'''MCP server for an agent that turns an article into a valuation.

People do not need this process. Notebooks still call ``company`` directly.

    python -m finagle.agent.mcp_server --transport stdio
    python -m finagle.agent.mcp_server --transport streamable-http

The SEC rejects anonymous requests. Set FINAGLE_SEC_USER_AGENT to a name and
an email address before calling get_10k_baseline or build_baseline. Install
the yahoo extra for the yfinance layer of build_baseline. Stdio is for an MCP client
on this machine. streamable-http is for a remote client; hosting it is separate.
'''

import argparse
import os
from typing import Optional

from finagle.agent.baseline import get_10k_baseline as load_10k_baseline
from finagle.agent.catalog import VALUE_FROM_ARTICLE, describe_inputs as catalog_inputs
from finagle.agent.runner import run_case as run_valuation
from finagle.agent.schema import case_tool_schema
from finagle.agent.sources import build_baseline as merge_baseline
from finagle.agent.validate import validate_case as check_case

_INSTRUCTIONS = (
    'Value a stock from an article already in the conversation, and from a filing '
    'when the user attached one. Call describe_inputs, then build_baseline with the '
    'article and attachment leaves to fill year 0 for the last fiscal year, then '
    'validate_case, then run_case. Use the value_from_article prompt. Do not invent '
    'method calls, and do not tune inputs to an article price target.'
)


def build_server():
    '''Return the MCP server with the valuation tools and the article prompt.'''
    server = _server_class()(
        'finagle-valuation',
        instructions=_INSTRUCTIONS,
    )

    @server.tool()
    def describe_inputs() -> dict:
        '''List every case field, the units, and how an article maps onto it.

        Call this before filling a case. Money is millions of the reporting
        currency, shares are millions of shares, and rates are decimals.
        '''
        return catalog_inputs()

    @server.tool()
    def get_10k_baseline(ticker: str, user_agent: str = '') -> dict:
        '''Year-0 figures from the latest 10-K, plus shares outstanding.

        Dollar amounts are millions. Debt and the other figures are year 0
        only, and the period is the latest 10-K fiscal year. user_agent can
        be omitted when FINAGLE_SEC_USER_AGENT is set to a name and an email.
        '''
        agent = (user_agent or '').strip() or os.environ.get('FINAGLE_SEC_USER_AGENT', '').strip()
        if not agent:
            return {
                'ok': False,
                'error': (
                    'Set FINAGLE_SEC_USER_AGENT to a name and an email address, '
                    'or pass user_agent.'
                ),
            }
        try:
            return load_10k_baseline(ticker, agent)
        except Exception as exc:
            return {'ok': False, 'error': '%s: %s' % (type(exc).__name__, exc)}

    @server.tool()
    def build_baseline(
        ticker: str,
        article: Optional[dict] = None,
        attachment: Optional[dict] = None,
        country: str = '',
        period_end: str = '',
        user_agent: str = '',
    ) -> dict:
        '''Fill year 0 from the article, then an attached filing, then the SEC, then yfinance.

        article and attachment map baseline fields (and optionally shares)
        to leaves you read, each with basis fiscal_year, period_end, and
        currency. Year 0 is the last completed fiscal year, never trailing
        twelve months. Canadian companies (a .TO, .V, .CN, or .NE ticker, or
        country CA) skip the SEC. Returns baseline leaves, shares, a price in
        the reporting currency, the source of each field, what was skipped,
        and warnings.
        '''
        agent = (user_agent or '').strip() or os.environ.get('FINAGLE_SEC_USER_AGENT', '').strip()
        try:
            return merge_baseline(
                ticker,
                article=article,
                attachment=attachment,
                user_agent=agent or None,
                country=country or None,
                period_end=period_end or None,
            )
        except Exception as exc:
            return {'ok': False, 'error': '%s: %s' % (type(exc).__name__, exc)}

    @server.tool()
    def validate_case(case: dict) -> dict:
        '''Check a case. Fix every error and missing field before run_case.

        Warnings do not stop a run. Each figure must be a leaf with value,
        source (article, attachment, 10k, yahoo, or assumption), and evidence.
        '''
        return check_case(case)

    @server.tool()
    def run_case(case: dict) -> dict:
        '''Compile a case and run value() on the existing company class.

        Returns value per share, the dividend-discount value per share, the
        forecast, log errors, and notebook code that reruns the same case.
        ok is false when validation fails or the model logs an error. The
        article price target is not an input.
        '''
        return run_valuation(case)

    @server.prompt(
        name='value_from_article',
        description='Turn a stock article into a finagle valuation case and run it.',
    )
    def value_from_article() -> str:
        '''Workflow from an article, an optional filing, the SEC, and yfinance to a value per share.'''
        return VALUE_FROM_ARTICLE

    _publish_case_schema(server)
    return server


def main(argv=None):
    parser = argparse.ArgumentParser(description='Serve the finagle valuation tools over MCP.')
    parser.add_argument(
        '--transport',
        default='stdio',
        choices=('stdio', 'streamable-http'),
        help='stdio for a local MCP client, streamable-http for a remote one',
    )
    args = parser.parse_args(argv)
    build_server().run(transport=args.transport)


def _server_class():
    try:
        from mcp.server.fastmcp import FastMCP
    except ModuleNotFoundError:
        from mcp.server.mcpserver import MCPServer as FastMCP
    return FastMCP


def _publish_case_schema(server):
    '''Show validate_case and run_case the case schema, not a free-form object.'''
    manager = getattr(server, '_tool_manager', None)
    if manager is None:
        return
    schema = case_tool_schema()
    for name in ('validate_case', 'run_case'):
        tool = manager.get_tool(name)
        if tool is not None and hasattr(tool, 'parameters'):
            tool.parameters = schema


if __name__ == '__main__':
    main()

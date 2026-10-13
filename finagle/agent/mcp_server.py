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
from finagle.agent.documents import read_pdf as read_pdf_file
from finagle.agent.runner import run_case as run_valuation
from finagle.agent.schema import case_tool_schema
from finagle.agent.sources import build_baseline as merge_baseline
from finagle.agent.update import update_case as revise_case
from finagle.agent.validate import validate_case as check_case

_INSTRUCTIONS = (
    'Value a stock from an article already in the conversation, and from other '
    'files when the user attached them. Read charts in the article and in '
    'attachments, including an investor presentation and a presentation URL the '
    'user supplies. Printed labels, OCR digits, and axis readings are article or '
    'attachment figures. Do not fetch the article URL. A figure stated in the '
    'request comes first. Call describe_inputs, then read_pdf on every PDF, '
    'then build_baseline with the prompt, article, attachments, and documents '
    'in that order. If unread_charts is set, cite those pages and call '
    'build_baseline again before validate_case. build_baseline fills the last fiscal '
    'year and the default assumptions. Then validate_case, then run_case. On a '
    'follow-up, call update_case with only what changed. Use the value_from_article '
    'prompt. Do not invent method calls, and do not tune inputs to an article '
    'price target.'
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
    def read_pdf(path: str, out_dir: str, name: str = '') -> dict:
        '''Extract a PDF, including a chart block for every large image.

        Call this on the article and each attachment PDF before build_baseline.
        name is "article" for the primary article, or the attachment name.
        The extract in out_dir holds page text plus OCR, or an UNREAD CHART
        line naming the png to open. Pass the return value in documents.
        '''
        try:
            return read_pdf_file(path, out_dir, name=name or None)
        except Exception as exc:
            return {'ok': False, 'error': '%s: %s' % (type(exc).__name__, exc)}

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
        prompt: Optional[dict] = None,
        article: Optional[dict] = None,
        attachment: Optional[dict] = None,
        attachments: Optional[list] = None,
        country: str = '',
        period_end: str = '',
        user_agent: str = '',
        documents: Optional[list] = None,
    ) -> dict:
        '''Fill the case from the prompt, the article, other attachments, the SEC, then yfinance.

        prompt, article, and each attachment map fields to leaves you read.
        Reported baseline leaves need basis fiscal_year, period_end, and
        currency. attachments is an ordered list of {name, leaves}; the
        request sets that order. documents is the list read_pdf returned.
        A chart page that is not cited comes back as unread_charts and no
        case. Year 0 is the last completed fiscal year,
        never trailing twelve months. Canadian companies (a .TO, .V, .CN, or
        .NE ticker, or country CA) skip the SEC. Also fills me, mc, the cost
        of equity, terminal growth, ROIC, the last close, year-0 leverage,
        and a proportional buyback price when those were not stated.
        '''
        agent = (user_agent or '').strip() or os.environ.get('FINAGLE_SEC_USER_AGENT', '').strip()
        try:
            return merge_baseline(
                ticker,
                prompt=prompt,
                article=article,
                attachment=attachment,
                attachments=attachments,
                user_agent=agent or None,
                country=country or None,
                period_end=period_end or None,
                documents=documents,
            )
        except Exception as exc:
            return {'ok': False, 'error': '%s: %s' % (type(exc).__name__, exc)}

    @server.tool()
    def update_case(case: dict, updates: dict) -> dict:
        '''Change a finished case and rerun it.

        updates is a partial case of leaves, source prompt. Derived
        assumptions whose inputs changed are recomputed from the numbers
        already in the case. The article, the attachments, the SEC, and
        yfinance are not read again.
        '''
        try:
            return revise_case(case, updates)
        except Exception as exc:
            return {'ok': False, 'error': '%s: %s' % (type(exc).__name__, exc)}

    @server.tool()
    def validate_case(case: dict) -> dict:
        '''Check a case. Fix every error and missing field before run_case.

        Warnings do not stop a run. Each figure must be a leaf with value,
        source (prompt, article, attachment, 10k, yahoo, or assumption), and evidence.
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

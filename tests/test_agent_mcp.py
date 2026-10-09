import pytest

pytest.importorskip('mcp')

import anyio  # noqa: E402
from finagle.agent.mcp_server import build_server  # noqa: E402


def test_server_exposes_the_case_schema_and_the_article_prompt():
    server = build_server()
    tools = anyio.run(server.list_tools)
    by_name = {tool.name: tool for tool in tools}
    assert set(by_name) == {
        'describe_inputs', 'get_10k_baseline', 'validate_case', 'run_case',
    }
    schema = by_name['validate_case'].input_schema
    assert schema['required'] == ['case']
    assert 'ticker' in schema['properties']['case']['properties']
    assert 'leaf' in schema['$defs']
    assert by_name['run_case'].input_schema == schema

    prompts = anyio.run(server.list_prompts)
    assert [prompt.name for prompt in prompts] == ['value_from_article']


def test_describe_inputs_and_the_prompt_are_callable():
    server = build_server()
    result = anyio.run(server.call_tool, 'describe_inputs', {})
    text = _text(result)
    assert 'millions of dollars' in text
    assert 'price_target' in text

    rendered = anyio.run(server.get_prompt, 'value_from_article')
    prompt_text = _text(rendered)
    assert 'get_10k_baseline' in prompt_text
    assert 'price target' in prompt_text.lower()


def test_missing_user_agent_is_an_error_result():
    server = build_server()
    result = anyio.run(server.call_tool, 'get_10k_baseline', {'ticker': 'ATKR'})
    text = _text(result)
    assert 'FINAGLE_SEC_USER_AGENT' in text


def _text(result):
    if isinstance(result, str):
        return result
    content = getattr(result, 'content', None)
    if content:
        parts = []
        for block in content:
            parts.append(getattr(block, 'text', '') or '')
        return '\n'.join(parts)
    messages = getattr(result, 'messages', None)
    if messages:
        parts = []
        for message in messages:
            block = getattr(message, 'content', message)
            parts.append(getattr(block, 'text', '') or str(block))
        return '\n'.join(parts)
    return str(result)

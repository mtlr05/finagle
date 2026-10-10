'''Tools that let an agent run a valuation without changing company or the notebooks.

``import finagle`` does not import this package. People keep constructing
``company`` and calling its methods directly.
'''

from finagle.agent.baseline import get_10k_baseline, sec_baseline, shares_from_facts
from finagle.agent.catalog import VALUE_FROM_ARTICLE, describe_inputs
from finagle.agent.compile import compile_case
from finagle.agent.runner import run_case
from finagle.agent.schema import CASE_JSON_SCHEMA, case_tool_schema
from finagle.agent.sources import build_baseline
from finagle.agent.validate import validate_case
from finagle.agent.yahoo import yahoo_baseline

__all__ = [
    'CASE_JSON_SCHEMA',
    'VALUE_FROM_ARTICLE',
    'build_baseline',
    'case_tool_schema',
    'compile_case',
    'describe_inputs',
    'get_10k_baseline',
    'run_case',
    'sec_baseline',
    'shares_from_facts',
    'validate_case',
    'yahoo_baseline',
]

'''Prompt-only HROW valuations.

Each test's user input is one sentence and no article. The first sentence is
a 30 percent sales-growth forecast. Each later sentence is filled with the
figure taken from the previous run. The four valuations match.
'''

import os
import re

import pytest

pytest.importorskip('yfinance')

from finagle.agent import build_baseline, run_case  # noqa: E402
from finagle.agent.schema import leaf  # noqa: E402

pytestmark = pytest.mark.skipif(
    not os.environ.get('FINAGLE_SEC_USER_AGENT', '').strip(),
    reason='FINAGLE_SEC_USER_AGENT is not set',
)

_CACHE = {}
_PROMPT = re.compile(
    r"^run in bot mode, ticker: HROW, this year's "
    r"(?P<kind>sales growth forecast|sales forecast is|EBITDA forecast is|"
    r"EBITDA growth forecast is) (?P<number>.+)$"
)


def test_sales_growth_forecast():
    result = _sales_growth_run()
    revenue0 = result['years'][0]['revenue']
    assert result['years'][1]['revenue'] == pytest.approx(revenue0 * 1.30)
    assert result['case']['forecast']['sales_growth']['value'] == pytest.approx([0.30])


def test_sales_forecast():
    first = _sales_growth_run()
    result = _sales_run()
    assert result['case']['forecast']['sales_growth']['value'] == pytest.approx([0.30])
    _assert_same_value(result, first)
    assert result['years'][1]['ebitda'] == pytest.approx(first['years'][1]['ebitda'])


def test_ebitda_forecast():
    sales = _sales_run()
    result = _ebitda_run()
    assert result['case']['forecast']['ebitda_growth']['value'] == pytest.approx(
        sales['case']['forecast']['ebitda_growth']['value']
    )
    assert result['years'][1]['ebitda'] == pytest.approx(sales['years'][1]['ebitda'])
    _assert_same_value(result, sales)


def test_ebitda_growth_forecast():
    ebitda = _ebitda_run()
    result = _ebitda_growth_run()
    assert result['years'][1]['ebitda'] == pytest.approx(ebitda['years'][1]['ebitda'])
    _assert_same_value(result, ebitda)
    _assert_same_value(result, _sales_growth_run())


def _sales_growth_run():
    return _cached(
        'sales_growth',
        "run in bot mode, ticker: HROW, this year's sales growth forecast 30%",
    )


def _sales_run():
    sales = _sales_growth_run()['years'][1]['revenue']
    return _cached(
        'sales',
        "run in bot mode, ticker: HROW, this year's sales forecast is %r" % float(sales),
    )


def _ebitda_run():
    ebitda = _sales_run()['years'][1]['ebitda']
    return _cached(
        'ebitda',
        "run in bot mode, ticker: HROW, this year's EBITDA forecast is %r" % float(ebitda),
    )


def _ebitda_growth_run():
    stated = _ebitda_run()
    year0 = stated['case']['baseline']['ebitda']['value']
    growth = stated['years'][1]['ebitda'] / float(year0) - 1
    return _cached(
        'ebitda_growth',
        "run in bot mode, ticker: HROW, this year's EBITDA growth forecast is %r" % float(growth),
    )


def _cached(key, sentence):
    if key not in _CACHE:
        _CACHE[key] = _run(sentence)
    return _CACHE[key]


def _run(sentence):
    kind, value = _parse(sentence)
    if kind in ('sales', 'ebitda'):
        peeked = _baseline()
        base = _number(peeked['revenue'] if kind == 'sales' else peeked['ebitda'])
        value = value / base - 1
        kind = 'sales_growth' if kind == 'sales' else 'ebitda_growth'
    prompt = {'forecast': {kind: leaf([value], 'prompt', sentence)}}
    built = build_baseline(
        'HROW', prompt=prompt, user_agent=os.environ['FINAGLE_SEC_USER_AGENT'],
    )
    assert built['ok'], built.get('error')
    case = built['case']
    _fill(case)
    result = run_case(case)
    assert result['ok'], result.get('validation') or result.get('log')
    return result


def _baseline():
    if 'baseline' not in _CACHE:
        built = build_baseline('HROW', user_agent=os.environ['FINAGLE_SEC_USER_AGENT'])
        assert built['ok'], built.get('error')
        _CACHE['baseline'] = built['baseline']
    return _CACHE['baseline']


def _fill(case):
    '''Assumptions the bot supplies when the prompt does not.'''
    case['year'] = 10
    baseline = case.setdefault('baseline', {})
    forecast = case.setdefault('forecast', {})
    rates = case.setdefault('rates', {})
    if 'nol' not in baseline:
        baseline['nol'] = leaf(0, 'assumption', 'No net operating loss was stated.')
    if 'noa' not in baseline:
        baseline['noa'] = leaf(0, 'assumption', 'No non-operating assets were stated.')
    if 't' not in rates:
        rates['t'] = leaf(0.21, 'assumption', 'US federal statutory tax rate.')
    if 'rd' not in rates:
        debt = _number(baseline['debt'])
        if not debt:
            raise AssertionError('year-0 debt is zero, so interest/debt is not a cost of debt')
        rates['rd'] = leaf(
            _number(baseline['interest']) / debt, 'assumption',
            'Year-0 interest divided by year-0 debt.',
        )
    forecast.setdefault('ebitda_basis', leaf(
        'organic', 'assumption', 'The prompt states no future deals.',
    ))
    forecast.setdefault('acquisition_disclosure', {
        'status': leaf('none', 'assumption', 'The prompt states no future deals.'),
    })
    forecast.setdefault('dwc', {
        'mode': leaf('zero', 'assumption', 'No working-capital change was stated.'),
    })
    case.setdefault('distribution', {
        'mode': leaf('none', 'assumption', 'No distribution was stated.'),
    })


def _parse(sentence):
    match = _PROMPT.match(sentence.strip())
    if match is None:
        raise AssertionError('unrecognized prompt: %s' % sentence)
    kind = match.group('kind')
    raw = match.group('number').strip()
    if kind == 'sales growth forecast':
        if not raw.endswith('%'):
            raise AssertionError(sentence)
        return 'sales_growth', float(raw[:-1]) / 100.0
    value = float(raw)
    if kind == 'sales forecast is':
        return 'sales', value
    if kind == 'EBITDA forecast is':
        return 'ebitda', value
    return 'ebitda_growth', value


def _number(node):
    value = node['value'] if isinstance(node, dict) else node
    if isinstance(value, list):
        value = value[0]
    return float(value)


def _assert_same_value(result, other):
    assert result['value_per_share'] == pytest.approx(other['value_per_share'])
    assert result['value_per_share_DDM'] == pytest.approx(other['value_per_share_DDM'])

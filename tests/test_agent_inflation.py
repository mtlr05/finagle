'''Expected inflation from FRED and the IMF, with no network.'''

import pytest

from finagle.agent.defaults import apply_defaults
from finagle.agent.inflation import expected_inflation
from finagle.agent.schema import leaf


FRED = """observation_date,T10YIE
2026-11-13,2.28
2026-11-14,2.31
2026-11-17,.
"""

IMF = """{
  "values": {"PCPIPCH": {"CAN": {"2028": 1.9, "2030": 2.1, "2029": null}}}
}"""


def test_fred_uses_the_last_observation_and_divides_by_100():
    quote = expected_inflation('USD', fetch=lambda url: FRED)
    assert 'T10YIE' in quote['leaf']['evidence']
    assert quote['ok']
    assert quote['value'] == pytest.approx(0.0231)
    assert quote['leaf']['source'] == 'fred'
    assert quote['leaf']['as_of'] == '2026-11-14'


def test_imf_uses_the_furthest_projection_year():
    quote = expected_inflation('CAD', fetch=lambda url: IMF)
    assert quote['ok']
    assert 'CAN' in quote['leaf']['evidence']
    assert '2030' in quote['leaf']['evidence']
    assert quote['value'] == pytest.approx(0.021)
    assert quote['leaf']['source'] == 'imf'


def test_an_unknown_currency_is_left_unset():
    quote = expected_inflation('XYZ', fetch=lambda url: FRED)
    assert not quote['ok']
    assert 'XYZ' in quote['error']


def test_a_fetch_exception_is_an_error():
    def broken(url):
        raise TimeoutError('timed out')

    quote = expected_inflation('USD', fetch=broken)
    assert not quote['ok']
    assert 'TimeoutError' in quote['error']


def test_defaults_fill_lam_and_inflation_from_the_client():
    case = {'baseline': {}, 'rates': {}}

    def fetch(url):
        assert 'T10YIE' in url
        return FRED

    filled = apply_defaults(case, currency='USD', fetch_yield=True, inflation_client=fetch)
    assert case['rates']['inflation']['value'] == pytest.approx(0.0231)
    assert case['rates']['inflation']['source'] == 'fred'
    assert case['rates']['lam']['value'] == 0.25
    assert case['rates']['lam']['source'] == 'assumption'
    assert 'rates.inflation' in filled['filled']
    assert 'rates.lam' in filled['filled']
    assert filled['warnings'] == []


def test_inflation_falls_back_to_two_percent_when_the_client_fails():
    case = {'rates': {}}

    def broken(url):
        raise OSError('down')

    filled = apply_defaults(case, currency='USD', fetch_yield=True, inflation_client=broken)
    assert case['rates']['inflation']['value'] == 0.02
    assert case['rates']['inflation']['source'] == 'assumption'
    assert '2 percent' in case['rates']['inflation']['evidence']
    assert any('could not be loaded' in warning for warning in filled['warnings'])


def test_validation_rejects_a_bad_pass_through_and_warns_when_growth_is_below_inflation():
    from finagle.agent.validate import validate_case

    high = validate_case({'rates': {'lam': leaf(1.2, 'prompt', 'too high')}})
    assert any('from 0 to 1' in error for error in high['errors'])
    text = validate_case({'rates': {'inflation': leaf('high', 'prompt', 'not a number')}})
    assert any('rates.inflation must be a number' in error for error in text['errors'])
    undefined = validate_case({'rates': {
        'roict': leaf(0.02, 'assumption', 'roic'),
        'lam': leaf(1, 'assumption', 'full pass-through'),
        'inflation': leaf(0.03, 'fred', 'breakeven'),
        'gt': leaf(0.02, 'assumption', 'growth'),
    }})
    assert any('reinvestment rate is undefined' in error for error in undefined['errors'])
    low_growth = validate_case({'rates': {
        'roict': leaf(0.15, 'assumption', 'roic'),
        'lam': leaf(1, 'assumption', 'full pass-through'),
        'inflation': leaf(0.03, 'fred', 'breakeven'),
        'gt': leaf(0.02, 'assumption', 'growth'),
    }})
    assert any('terminal net investment is negative' in warning for warning in low_growth['warnings'])
    assert not any('reinvestment' in error for error in low_growth['errors'])


def test_a_stated_inflation_is_kept_and_an_update_does_not_refetch():
    from finagle.agent.update import update_case

    case = {
        'ticker': 'X',
        'year': 1,
        'rates': {'inflation': leaf(0.031, 'prompt', 'stated inflation')},
    }
    calls = []

    def fetch(url):
        calls.append(url)
        return FRED

    apply_defaults(case, currency='USD', fetch_yield=False, inflation_client=fetch)
    assert case['rates']['inflation']['value'] == 0.031
    assert calls == []
    updated = update_case(case, {
        'rates': {'gt': leaf(0.02, 'prompt', 'terminal growth')},
    }, run=False)
    assert updated['case']['rates']['inflation']['value'] == 0.031
    assert calls == []

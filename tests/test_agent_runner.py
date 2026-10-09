import os

import pandas as pd
import pytest

from finagle.agent.runner import run_case, sanity_flags


def leaf(value):
    return {
        'value': value,
        'source': 'assumption',
        'evidence': 'replay of an existing notebook case',
    }


def _rates(rd, re, t, gt, roict):
    return {
        'rd': leaf(rd),
        're': leaf(re),
        't': leaf(t),
        'gt': leaf(gt),
        'roict': leaf(roict),
    }


def _baseline(date, ebitda, capex, da, tax, interest, sbc, debt, cash, nol, noa):
    return {
        'date': leaf(date),
        'ebitda': leaf(ebitda),
        'ebitda_definition': leaf('before_sbc'),
        'capex': leaf(capex),
        'da': leaf(da),
        'tax': leaf(tax),
        'interest': leaf(interest),
        'sbc': leaf(sbc),
        'debt': leaf(debt),
        'cash': leaf(cash),
        'nol': leaf(nol),
        'noa': leaf(noa),
    }


def disck_case():
    ebitda = [13.7, 13.8, 14.0, 14.7, 16.0, 18.1, 19.6, 21.4, 23.7, 26.4, 29.7]
    capex = [1.4, 1.5, 1.6, 1.6, 1.8, 2.0, 2.0, 2.1, 2.2, 2.4, 2.5]
    debt = [51.7, 43.3, 34.6, 36.6, 40.0, 45.2, 48.9, 53.6, 59.3, 74.3, 80]
    return {
        'path': 'ebitda',
        'ticker': 'DISCK',
        'year': 10,
        'market': {'shares': leaf(2.3)},
        'rates': _rates(0.065, 0.10, 0.21, 0.02, 0.75),
        'baseline': _baseline(
            '2021-12-31', 13.7, 1.4, [6.8, 7.0, 7.3, 8.0, 9.0],
            0.68, 3.6, 0, 51.7, 0, 0, 0,
        ),
        'forecast': {
            'ebitda': leaf(ebitda),
            'capex': leaf(capex),
            'sbc': leaf([0] * 11),
            'dwc': {'mode': leaf('zero')},
        },
        'debt': {
            'mode': leaf('target'),
            'leverage': leaf(2.5),
            'start_year': leaf(1),
            'values': leaf(debt),
        },
        'distribution': {
            'mode': leaf('max_buybacks'),
            'price': leaf(28.22),
            'price_path': leaf('proportional'),
        },
    }


def frg_acquire_case():
    capex = [40, 46.0, 50.6, 55.7, 61.2, 67.3, 73.2, 78.6, 83.3, 87.2, 90.1]
    return {
        'path': 'ebitda',
        'ticker': 'FRG',
        'year': 10,
        'market': {'shares': leaf(0.744 + 40.295)},
        'rates': _rates(0.065, 0.10, 0.21, 0.02, 0.15),
        'baseline': _baseline(
            '2021-12-26', 330, 40, [51.847, 54.4],
            0, 70, 0, 1072.9, 159.72, 127.4, 55.86,
        ),
        'forecast': {
            'ebitda_growth': leaf([0.15, 0.1, 0.1, 0.1, 0.1]),
            'capex': leaf(capex),
            'sbc': leaf([0] * 11),
            'dwc': {'mode': leaf('zero')},
        },
        'debt': {
            'mode': leaf('target'),
            'leverage': leaf(2.5),
            'start_year': leaf(1),
        },
        'acquisitions': [{
            'year': leaf(0),
            'ebitda_frac': leaf(0.2696),
            'multiple': leaf(6.52),
            'leverage': leaf(6.52),
            'next_growth': leaf(0.1),
            'capex_frac': leaf(0.22),
            'pay_from_cash': leaf(True),
        }],
        'distribution': {
            'mode': leaf('max_buybacks'),
            'price': leaf(45),
            'price_path': leaf('proportional'),
        },
    }


def pl_allocate_case():
    ebitda = [-64, -35, 42, 162, 242, 354, 502, 694, 933, 1219, 1548, 1907, 2278, 2637, 2956, 3205]
    capex = [26, 38, 50, 61, 87, 120, 161, 210, 268, 332, 400, 468, 531, 584, 622, 641]
    return {
        'path': 'ebitda',
        'ticker': 'PL',
        'year': 15,
        'market': {'shares': leaf(270 + 12.8 + 21.459 + 8)},
        'rates': _rates(0.075, 0.10, 0.21, 0.02, 0.15),
        'baseline': _baseline(
            '2023-01-31', -64, 26, 60,
            0, 0, 80, 0, 598 - 45, 114.6, 0,
        ),
        'forecast': {
            'ebitda': leaf(ebitda),
            'capex': leaf(capex),
            'sbc': leaf([80, 80, 80, 80, 80]),
            'sbc_rate_terminal': leaf(0.17),
            'dwc': {'mode': leaf('zero')},
        },
        'debt': {
            'mode': leaf('target'),
            'leverage': leaf(1),
            'start_year': leaf(4),
        },
        'distribution': {
            'mode': leaf('buyback_schedule'),
            'schedule': leaf([0]),
            'price': leaf(10),
            'price_path': leaf('constant'),
        },
    }


def _assert_matches(result, pickle_name):
    assert result['ok'], {
        'validation': result['validation'],
        'error': result.get('error'),
        'log': result['log'],
        'flags': result['sanity_flags'],
    }
    answer = pd.read_pickle(os.path.join(os.path.dirname(__file__), pickle_name))
    equity = [row['equity'] for row in result['years']]
    firm = [row['firm'] for row in result['years']]
    assert equity == pytest.approx(list(answer.loc['equity']), rel=1e-9, abs=1e-9)
    assert firm == pytest.approx(list(answer.loc['firm']), rel=1e-9, abs=1e-9)


def test_disck_matches_the_ebitda_snapshot():
    result = run_case(disck_case())
    _assert_matches(result, 'fcf_from_ebitda.pkl')
    assert [step['method'] for step in result['steps']] == [
        'fcf_from_ebitda', 'fcf_to_debt', 'fcf_to_buyback', 'value',
    ]
    assert 'import finagle as cmp' in result['notebook_code']
    assert result['value_per_share'] == pytest.approx(result['equity'] / 2.3)


def test_frg_acquisition_matches_the_snapshot():
    result = run_case(frg_acquire_case())
    _assert_matches(result, 'fcf_to_acquire.pkl')
    methods = [step['method'] for step in result['steps']]
    assert methods == [
        'forecast_ebitda', 'load_financials', 'fcf_from_ebitda',
        'fcf_to_acquire', 'fcf_to_debt', 'fcf_to_buyback', 'value',
    ]
    acquire = next(step for step in result['steps'] if step['method'] == 'fcf_to_acquire')
    assert acquire['args']['ebitda_frac'] == pytest.approx(0.2696)
    assert 'ebitda_frac=' in result['notebook_code']


def test_pl_allocate_matches_the_snapshot():
    result = run_case(pl_allocate_case())
    _assert_matches(result, 'fcf_to_allocate.pkl')
    assert [step['method'] for step in result['steps']] == [
        'forecast_sbc', 'load_financials', 'fcf_from_ebitda',
        'fcf_to_debt', 'fcf_to_allocate', 'value',
    ]
    allocate = next(step for step in result['steps'] if step['method'] == 'fcf_to_allocate')
    assert allocate['args']['buybacks'] == [0]
    assert allocate['args']['dp'] == 'constant'


def test_run_does_not_write_the_log_into_the_working_directory(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    result = run_case(disck_case())
    assert result['ok']
    assert list(tmp_path.iterdir()) == []


def test_acquired_ebitda_becomes_a_fraction_of_that_years_ebitda():
    case = {
        'path': 'ebitda',
        'ticker': 'DEAL',
        'year': 3,
        'market': {'shares': leaf(10)},
        'rates': _rates(0.05, 0.10, 0.21, 0.02, 0.30),
        'baseline': _baseline(
            '2021-12-31', 10, 1, 1, 1, 0.5, 0, 5, 10, 0, 0,
        ),
        'forecast': {
            'ebitda': leaf([10, 11, 12, 13]),
            'capex': leaf([1, 1, 1, 1]),
            'sbc': leaf([0, 0, 0, 0]),
            'dwc': {'mode': leaf('zero')},
        },
        'debt': {'mode': leaf('hold')},
        'acquisitions': [{
            'year': leaf(1),
            'acquired_ebitda': leaf(2.2),
            'multiple': leaf(8),
            'leverage': leaf(0),
            'next_growth': leaf(0.1),
            'capex_frac': leaf(0.1),
            'pay_from_cash': leaf(False),
        }],
        'distribution': {'mode': leaf('none')},
    }
    result = run_case(case)
    assert result['ok'], {'error': result.get('error'), 'log': result['log']}
    acquire = next(step for step in result['steps'] if step['method'] == 'fcf_to_acquire')
    assert acquire['args']['ebitda_frac'] == pytest.approx(2.2 / 11)
    assert 'ebitda_frac=' in result['notebook_code']


def test_deal_value_uses_the_multiple_before_the_fraction():
    case = {
        'path': 'ebitda',
        'ticker': 'DEAL2',
        'year': 3,
        'market': {'shares': leaf(10)},
        'rates': _rates(0.05, 0.10, 0.21, 0.02, 0.30),
        'baseline': _baseline(
            '2021-12-31', 10, 1, 1, 1, 0.5, 0, 5, 10, 0, 0,
        ),
        'forecast': {
            'ebitda': leaf([10, 11, 12, 13]),
            'capex': leaf([1, 1, 1, 1]),
            'sbc': leaf([0, 0, 0, 0]),
            'dwc': {'mode': leaf('zero')},
        },
        'debt': {'mode': leaf('hold')},
        'acquisitions': [{
            'year': leaf(1),
            'deal_value': leaf(22),
            'multiple': leaf(10),
            'leverage': leaf(0),
            'next_growth': leaf(0.1),
            'capex_frac': leaf(0.1),
            'pay_from_cash': leaf(False),
        }],
        'distribution': {'mode': leaf('none')},
    }
    result = run_case(case)
    assert result['ok'], {'error': result.get('error'), 'log': result['log']}
    acquire = next(step for step in result['steps'] if step['method'] == 'fcf_to_acquire')
    assert acquire['args']['ebitda_frac'] == pytest.approx(2.2 / 11)


def test_direct_fcfe_matches_the_snapshot():
    case = {
        'path': 'fcfe',
        'ticker': 'abc',
        'year': 6,
        'rates': {'re': leaf(0.0557 + 0.0214), 'gt': leaf(0.0214)},
        'baseline': {'date': leaf('2021-12-31')},
        'fcfe': leaf([0, 155.76, 161.20, 166.84, 172.67, 178.71, 182.53]),
    }
    result = run_case(case)
    assert result['ok'], {'error': result.get('error'), 'log': result['log'], 'validation': result['validation']}
    answer = pd.read_pickle(os.path.join(os.path.dirname(__file__), 'value.pkl'))
    equity = [row['equity'] for row in result['years']]
    assert equity == pytest.approx(list(answer.loc['equity']), rel=1e-9, abs=1e-9)
    assert result['value_per_share'] is None
    assert result['firm'] is None


def test_validation_failure_does_not_run():
    result = run_case({'ticker': 'NOPE'})
    assert result['ok'] is False
    assert result['value_per_share'] is None
    assert result['validation']['missing']
    assert result['notebook_code'] == ''


def test_notebook_code_reruns_the_resolved_acquisition(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    result = run_case(frg_acquire_case())
    assert result['ok']
    namespace = {}
    exec(result['notebook_code'], namespace)
    assert namespace['model'].fin['equity'].iloc[0] == pytest.approx(result['equity'])


def test_earnings_path_reports_equity_only():
    case = {
        'path': 'earnings',
        'ticker': 'xyz',
        'year': 5,
        'market': {'shares': leaf(1)},
        'rates': {'re': leaf(0.12), 'gt': leaf(0.01)},
        'baseline': {'date': leaf('2021-12-31')},
        'earnings': {
            'e': leaf(1),
            'payout': leaf(0.5),
            'gf': leaf([0.05, 0.04]),
            'roe': leaf(0.15),
        },
    }
    result = run_case(case)
    assert result['ok'], {
        'error': result.get('error'),
        'log': result['log'],
        'validation': result['validation'],
    }
    assert result['value_per_share'] is None
    assert result['firm'] is None
    answer = pd.read_pickle(os.path.join(os.path.dirname(__file__), 'fcf_from_earnings.pkl'))
    equity = [row['equity'] for row in result['years']]
    assert equity == pytest.approx(list(answer.loc['equity']), rel=1e-9, abs=1e-9)


def test_sanity_flags_follow_the_cash_rules():
    assert sanity_flags(0.1, 0.02, [1, -0.1], [0, -5, 1], 0.08) == ['negative cash']
    assert sanity_flags(0.1, 0.02, [1], [0, -5, -1], 0.08) == ['negative FCFE after year 1']
    assert 're <= gt' in sanity_flags(0.02, 0.02, [], [], 0.03)
    assert 'wacc <= gt' in sanity_flags(0.1, 0.02, [], [], 0.02)
    assert sanity_flags(0.1, 0.02, [0], [1, 1, 1], 0.08) == []

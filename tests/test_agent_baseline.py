import pandas as pd

from finagle.agent.baseline import (
    get_10k_baseline,
    leaves_from_baseline,
    sec_baseline,
    shares_from_facts,
)
from finagle.sec import Baseline, baseline_from_facts
from tests.test_sec import PERIOD, _clean


def test_leaves_unwrap_year_zero_lists_and_keep_the_tag():
    baseline = baseline_from_facts(_clean(), PERIOD)
    leaves = leaves_from_baseline(baseline)

    assert leaves['ebitda']['value'] == 25.0
    assert leaves['ebitda']['source'] == '10k'
    assert 'OperatingIncomeLoss' in leaves['ebitda']['evidence']
    assert leaves['ebitda']['period_end'] == PERIOD
    assert leaves['cash']['value'] == 10.0
    assert leaves['date']['value'] == PERIOD
    assert isinstance(leaves['debt']['value'], float)


def test_shares_prefer_the_period_end_total():
    period = '2023-09-30'
    facts = _share_facts([
        _share(46_000_000, end='2024-02-15', filed='2024-02-20'),
        _share(40_000_000, end=period, filed='2023-11-01'),
        _share(5_000_000, end=period, filed='2023-11-01', segment={'dimension': 'class'}),
    ])
    leaf = shares_from_facts(facts, period)
    assert leaf['value'] == 40.0
    assert leaf['evidence'] == 'dei:EntityCommonStockSharesOutstanding'
    assert 'class' not in leaf.get('note', '')


def test_shares_use_the_cover_date_when_the_period_is_missing():
    facts = _share_facts([
        _share(46_000_000, end='2024-02-15', filed='2024-02-20'),
    ])
    leaf = shares_from_facts(facts, '2023-12-31')
    assert leaf['value'] == 46.0
    assert 'cover date' in leaf['note']
    assert leaf['period_end'] == '2024-02-15'


def test_dimensioned_share_count_is_marked_as_one_class():
    facts = _share_facts([
        _share(12_000_000, end=PERIOD, segment={'dimension': 'class'}),
    ])
    leaf = shares_from_facts(facts, PERIOD)
    assert leaf['value'] == 12.0
    assert 'one share class' in leaf['note']


def test_missing_share_tag_is_a_null_leaf():
    leaf = shares_from_facts({'facts': {'dei': {}}}, PERIOD)
    assert leaf['value'] is None
    assert leaf['source'] == '10k'


def test_get_10k_baseline_maps_the_helper_and_the_share_count(monkeypatch):
    baseline = baseline_from_facts(_clean(), PERIOD)
    facts = _share_facts([_share(46_000_000, end=PERIOD)])

    monkeypatch.setattr('finagle.agent.baseline.last_10k', lambda *args, **kwargs: baseline)
    monkeypatch.setattr('finagle.agent.baseline._cik_for_ticker', lambda *args, **kwargs: '0000000001')
    monkeypatch.setattr('finagle.agent.baseline._get_json', lambda *args, **kwargs: facts)

    result = get_10k_baseline('ATKR', 'Name you@example.com')
    assert result['baseline']['ebitda']['value'] == 25.0
    assert result['baseline']['debt']['value'] == 45.0
    assert result['shares']['value'] == 46.0
    assert 'dwc' in result['not_in_10k']
    assert '10-K' in result['note']


def test_sec_leaves_are_fiscal_year_dollars():
    leaves = leaves_from_baseline(baseline_from_facts(_clean(), PERIOD))
    assert leaves['ebitda']['basis'] == 'fiscal_year'
    assert leaves['ebitda']['currency'] == 'USD'
    assert leaves['date']['basis'] == 'fiscal_year'
    assert 'currency' not in leaves['date']


def _sec_feed(monkeypatch, report_dates):
    facts = _clean()
    facts['facts']['dei'] = _share_facts([_share(46_000_000, end=PERIOD)])['facts']['dei']
    submissions = {'filings': {'recent': {
        'form': ['10-Q'] + ['10-K'] * len(report_dates),
        'reportDate': ['2024-09-30'] + list(report_dates),
    }}}

    def fake_get_json(url, user_agent):
        if 'submissions' in url:
            return submissions
        if 'companyfacts' in url:
            return facts
        return {'0': {'ticker': 'ATKR', 'cik_str': 1}}

    monkeypatch.setattr('finagle.agent.baseline._get_json', fake_get_json)


def test_sec_baseline_reads_the_requested_fiscal_year(monkeypatch):
    _sec_feed(monkeypatch, ['2024-12-31', PERIOD])
    result = sec_baseline('ATKR', 'Name you@example.com', period_end=PERIOD)
    assert result['ok']
    assert result['period_end'] == PERIOD
    assert result['latest_period_end'] == '2024-12-31'
    assert result['baseline']['ebitda']['value'] == 25.0
    assert result['baseline']['date'] == {
        'value': PERIOD, 'source': '10k', 'evidence': '10-K report date',
        'period_end': PERIOD, 'basis': 'fiscal_year',
    }
    assert result['shares']['value'] == 46.0


def test_sec_baseline_without_that_10k_lists_the_years(monkeypatch):
    _sec_feed(monkeypatch, ['2024-12-31', PERIOD])
    result = sec_baseline('ATKR', 'Name you@example.com', period_end='2022-12-31')
    assert not result['ok']
    assert result['periods'] == ['2024-12-31', PERIOD]
    assert result['latest_period_end'] == '2024-12-31'


def test_missing_fact_stays_visible():
    provenance = pd.DataFrame([{
        'key': 'interest',
        'value': None,
        'tags': '',
        'raw dollars': None,
        'period end': PERIOD,
        'note': 'Interest expense was not found in the latest 10-K.',
    }])
    baseline = Baseline(provenance, {})
    leaves = leaves_from_baseline(baseline)
    assert leaves['interest']['value'] is None
    assert 'not found' in leaves['interest']['evidence']


def _share(val, end=PERIOD, filed='2024-02-20', segment=None):
    fact = {
        'end': end,
        'val': val,
        'form': '10-K',
        'fp': 'FY',
        'filed': filed,
        'accn': filed,
    }
    if segment is not None:
        fact['segment'] = segment
    return fact


def _share_facts(rows):
    return {
        'facts': {
            'dei': {
                'EntityCommonStockSharesOutstanding': {
                    'units': {'shares': rows},
                },
            },
        },
    }

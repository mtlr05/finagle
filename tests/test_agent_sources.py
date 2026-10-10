import copy
import json
import os
import sys

import pandas as pd
import pytest

from finagle.agent import build_baseline, validate_case, yahoo_baseline
from finagle.agent.schema import leaf

ROOT = os.path.dirname(os.path.dirname(__file__))
FY23 = '2023-12-31'
FY22 = '2022-12-31'
UA = 'Name you@example.com'

YAHOO_ROWS = {
    FY23: {
        'Total Revenue': 1300e6, 'Operating Income': 200e6, 'Tax Provision': 40e6,
        'Interest Expense': 25e6, 'Depreciation And Amortization': 45e6,
        'Stock Based Compensation': 15e6, 'Capital Expenditure': -50e6,
        'Total Debt': 420e6, 'Cash Cash Equivalents And Short Term Investments': 95e6,
    },
    FY22: {
        'Total Revenue': 1100e6, 'Operating Income': 150e6, 'Tax Provision': 30e6,
        'Interest Expense': 20e6, 'Depreciation And Amortization': 40e6,
        'Stock Based Compensation': 10e6, 'Capital Expenditure': -45e6,
        'Total Debt': 400e6, 'Cash Cash Equivalents And Short Term Investments': 80e6,
    },
}
_INCOME = ('Total Revenue', 'Operating Income', 'Tax Provision', 'Interest Expense')
_CASH = ('Depreciation And Amortization', 'Stock Based Compensation', 'Capital Expenditure')
_BALANCE = ('Total Debt', 'Cash Cash Equivalents And Short Term Investments')


def _table(rows_by_period, names):
    return pd.DataFrame({
        pd.Timestamp(period): {name: rows.get(name) for name in names if name in rows}
        for period, rows in rows_by_period.items()
    })


class FakeTicker:
    def __init__(self, rows_by_period, info):
        self.income_stmt = _table(rows_by_period, _INCOME)
        self.cashflow = _table(rows_by_period, _CASH)
        self.balance_sheet = _table(rows_by_period, _BALANCE)
        self.info = info


class FakeFx:
    def __init__(self, rate):
        self.rate = rate

    def history(self, period):
        return pd.DataFrame({'Close': [self.rate]}, index=[pd.Timestamp('2024-08-02')])


class FakeYahoo:
    def __init__(self, rows_by_period=None, info=None, fx=None):
        self.rows = YAHOO_ROWS if rows_by_period is None else rows_by_period
        self.info = info or {
            'sharesOutstanding': 52e6, 'currentPrice': 44.0,
            'currency': 'USD', 'financialCurrency': 'USD',
        }
        self.fx = fx
        self.calls = []

    def __call__(self, symbol):
        self.calls.append(symbol)
        if symbol.endswith('=X'):
            return FakeFx(self.fx)
        return FakeTicker(self.rows, self.info)


def _sec_leaves(period, **values):
    defaults = dict(revenue=1250.0, ebitda=255.0, capex=48.0, da=42.0, tax=39.0,
                    interest=None, sbc=13.0, debt=410.0, cash=92.0)
    defaults.update(values)
    leaves = {'date': leaf(period, '10k', '10-K report date', period_end=period, basis='fiscal_year')}
    for key, value in defaults.items():
        leaves[key] = leaf(value, '10k', 'us-gaap tag for ' + key, period_end=period,
                           basis='fiscal_year', currency='USD')
    return leaves


class FakeSec:
    def __init__(self, years=(FY23, FY22)):
        self.years = {year: _sec_leaves(year) for year in years}
        self.latest = max(years)
        self.calls = []

    def __call__(self, ticker, user_agent, period_end=None):
        self.calls.append((ticker, period_end))
        period = period_end or self.latest
        if period not in self.years:
            return {'ok': False, 'periods': sorted(self.years, reverse=True),
                    'latest_period_end': self.latest, 'error': 'no 10-K'}
        return {
            'ok': True, 'period_end': period, 'latest_period_end': self.latest,
            'baseline': copy.deepcopy(self.years[period]),
            'shares': leaf(49.0, '10k', 'dei:EntityCommonStockSharesOutstanding', period_end=period),
        }


@pytest.fixture
def sec(monkeypatch):
    fake = FakeSec()
    monkeypatch.setattr('finagle.agent.sources.sec_baseline', fake)
    return fake


def _article(period=FY23, currency='USD', **values):
    leaves = {}
    for key, value in values.items():
        leaves[key] = leaf(value, 'article', 'quoted ' + key, period_end=period,
                           basis='fiscal_year', currency=currency)
    return leaves


def test_each_field_comes_from_the_first_source_that_has_it(sec):
    article = _article(ebitda=240.0)
    article['ebitda_definition'] = leaf('before_sbc', 'article', 'adds back SBC')
    attachment = {
        'ebitda': leaf(999.0, 'attachment', 'p. 9: EBITDA', period_end=FY23,
                       basis='fiscal_year', currency='USD'),
        'capex': leaf(36.0, 'attachment', 'p. 47: Capital expenditure 36', period_end=FY23,
                      basis='fiscal_year', currency='USD'),
    }
    yahoo = FakeYahoo()
    result = build_baseline('XMPL', article=article, attachment=attachment,
                            user_agent=UA, yahoo_client=yahoo)

    assert result['ok']
    assert result['period_end'] == FY23
    assert result['period_source'] == 'article'
    used = result['sources_used']
    assert used['ebitda'] == 'article'
    assert used['ebitda_definition'] == 'article'
    assert used['capex'] == 'attachment'
    assert used['revenue'] == '10k'
    assert used['interest'] == 'yahoo'
    assert used['shares'] == '10k'
    assert used['price'] == 'yahoo'
    assert result['baseline']['ebitda']['value'] == 240.0
    assert result['baseline']['capex']['value'] == 36.0
    assert result['baseline']['interest']['value'] == 25.0
    assert result['missing'] == ['nol', 'noa']
    assert [status['status'] for status in result['layers'].values()] == ['used'] * 4
    assert list(result['layers']) == ['article', 'attachment', '10k', 'yahoo']
    assert sec.calls == [('XMPL', FY23)]
    assert result['warnings'] == []


def test_the_merged_baseline_validates():
    root_case = os.path.join(ROOT, 'examples', 'article_case.json')
    with open(root_case, encoding='utf-8') as handle:
        case = json.load(handle)
    article = case['baseline']
    result = build_baseline('XMPL', article=article, yahoo_client=FakeYahoo(rows_by_period={}))
    assert result['ok']
    assert result['period_end'] == FY23
    assert set(result['sources_used'].values()) == {'article'}
    assert result['layers']['10k']['status'] == 'no_user_agent'
    assert result['layers']['yahoo']['status'] == 'no_data_for_year'
    case['baseline'] = result['baseline']
    report = validate_case(case)
    assert report['errors'] == []
    assert report['missing'] == []


def test_lower_sources_only_fill_the_article_year(sec):
    yahoo = FakeYahoo()
    result = build_baseline('XMPL', article=_article(period=FY22, ebitda=180.0),
                            user_agent=UA, yahoo_client=yahoo)
    assert result['period_end'] == FY22
    assert sec.calls == [('XMPL', FY22)]
    assert result['baseline']['revenue']['period_end'] == FY22
    assert result['baseline']['interest']['value'] == 20.0
    assert result['baseline']['date']['value'] == FY22
    assert any('newer fiscal year ending 2023-12-31' in text and 'SEC' in text
               for text in result['warnings'])
    assert any('Yahoo' in text for text in result['warnings'])


def test_trailing_twelve_months_and_other_years_are_skipped(sec):
    article = _article(ebitda=240.0)
    article['revenue'] = leaf(1250.0, 'article', 'twelve months ended June 30', period_end='2024-06-30',
                              basis='ttm', currency='USD')
    article['capex'] = leaf(30.0, 'article', 'capex in fiscal 2022', period_end=FY22,
                            basis='fiscal_year', currency='USD')
    result = build_baseline('XMPL', article=article, user_agent=UA, yahoo_client=FakeYahoo())
    assert result['period_end'] == FY23
    reasons = {item['field']: item['reason'] for item in result['skipped']}
    assert 'ttm' in reasons['revenue']
    assert FY22 in reasons['capex']
    assert result['sources_used']['revenue'] == '10k'
    assert result['sources_used']['capex'] == '10k'


def test_a_canadian_ticker_never_calls_the_sec(sec):
    yahoo = FakeYahoo()
    result = build_baseline('XMPL.TO', user_agent=UA, yahoo_client=yahoo)
    assert result['ok']
    assert sec.calls == []
    assert result['canadian']
    assert result['layers']['10k'] == {'status': 'skipped_canadian'}
    assert result['period_source'] == 'yahoo'
    assert result['period_end'] == FY23
    assert result['baseline']['ebitda']['value'] == pytest.approx(260.0)
    assert result['baseline']['ebitda_definition']['value'] == 'before_sbc'
    assert result['sources_used']['ebitda_definition'] == 'yahoo'
    assert yahoo.calls == ['XMPL.TO']


def test_country_ca_also_skips_the_sec(sec):
    result = build_baseline('XMPL.V', country='CA', user_agent=UA, yahoo_client=FakeYahoo())
    assert result['ok']
    assert sec.calls == []


def test_a_canadian_ticker_without_a_suffix_is_rejected(sec):
    yahoo = FakeYahoo()
    result = build_baseline('XMPL', country='CA', user_agent=UA, yahoo_client=yahoo)
    assert not result['ok']
    assert 'XMPL.TO' in result['error']
    assert sec.calls == []
    assert yahoo.calls == []


def test_a_tsx_listing_reporting_in_dollars_gets_a_converted_price():
    yahoo = FakeYahoo(
        info={'sharesOutstanding': 52e6, 'currentPrice': 60.0,
              'currency': 'CAD', 'financialCurrency': 'USD'},
        fx=0.73,
    )
    result = build_baseline('XMPL.TO', yahoo_client=yahoo)
    assert result['currency'] == 'USD'
    assert result['price']['currency'] == 'USD'
    assert result['price']['value'] == pytest.approx(60.0 * 0.73)
    assert 'CADUSD=X' in result['price']['evidence']
    assert 'CADUSD=X' in yahoo.calls


def test_a_lower_source_in_another_currency_is_skipped():
    yahoo = FakeYahoo(
        info={'sharesOutstanding': 52e6, 'currentPrice': 60.0,
              'currency': 'CAD', 'financialCurrency': 'USD'},
        fx=0.73,
    )
    result = build_baseline('XMPL.TO', article=_article(currency='CAD', ebitda=330.0),
                            yahoo_client=yahoo)
    assert result['currency'] == 'CAD'
    assert result['baseline']['ebitda']['value'] == 330.0
    assert 'revenue' not in result['baseline']
    skipped = {item['field']: item['reason'] for item in result['skipped']}
    assert skipped['revenue'] == 'in USD, but the baseline is in CAD'
    assert result['price']['currency'] == 'CAD'
    assert result['price']['value'] == 60.0


def test_yfinance_not_installed_is_reported(sec, monkeypatch):
    monkeypatch.setitem(sys.modules, 'yfinance', None)
    result = build_baseline('XMPL', user_agent=UA)
    assert result['ok']
    assert result['layers']['yahoo']['status'] == 'not_installed'
    assert result['sources_used']['revenue'] == '10k'
    assert result['price'] is None
    assert 'interest' in result['missing']


def test_a_failing_yfinance_call_is_reported(sec):
    def broken(symbol):
        raise ConnectionError('Yahoo is down')

    result = build_baseline('XMPL', user_agent=UA, yahoo_client=broken)
    assert result['ok']
    assert result['layers']['yahoo']['status'] == 'error'
    assert 'Yahoo is down' in result['layers']['yahoo']['detail']
    assert result['layers']['10k']['status'] == 'used'


def test_no_source_with_a_year_is_not_ok():
    result = build_baseline('XMPL.TO', yahoo_client=FakeYahoo(rows_by_period={}))
    assert not result['ok']
    assert result['layers']['yahoo']['status'] == 'no_data_for_year'


def test_an_explicit_period_end_wins(sec):
    result = build_baseline('XMPL', article=_article(ebitda=240.0), period_end=FY22,
                            user_agent=UA, yahoo_client=FakeYahoo())
    assert result['period_end'] == FY22
    assert result['period_source'] == 'argument'
    assert result['sources_used']['ebitda'] == '10k'


def test_a_52_week_year_end_is_the_same_fiscal_year():
    rows = {'2023-12-30': YAHOO_ROWS[FY23]}
    result = build_baseline('XMPL.TO', article=_article(ebitda=240.0), yahoo_client=FakeYahoo(rows))
    assert result['period_end'] == FY23
    revenue = result['baseline']['revenue']
    assert revenue['source'] == 'yahoo'
    assert revenue['period_end'] == FY23
    assert '2023-12-30' in revenue['note']


def test_yahoo_baseline_maps_annual_rows():
    result = yahoo_baseline('XMPL', client=FakeYahoo())
    assert result['ok']
    assert result['period_end'] == FY23
    assert result['periods'] == [FY23, FY22]
    leaves = result['baseline']
    assert leaves['revenue']['value'] == 1300.0
    assert leaves['revenue']['evidence'] == 'Total Revenue'
    assert leaves['capex']['value'] == 50.0
    assert leaves['interest']['value'] == 25.0
    assert leaves['ebitda']['value'] == pytest.approx(200.0 + 45.0 + 15.0)
    assert leaves['ebitda']['evidence'] == (
        'Operating Income + Depreciation And Amortization + Stock Based Compensation'
    )
    assert leaves['debt']['basis'] == 'fiscal_year'
    assert leaves['debt']['currency'] == 'USD'
    assert result['shares']['value'] == 52.0
    assert result['price']['value'] == 44.0


def test_yahoo_baseline_falls_back_and_marks_missing_rows():
    rows = {FY23: dict(YAHOO_ROWS[FY23])}
    rows[FY23].pop('Stock Based Compensation')
    rows[FY23].pop('Depreciation And Amortization')
    rows[FY23]['Reconciled Depreciation'] = 44e6
    yahoo = FakeYahoo(rows)
    yahoo_rows = FakeTicker(rows, yahoo.info)
    yahoo_rows.income_stmt.loc['Reconciled Depreciation'] = [44e6]

    result = yahoo_baseline('XMPL', client=lambda symbol: yahoo_rows)
    leaves = result['baseline']
    assert leaves['da']['value'] == 44.0
    assert leaves['da']['evidence'] == 'Reconciled Depreciation'
    assert leaves['sbc']['value'] is None
    assert 'Stock Based Compensation' in leaves['sbc']['note']
    assert leaves['ebitda']['value'] == pytest.approx(244.0)
    assert 'was not added' in leaves['ebitda']['note']


def test_yahoo_baseline_for_a_missing_year_lists_the_years():
    result = yahoo_baseline('XMPL', period_end='2020-12-31', client=FakeYahoo())
    assert not result['ok']
    assert result['periods'] == [FY23, FY22]

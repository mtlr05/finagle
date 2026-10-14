'''Fill inputs that no source stated.

Each fill is an assumption whose evidence is the formula and whose
``derived_from`` lists the paths it was computed from, so a later edit can
recompute it. A leaf that already has a value is left alone.
'''

from finagle.agent.schema import is_leaf, leaf
from finagle.agent.yahoo import risk_free_rate

EQUITY_PREMIUM = 0.055
ROICT = 0.15


def apply_defaults(case, currency=None, yahoo_client=None, fetch_yield=True):
    '''Fill the gaps in ``case``. Returns the paths filled and any warnings.

    ``fetch_yield`` loads the 10-year yield when ``rates.risk_free`` is
    missing. An update passes false and uses the yield already stored.
    '''
    if not isinstance(case, dict):
        raise ValueError('case must be an object')
    filled = []
    warnings = []
    baseline = case.setdefault('baseline', {})
    forecast = case.setdefault('forecast', {})
    rates = case.setdefault('rates', {})
    debt = case.setdefault('debt', {})
    distribution = case.setdefault('distribution', {})

    _margins(baseline, forecast, filled)
    _yield_and_rates(rates, currency, yahoo_client, fetch_yield, filled, warnings)
    _sales_growth(baseline, forecast, filled)
    _sales_from_ebitda(baseline, forecast, filled)
    _next_sales_growth(forecast, rates, filled)
    _drop_incomplete_margin(forecast, filled)
    _leverage(baseline, debt, filled)
    _buyback_price_path(distribution, filled)
    _drop_empty(case, ('forecast', 'debt', 'distribution', 'rates', 'baseline'))
    return {'filled': filled, 'warnings': warnings}


def _margins(baseline, forecast, filled):
    margin = forecast.setdefault('margin_path', {})
    ebitda = _number(_val(baseline.get('ebitda')))
    revenue = _number(_val(baseline.get('revenue')))
    if _missing(margin.get('me')) and ebitda not in (None, 0) and revenue not in (None, 0):
        _put(margin, 'me', 'forecast.margin_path.me', ebitda / revenue,
             'Year-0 EBITDA divided by year-0 revenue.',
             ['baseline.ebitda', 'baseline.revenue'], filled)
    me = _number(_val(margin.get('me')))
    if _missing(margin.get('mc')) and me not in (None, 0):
        gross = _number(_val(baseline.get('gross_margin')))
        if gross is None:
            _put(margin, 'mc', 'forecast.margin_path.mc', me,
                 'Gross margin was not found, so the contribution margin equals the EBITDA margin.',
                 ['forecast.margin_path.me'], filled)
        else:
            _put(margin, 'mc', 'forecast.margin_path.mc', (me + gross) / 2.0,
                 'Average of the EBITDA margin and the gross margin.',
                 ['forecast.margin_path.me', 'baseline.gross_margin'], filled)


def _yield_and_rates(rates, currency, yahoo_client, fetch_yield, filled, warnings):
    if fetch_yield and _missing(rates.get('risk_free')) and currency:
        try:
            quote = risk_free_rate(currency, client=yahoo_client)
        except ImportError:
            warnings.append('yfinance is not installed, so the 10-year yield was not loaded.')
            quote = None
        if quote and quote.get('ok'):
            rates['risk_free'] = quote['leaf']
            filled.append('rates.risk_free')
        elif quote:
            warnings.append(quote['error'])
    rf = _number(_val(rates.get('risk_free')))
    if _missing(rates.get('re')) and rf is not None:
        _put(rates, 're', 'rates.re', EQUITY_PREMIUM + rf,
             '0.055 plus the 10-year yield.', ['rates.risk_free'], filled)
    if _missing(rates.get('gt')) and rf is not None:
        _put(rates, 'gt', 'rates.gt', rf,
             'Terminal growth equals the 10-year yield.', ['rates.risk_free'], filled)
    if _missing(rates.get('roict')):
        _put(rates, 'roict', 'rates.roict', ROICT,
             'Terminal ROIC of 15 percent.', [], filled)


def _sales_growth(baseline, forecast, filled):
    if not _missing(forecast.get('ebitda_growth')) or not _missing(forecast.get('ebitda')):
        return
    sales = _series(forecast.get('sales_growth'))
    margin = forecast.get('margin_path') or {}
    me = _number(_val(margin.get('me')))
    mc = _number(_val(margin.get('mc')))
    if not sales or me in (None, 0) or mc in (None, 0):
        return
    derived = _rolled_ebitda_growth(
        sales, me, mc,
        _number(_val(baseline.get('revenue'))),
        _number(_val(baseline.get('ebitda'))),
    )
    _put(forecast, 'ebitda_growth', 'forecast.ebitda_growth', derived,
         'Each year\'s sales growth times the contribution margin divided by that year\'s EBITDA margin.',
         ['forecast.sales_growth', 'forecast.margin_path.me', 'forecast.margin_path.mc'], filled)


def _rolled_ebitda_growth(sales, me, mc, revenue, ebitda):
    '''EBITDA growth from sales growth, updating the margin after each year.

    Year 0 uses the stored EBITDA margin. Each later year uses EBITDA divided
    by revenue after the previous year's growth. Without both year-0 levels,
    every year uses the stored margin.
    '''
    if revenue in (None, 0) or ebitda is None:
        return [mc / me * rate for rate in sales]
    derived = []
    for rate in sales:
        derived.append(mc / me * rate)
        revenue = revenue * (1.0 + rate)
        ebitda = ebitda * (1.0 + derived[-1])
        if revenue != 0:
            me = ebitda / revenue
    return derived


def _sales_from_ebitda(baseline, forecast, filled):
    '''Sales growth that the margin roll turns back into the stated EBITDA growth.'''
    if not _missing(forecast.get('sales_growth')):
        return
    ebitda_growth = _series(forecast.get('ebitda_growth'))
    if not ebitda_growth:
        return
    margin = forecast.get('margin_path') or {}
    me = _number(_val(margin.get('me')))
    mc = _number(_val(margin.get('mc')))
    if me in (None, 0) or mc in (None, 0):
        return
    revenue = _number(_val(baseline.get('revenue')))
    ebitda = _number(_val(baseline.get('ebitda')))
    derived = _rolled_sales_growth(ebitda_growth, me, mc, revenue, ebitda)
    sources = ['forecast.ebitda_growth', 'forecast.margin_path.me', 'forecast.margin_path.mc']
    if revenue not in (None, 0) and ebitda is not None:
        sources.extend(['baseline.revenue', 'baseline.ebitda'])
    _put(forecast, 'sales_growth', 'forecast.sales_growth', derived,
         'Each year\'s EBITDA growth times that year\'s EBITDA margin divided by the contribution margin.',
         sources, filled)


def _rolled_sales_growth(ebitda_growth, me, mc, revenue, ebitda):
    '''Inverse of ``_rolled_ebitda_growth``.

    Without both year-0 levels, every year uses the stored margin.
    '''
    if revenue in (None, 0) or ebitda is None:
        return [rate * me / mc for rate in ebitda_growth]
    derived = []
    for rate in ebitda_growth:
        sales_rate = rate * me / mc
        derived.append(sales_rate)
        revenue = revenue * (1.0 + sales_rate)
        ebitda = ebitda * (1.0 + rate)
        if revenue != 0:
            me = ebitda / revenue
    return derived


def _next_sales_growth(forecast, rates, filled):
    margin = forecast.get('margin_path') or {}
    if _missing(margin.get('me')) or _missing(margin.get('mc')) or not _missing(margin.get('gsnext')):
        return
    sales = _series(forecast.get('sales_growth'))
    if sales:
        _put(margin, 'gsnext', 'forecast.margin_path.gsnext', sales[-1],
             'Last sales growth.', ['forecast.sales_growth'], filled)
        return
    gt = _number(_val(rates.get('gt')))
    if gt is not None:
        _put(margin, 'gsnext', 'forecast.margin_path.gsnext', gt,
             'No sales growth was stated and EBITDA growth could not be turned into one, '
             'so the next sales growth equals terminal growth.',
             ['rates.gt'], filled)


def _drop_incomplete_margin(forecast, filled):
    margin = forecast.get('margin_path')
    if not isinstance(margin, dict):
        return
    if not any(_missing(margin.get(name)) for name in ('me', 'mc', 'gsnext')):
        return
    forecast.pop('margin_path', None)
    growth = forecast.get('ebitda_growth')
    if is_leaf(growth) and 'forecast.margin_path.me' in (growth.get('derived_from') or []):
        forecast.pop('ebitda_growth', None)
    sales = forecast.get('sales_growth')
    if is_leaf(sales) and 'forecast.margin_path.me' in (sales.get('derived_from') or []):
        forecast.pop('sales_growth', None)
    filled[:] = [
        path for path in filled
        if not path.startswith('forecast.margin_path')
        and path not in ('forecast.ebitda_growth', 'forecast.sales_growth')
    ]


def _leverage(baseline, debt, filled):
    if not _missing(debt.get('mode')):
        return
    ebitda = _number(_val(baseline.get('ebitda')))
    balance = _number(_val(baseline.get('debt')))
    if ebitda in (None, 0) or balance is None:
        return
    ratio = balance / ebitda
    _put(debt, 'mode', 'debt.mode', 'target',
         'No debt policy was stated, so the target leverage is the year-0 Debt/EBITDA.',
         ['baseline.debt', 'baseline.ebitda'], filled)
    if _missing(debt.get('leverage')):
        _put(debt, 'leverage', 'debt.leverage', ratio,
             'Year-0 debt divided by year-0 EBITDA.',
             ['baseline.debt', 'baseline.ebitda'], filled)
    if _missing(debt.get('start_year')):
        _put(debt, 'start_year', 'debt.start_year', 1,
             'The target applies from the first forecast year.', [], filled)


def _buyback_price_path(distribution, filled):
    mode = _val(distribution.get('mode'))
    if mode in ('max_buybacks', 'buyback_schedule') and _missing(distribution.get('price_path')):
        _put(distribution, 'price_path', 'distribution.price_path', 'proportional',
             'Buybacks hold the year-1 EV/EBITDA multiple.', [], filled)


def _put(bucket, key, path, value, evidence, derived_from, filled):
    bucket[key] = leaf(value, 'assumption', evidence, derived_from=derived_from)
    filled.append(path)


def _drop_empty(case, names):
    for name in names:
        node = case.get(name)
        if isinstance(node, dict) and not node:
            case.pop(name, None)


def _missing(node):
    return not is_leaf(node) or node.get('value') is None


def _val(node):
    if is_leaf(node):
        return node.get('value')
    return None


def _number(value):
    if isinstance(value, bool) or isinstance(value, (list, dict)):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    return None


def _series(node):
    value = _val(node)
    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        return [float(value)]
    if isinstance(value, list) and value and all(isinstance(item, (int, float)) and not isinstance(item, bool) for item in value):
        return [float(item) for item in value]
    return None

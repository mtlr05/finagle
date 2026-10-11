'''Map yfinance annual statements into case leaves.

yfinance is an optional extra (``pip install -e ".[yahoo]"``) and is only
imported when no ``client`` is passed. It is unofficial: Yahoo limits its
data to personal use, and row names change between versions, so every row
lookup has fallbacks and a missing row becomes a null leaf with a note.

Only annual statements are read. Trailing-twelve-month tables are never
used, because EBITDA growth is measured from the last fiscal year.
'''

import math
from datetime import date, datetime

import pandas as pd

# 52- and 53-week fiscal years end a few days apart from year to year.
PERIOD_TOLERANCE_DAYS = 7
_SCALE = 1_000_000

_ROWS = {
    'revenue': ('Total Revenue', 'Operating Revenue'),
    'operating_income': ('Operating Income', 'Total Operating Income As Reported'),
    'da': ('Depreciation And Amortization', 'Reconciled Depreciation',
           'Depreciation Amortization Depletion'),
    'sbc': ('Stock Based Compensation',),
    'tax': ('Tax Provision',),
    'interest': ('Interest Expense', 'Interest Expense Non Operating'),
    'capex': ('Capital Expenditure',),
    'debt': ('Total Debt',),
    'cash': ('Cash Cash Equivalents And Short Term Investments', 'Cash And Cash Equivalents'),
}
_POSITIVE = ('capex', 'interest')
_NOTES = {
    'capex': 'Positive cash outflow.',
    'interest': 'Interest expense, as a positive number.',
    'debt': 'Yahoo Total Debt can include lease liabilities.',
}
_NOT_IN_YAHOO = ['dwc', 'nol', 'noa', 're', 'rd', 't', 'gt', 'roict', 'dividend']
_GROSS_ROWS = ('Gross Profit',)

# Last close of the 10-year yield for the reporting currency. EUR uses the
# German bund. A currency that is not listed leaves the cost of equity unset.
YIELD_SYMBOLS = {
    'USD': '^TNX',
    'CAD': 'CA10Y=RR',
    'EUR': 'DE10Y=RR',
    'GBP': 'GB10Y=RR',
    'JPY': 'JP10Y=RR',
    'AUD': 'AU10Y=RR',
    'CHF': 'CH10Y=RR',
}


def yahoo_baseline(ticker, period_end=None, client=None):
    '''Year-0 leaves from the yfinance annual statements for one fiscal year.

    ``client`` is a callable that takes a symbol and returns an object with
    ``income_stmt``, ``cashflow``, ``balance_sheet`` and ``info``, like
    ``yfinance.Ticker``. ``period_end`` picks the statement column within 7
    days of that date; without it the latest column is used. When no column
    matches, ``ok`` is false and ``periods`` lists the years Yahoo has.

    Money is in millions of the statement currency (``financialCurrency``).
    ``price`` is in that currency too. When the listing trades in another
    currency, ``price_trading`` is the quoted price and ``price`` is
    converted with the Yahoo FX pair ``{trading}{reporting}=X``.
    '''
    if client is None:
        import yfinance
        client = yfinance.Ticker
    ticker = str(ticker).strip()
    if not ticker:
        raise ValueError('ticker is required')
    handle = client(ticker)
    frames = [
        _frame(getattr(handle, 'income_stmt', None)),
        _frame(getattr(handle, 'cashflow', None)),
        _frame(getattr(handle, 'balance_sheet', None)),
    ]
    info = _info(handle)
    periods = _periods(frames)
    latest = periods[0] if periods else None
    result = {
        'ticker': ticker,
        'latest_period_end': latest,
        'periods': periods,
    }
    if not periods:
        result.update(ok=False, period_end=period_end,
                      error='Yahoo has no annual statements for %s' % ticker)
        return result
    chosen = _match(periods, period_end) if period_end else latest
    if chosen is None:
        result.update(ok=False, period_end=period_end,
                      error='Yahoo has no annual statement within %d days of %s'
                      % (PERIOD_TOLERANCE_DAYS, period_end))
        return result

    reporting = _currency(info.get('financialCurrency') or info.get('currency'))
    trading = _currency(info.get('currency') or reporting)
    leaves = _leaves(frames, chosen, reporting)
    today = date.today().isoformat()
    result.update(
        ok=True,
        period_end=chosen,
        currency=reporting,
        trading_currency=trading,
        units='millions of %s; shares in millions' % (reporting or 'the reporting currency'),
        baseline=leaves,
        shares=_shares(info, today),
        not_in_yahoo=list(_NOT_IN_YAHOO),
        note=(
            'Annual statements, not trailing twelve months. yfinance is '
            'unofficial; check figures against the filing when they matter.'
        ),
    )
    quoted, evidence = _quoted_price(handle, info)
    if trading and reporting and trading != reporting:
        result['price_trading'] = _price_leaf(quoted, trading, today, evidence)
        result['price'] = _converted_price(client, quoted, trading, reporting, today, evidence)
    else:
        result['price'] = _price_leaf(quoted, trading or reporting, today, evidence)
    return result


def _leaves(frames, period_end, currency):
    found = {}
    for key, names in _ROWS.items():
        found[key] = _lookup(frames, names, period_end)

    leaves = {
        'date': {
            'value': period_end,
            'source': 'yahoo',
            'evidence': 'annual statement column',
            'period_end': period_end,
            'basis': 'fiscal_year',
        },
    }
    for key in ('revenue', 'capex', 'da', 'tax', 'interest', 'sbc', 'debt', 'cash'):
        row, raw = found[key]
        value = None if raw is None else raw / _SCALE
        if value is not None and key in _POSITIVE:
            value = abs(value)
        note = _NOTES.get(key, '')
        if raw is None:
            note = 'Yahoo has no %s row for this year.' % ' / '.join(_ROWS[key])
        leaves[key] = _leaf(value, row or _ROWS[key][0], period_end, currency, note)

    leaves['ebitda'] = _ebitda(found, period_end, currency)
    revenue = leaves['revenue']['value']
    grow, graw = _lookup(frames, _GROSS_ROWS, period_end)
    if graw is not None and revenue:
        leaves['gross_margin'] = _leaf(
            (graw / _SCALE) / revenue,
            '%s / %s' % (grow, leaves['revenue']['evidence']),
            period_end,
            None,
            'Gross profit divided by revenue.',
        )
    leaves['ebitda_definition'] = {
        'value': 'before_sbc',
        'source': 'yahoo',
        'evidence': 'rebuilt as Operating Income + D&A + Stock Based Compensation',
    }
    return leaves


def _ebitda(found, period_end, currency):
    oi_row, operating = found['operating_income']
    da_row, da = found['da']
    sbc_row, sbc = found['sbc']
    if operating is None:
        return _leaf(None, 'Operating Income', period_end, currency,
                     'Operating income was not found, so EBITDA was not calculated.')
    if da is None:
        return _leaf(None, 'Depreciation And Amortization', period_end, currency,
                     'Depreciation and amortization were not found, so EBITDA was not calculated.')
    raw = operating + da
    rows = [oi_row, da_row]
    if sbc is None:
        note = ('Operating income plus depreciation and amortization. '
                'Stock-based compensation was not found, so it was not added.')
    else:
        raw += sbc
        rows.append(sbc_row)
        note = ('Operating income plus depreciation and amortization plus '
                'stock-based compensation. Yahoo\'s own EBITDA row does not add back SBC.')
    return _leaf(raw / _SCALE, ' + '.join(rows), period_end, currency, note)


def _leaf(value, evidence, period_end, currency, note=''):
    item = {
        'value': value,
        'source': 'yahoo',
        'evidence': evidence,
        'period_end': period_end,
        'basis': 'fiscal_year',
    }
    if currency:
        item['currency'] = currency
    if note:
        item['note'] = note
    return item


def _shares(info, today):
    raw = _number(info.get('sharesOutstanding'))
    item = {
        'value': None if raw is None else raw / _SCALE,
        'source': 'yahoo',
        'evidence': 'info.sharesOutstanding',
        'as_of': today,
    }
    if raw is None:
        item['note'] = 'Yahoo has no sharesOutstanding for this ticker.'
    return item


def _quoted_price(handle, info):
    value = _number(info.get('previousClose'))
    if value is not None:
        return value, 'info.previousClose'
    close, _as_of = _last_close(handle)
    if close is not None:
        return close, 'history Close'
    return None, 'info.previousClose'


def _last_close(handle):
    try:
        history = handle.history(period='5d')
    except Exception:
        return None, None
    if history is None or 'Close' not in getattr(history, 'columns', ()):
        return None, None
    closes = history['Close'].dropna()
    if closes.empty:
        return None, None
    return float(closes.iloc[-1]), _date_text(closes.index[-1])


def _price_leaf(value, currency, today, evidence):
    item = {'value': value, 'source': 'yahoo', 'evidence': evidence, 'as_of': today}
    if currency:
        item['currency'] = currency
    if value is None:
        item['note'] = 'Yahoo has no previous close for this ticker.'
    return item


def _converted_price(client, quoted, trading, reporting, today, evidence):
    pair = '%s%s=X' % (trading, reporting)
    rate, as_of, error = _fx_rate(client, pair)
    if quoted is None or rate is None:
        item = _price_leaf(None, reporting, today, evidence + ' converted with ' + pair)
        item['note'] = (
            'The %s quote could not be converted to %s: %s'
            % (trading, reporting, error or 'no price')
        )
        return item
    item = _price_leaf(quoted * rate, reporting, today,
                       '%s %g %s x %s %.6g on %s'
                       % (evidence, quoted, trading, pair, rate, as_of))
    item['note'] = 'Converted from the %s listing price to the reporting currency.' % trading
    return item


def risk_free_rate(currency, client=None):
    '''Last close of the 10-year yield for the reporting currency, as a decimal.

    A quote above 20 is the CBOE index and is divided by 1,000. A quote above
    1 is a percent and is divided by 100. Anything else is already a decimal.
    An unknown currency, or a quote that cannot be loaded, returns ok false.
    '''
    code = str(currency or '').strip().upper()
    symbol = YIELD_SYMBOLS.get(code)
    if not symbol:
        return {
            'ok': False,
            'currency': code or None,
            'error': 'No 10-year yield is mapped for %s, so the cost of equity and '
                     'terminal growth are left unset.' % (code or 'this currency'),
        }
    if client is None:
        import yfinance
        client = yfinance.Ticker
    try:
        handle = client(symbol)
        info = _info(handle)
        raw, evidence = _quoted_price(handle, info)
        as_of = date.today().isoformat()
        if evidence == 'history Close':
            _close, closed_on = _last_close(handle)
            as_of = closed_on or as_of
    except Exception as exc:
        return {
            'ok': False,
            'symbol': symbol,
            'error': 'The 10-year yield %s could not be loaded: %s: %s'
                     % (symbol, type(exc).__name__, exc),
        }
    if raw is None:
        return {
            'ok': False,
            'symbol': symbol,
            'error': 'Yahoo has no last close for the 10-year yield %s.' % symbol,
        }
    value, how = _yield_decimal(raw)
    return {
        'ok': True,
        'symbol': symbol,
        'raw': raw,
        'value': value,
        'leaf': {
            'value': value,
            'source': 'yahoo',
            'evidence': '%s %s %g on %s, %s' % (symbol, evidence, raw, as_of, how),
            'as_of': as_of,
        },
    }


def _yield_decimal(raw):
    if raw > 20:
        return raw / 1000.0, 'CBOE index divided by 1,000'
    if raw > 1:
        return raw / 100.0, 'percent divided by 100'
    return raw, 'already a decimal'


def _fx_rate(client, pair):
    try:
        handle = client(pair)
        history = handle.history(period='5d')
    except Exception as exc:
        return None, None, '%s: %s' % (type(exc).__name__, exc)
    if history is None or 'Close' not in getattr(history, 'columns', ()):
        return None, None, 'no %s history' % pair
    closes = history['Close'].dropna()
    if closes.empty:
        return None, None, 'no %s history' % pair
    return float(closes.iloc[-1]), _date_text(closes.index[-1]), None


def _lookup(frames, names, period_end):
    for name in names:
        for frame in frames:
            if frame is None:
                continue
            row = _row(frame, name)
            if row is None:
                continue
            column = _column(frame, period_end)
            if column is None:
                continue
            value = _number(frame.at[row, column])
            if value is not None:
                return row, value
    return None, None


def _row(frame, name):
    wanted = _squash(name)
    for row in frame.index:
        if _squash(row) == wanted:
            return row
    return None


def _squash(text):
    return ''.join(str(text).split()).lower()


def _column(frame, period_end):
    for column in frame.columns:
        if _date_text(column) == period_end:
            return column
    return None


def _periods(frames):
    found = set()
    for frame in frames:
        if frame is None:
            continue
        for column in frame.columns:
            text = _date_text(column)
            if text:
                found.add(text)
    return sorted(found, reverse=True)


def _match(periods, period_end):
    target = _parse(period_end)
    if target is None:
        return None
    best = None
    for text in periods:
        days = abs((_parse(text) - target).days)
        if days <= PERIOD_TOLERANCE_DAYS and (best is None or days < best[0]):
            best = (days, text)
    return best[1] if best else None


def _frame(value):
    if isinstance(value, pd.DataFrame) and not value.empty:
        return value
    return None


def _info(handle):
    try:
        info = handle.info
    except Exception:
        return {}
    return info if isinstance(info, dict) else {}


def _currency(value):
    return str(value).strip().upper() if value else None


def _date_text(value):
    try:
        return pd.Timestamp(value).strftime('%Y-%m-%d')
    except (TypeError, ValueError):
        return None


def _parse(text):
    try:
        return datetime.strptime(str(text)[:10], '%Y-%m-%d')
    except (TypeError, ValueError):
        return None


def _number(value):
    if value is None or isinstance(value, bool):
        return None
    try:
        number = float(value)
    except (TypeError, ValueError):
        return None
    if math.isnan(number) or math.isinf(number):
        return None
    return number

'''Expected inflation for the reporting currency.

USD uses the FRED 10-year breakeven series, which matches the 10-year yield
behind terminal growth. Other mapped currencies use the IMF World Economic
Outlook CPI projection for the furthest year. Neither call needs a key.
``fetch`` is a callable from a URL to response text, so tests pass a fake.
'''

import json
import urllib.request
from datetime import date

FRED_SERIES = 'T10YIE'
FRED_URL = 'https://fred.stlouisfed.org/graph/fredgraph.csv?id=' + FRED_SERIES
IMF_URL = 'https://www.imf.org/external/datamapper/api/v1/PCPIPCH/%s'

# Same currencies as the 10-year yield. EUR uses Germany, matching the bund.
IMF_COUNTRIES = {
    'CAD': 'CAN',
    'EUR': 'DEU',
    'GBP': 'GBR',
    'JPY': 'JPN',
    'AUD': 'AUS',
    'CHF': 'CHE',
}


def expected_inflation(currency, fetch=None):
    '''Expected inflation for ``currency``, as a decimal.

    An unknown currency, or a quote that cannot be loaded, returns ok false.
    '''
    code = str(currency or '').strip().upper()
    if code == 'USD':
        return _fred(fetch)
    iso = IMF_COUNTRIES.get(code)
    if not iso:
        return {
            'ok': False,
            'currency': code or None,
            'error': (
                'No inflation series is mapped for %s, so inflation is left unset.'
                % (code or 'this currency')
            ),
        }
    return _imf(iso, fetch)


def _fred(fetch):
    try:
        text = _load(FRED_URL, fetch)
    except Exception as exc:
        return _failed('fred', FRED_SERIES, exc)
    parsed = _last_fred(text)
    if parsed is None:
        return {
            'ok': False,
            'symbol': FRED_SERIES,
            'error': 'FRED has no observation for %s.' % FRED_SERIES,
        }
    observed, raw = parsed
    value = raw / 100.0
    return {
        'ok': True,
        'symbol': FRED_SERIES,
        'raw': raw,
        'value': value,
        'leaf': {
            'value': value,
            'source': 'fred',
            'evidence': 'FRED %s %g on %s, percent divided by 100' % (FRED_SERIES, raw, observed),
            'as_of': observed,
        },
    }


def _imf(iso, fetch):
    url = IMF_URL % iso
    try:
        text = _load(url, fetch)
        payload = json.loads(text)
    except Exception as exc:
        return _failed('imf', iso, exc)
    parsed = _furthest_imf(payload, iso)
    if parsed is None:
        return {
            'ok': False,
            'symbol': iso,
            'error': 'IMF has no CPI projection for %s.' % iso,
        }
    year, raw = parsed
    value = raw / 100.0
    return {
        'ok': True,
        'symbol': iso,
        'raw': raw,
        'value': value,
        'leaf': {
            'value': value,
            'source': 'imf',
            'evidence': (
                'IMF WEO PCPIPCH %s %s %g, percent divided by 100' % (iso, year, raw)
            ),
            'as_of': '%s-01-01' % year,
        },
    }


def _load(url, fetch):
    if fetch is not None:
        return fetch(url)
    request = urllib.request.Request(url, headers={'User-Agent': 'finagle'})
    with urllib.request.urlopen(request, timeout=30) as response:
        return response.read().decode('utf-8')


def _last_fred(text):
    found = None
    for line in str(text or '').splitlines():
        parts = [part.strip() for part in line.split(',')]
        if len(parts) < 2 or not parts[0][:1].isdigit():
            continue
        raw = parts[1]
        if raw in ('', '.'):
            continue
        try:
            number = float(raw)
        except ValueError:
            continue
        found = (parts[0][:10], number)
    return found


def _furthest_imf(payload, iso):
    values = payload.get('values') if isinstance(payload, dict) else None
    series = values.get('PCPIPCH') if isinstance(values, dict) else None
    by_year = series.get(iso) if isinstance(series, dict) else None
    if not isinstance(by_year, dict):
        return None
    found = None
    for year, raw in by_year.items():
        try:
            number = float(raw)
            year_number = int(str(year)[:4])
        except (TypeError, ValueError):
            continue
        if found is None or year_number > found[0]:
            found = (year_number, number)
    if found is None:
        return None
    return str(found[0]), found[1]


def _failed(source, symbol, exc):
    return {
        'ok': False,
        'symbol': symbol,
        'error': 'The %s inflation series %s could not be loaded: %s: %s'
                 % (source, symbol, type(exc).__name__, exc),
    }

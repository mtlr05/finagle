'''Fill year 0 from several sources in a fixed priority.

``build_baseline`` takes the leaves the bot read from the article and from
an attached filing, then fills the remaining fields from the SEC and from
yfinance. The first source with a value for the chosen fiscal year wins
each field. Canadian companies never call the SEC.
'''

import copy
from datetime import datetime

from finagle.agent.baseline import sec_baseline
from finagle.agent.schema import BASELINE_PRIORITY
from finagle.agent.yahoo import PERIOD_TOLERANCE_DAYS, yahoo_baseline

CANADIAN_SUFFIXES = ('.TO', '.V', '.CN', '.NE')
FIELDS = (
    'date', 'revenue', 'ebitda', 'capex', 'da', 'tax', 'interest', 'sbc',
    'debt', 'cash', 'nol', 'noa',
)
_UNITLESS = ('date', 'ebitda_definition')
_NO_MATCH = 10 ** 6


def build_baseline(ticker, article=None, attachment=None, user_agent=None,
                   country=None, period_end=None, yahoo_client=None):
    '''Merge year-0 leaves: article, then attachment, then SEC, then yfinance.

    ``article`` and ``attachment`` map baseline fields (and optionally
    ``shares``) to provenance leaves the bot wrote. Reported leaves need
    ``basis: "fiscal_year"`` and a ``period_end``; trailing-twelve-month and
    quarterly leaves are left out and listed in ``skipped``.

    The baseline year is ``period_end`` when given, else the year of the
    article's figures, else the attachment's, else the latest SEC 10-K,
    else the latest yfinance annual statement. Lower sources only fill that
    year. ``user_agent`` is needed for the SEC. A ticker ending in .TO, .V,
    .CN or .NE, or ``country="CA"``, is Canadian and skips the SEC.
    '''
    ticker = str(ticker or '').strip()
    if not ticker:
        raise ValueError('ticker is required')
    canadian = _is_canadian(ticker, country)
    if canadian and not ticker.upper().endswith(CANADIAN_SUFFIXES):
        return {
            'ok': False,
            'ticker': ticker,
            'error': (
                '%s is Canadian but has no exchange suffix. Pass the TSX symbol '
                'yfinance uses, such as %s.TO (or .V for the TSX Venture).'
                % (ticker, ticker.split('.')[0].upper())
            ),
        }

    skipped = []
    warnings = []
    layers = {}
    given = {}
    for name, leaves in (('article', article), ('attachment', attachment)):
        if not leaves:
            layers[name] = {'status': 'not_given'}
            continue
        given[name] = _clean(name, leaves)

    period = str(period_end) if period_end else None
    period_source = 'argument' if period else None
    for name in ('article', 'attachment'):
        if period is None and name in given:
            period = _source_period(given[name])
            if period:
                period_source = name

    feeds = {}
    if canadian:
        layers['10k'] = {'status': 'skipped_canadian'}
    elif not user_agent or not str(user_agent).strip():
        layers['10k'] = {'status': 'no_user_agent',
                         'detail': 'The SEC needs a user agent with a name and email.'}
    else:
        feeds['10k'] = _load_sec(ticker, user_agent, period, layers)
        if period is None and feeds['10k']:
            period = feeds['10k']['period_end']
            period_source = '10k'

    feeds['yahoo'] = _load_yahoo(ticker, period, yahoo_client, layers)
    if period is None and feeds['yahoo']:
        period = feeds['yahoo']['period_end']
        period_source = 'yahoo'

    if period is None:
        return {
            'ok': False,
            'ticker': ticker,
            'error': 'No source has an annual baseline year for %s.' % ticker,
            'skipped': skipped,
            'layers': layers,
        }

    candidates = {}
    for name in BASELINE_PRIORITY:
        if name in given:
            candidates[name] = _filter_reported(name, given[name], period, skipped)
        elif feeds.get(name):
            candidates[name] = _normalize(name, feeds[name]['baseline'], period)

    currency = _reporting_currency(candidates)
    baseline = {}
    sources_used = {}
    for field in FIELDS:
        for name in BASELINE_PRIORITY:
            node = candidates.get(name, {}).get(field)
            if node is None or node.get('value') is None:
                continue
            if not _same_currency(node, currency):
                skipped.append({
                    'source': name, 'field': field,
                    'reason': 'in %s, but the baseline is in %s' % (node.get('currency'), currency),
                })
                continue
            baseline[field] = node
            sources_used[field] = name
            break

    if 'date' not in baseline:
        baseline['date'] = _date_leaf(period, period_source)
        sources_used['date'] = baseline['date']['source']
    else:
        baseline['date']['value'] = period

    ebitda_source = sources_used.get('ebitda')
    if ebitda_source:
        definition = candidates.get(ebitda_source, {}).get('ebitda_definition')
        if definition is None and ebitda_source == '10k':
            definition = {
                'value': 'before_sbc',
                'source': '10k',
                'evidence': 'last_10k adds stock-based compensation back to EBITDA',
            }
        if definition is not None:
            baseline['ebitda_definition'] = definition
            sources_used['ebitda_definition'] = ebitda_source

    shares, shares_source = _first_shares(given, feeds)
    if shares_source:
        sources_used['shares'] = shares_source
    price = _yahoo_price(feeds.get('yahoo'), currency, skipped)
    if price is not None:
        sources_used['price'] = 'yahoo'

    for name in BASELINE_PRIORITY:
        if name in layers and layers[name]['status'] != 'loaded':
            continue
        layers[name] = {'status': 'used' if name in sources_used.values() else 'unused'}
    layers = {name: layers[name] for name in BASELINE_PRIORITY if name in layers}

    for name in ('10k', 'yahoo'):
        latest = (feeds.get(name) or {}).get('latest_period_end') or layers.get(name, {}).get('latest_period_end')
        if latest and str(latest) > period and _days(period, latest) > PERIOD_TOLERANCE_DAYS:
            warnings.append(
                'A newer fiscal year ending %s exists in the %s data. The baseline '
                'stays on the year ending %s, so growth rates must be measured from it.'
                % (latest, 'SEC' if name == '10k' else 'Yahoo', period)
            )

    missing = [field for field in FIELDS if field not in baseline]
    if 'ebitda_definition' not in baseline:
        missing.append('ebitda_definition')
    return {
        'ok': True,
        'ticker': ticker,
        'canadian': canadian,
        'period_end': period,
        'period_source': period_source,
        'currency': currency,
        'baseline': baseline,
        'shares': shares,
        'price': price,
        'sources_used': sources_used,
        'missing': missing,
        'skipped': skipped,
        'warnings': warnings,
        'layers': layers,
        'note': (
            'Year 0 is the fiscal year ending %s. Copy these leaves into case.baseline, '
            'shares and price into case.market, and fill what is missing (nol and noa are '
            'never in the data feeds) as assumptions.' % period
        ),
    }


def _is_canadian(ticker, country):
    if country and str(country).strip().upper() in ('CA', 'CAN', 'CANADA'):
        return True
    return ticker.upper().endswith(CANADIAN_SUFFIXES)


def _clean(name, leaves):
    cleaned = {}
    for field, node in leaves.items():
        if isinstance(node, dict) and 'value' in node:
            node = copy.deepcopy(node)
            node.setdefault('source', name)
            cleaned[field] = node
    return cleaned


def _source_period(leaves):
    date = leaves.get('date')
    if _fiscal(date) and date.get('value'):
        return str(date['value'])
    ends = [
        str(node['period_end']) for field, node in leaves.items()
        if field in FIELDS and _fiscal(node) and node.get('period_end')
    ]
    return max(ends) if ends else None


def _fiscal(node):
    return isinstance(node, dict) and node.get('basis') == 'fiscal_year'


def _filter_reported(name, leaves, period, skipped):
    kept = {}
    for field, node in leaves.items():
        if field == 'shares':
            continue
        if field == 'ebitda_definition' or node.get('source') == 'assumption':
            kept[field] = node
            continue
        basis = node.get('basis')
        end = node.get('period_end')
        if basis != 'fiscal_year':
            skipped.append({
                'source': name, 'field': field,
                'reason': 'basis is %s; year 0 must be a fiscal year' % (basis or 'missing'),
            })
            continue
        if not end:
            skipped.append({'source': name, 'field': field, 'reason': 'no period_end'})
            continue
        if _days(period, end) > PERIOD_TOLERANCE_DAYS:
            skipped.append({
                'source': name, 'field': field,
                'reason': 'period_end %s is not the baseline year ending %s' % (end, period),
            })
            continue
        kept[field] = _restamp(node, period)
    return kept


def _normalize(name, leaves, period):
    return {field: _restamp(copy.deepcopy(node), period) for field, node in leaves.items()}


def _restamp(node, period):
    end = node.get('period_end')
    if end and str(end) != period:
        note = node.get('note', '')
        extra = 'Reported period end %s, within %d days of %s.' % (end, PERIOD_TOLERANCE_DAYS, period)
        node['note'] = (note + ' ' + extra).strip()
        node['period_end'] = period
    return node


def _reporting_currency(candidates):
    for name in BASELINE_PRIORITY:
        for field in FIELDS:
            node = candidates.get(name, {}).get(field)
            if field in _UNITLESS or not node or node.get('value') is None:
                continue
            if node.get('currency'):
                return str(node['currency']).upper()
    return None


def _same_currency(node, currency):
    found = node.get('currency')
    return not found or not currency or str(found).upper() == currency


def _date_leaf(period, period_source):
    if period_source in (None, 'argument'):
        return {
            'value': period,
            'source': 'assumption',
            'evidence': 'period_end passed to build_baseline',
        }
    return {
        'value': period,
        'source': period_source,
        'evidence': 'period_end of the %s fiscal-year figures' % period_source,
        'period_end': period,
        'basis': 'fiscal_year',
    }


def _first_shares(given, feeds):
    for name in BASELINE_PRIORITY:
        if name in given:
            node = given[name].get('shares')
        else:
            node = (feeds.get(name) or {}).get('shares')
        if node and node.get('value') is not None:
            return node, name
    return None, None


def _yahoo_price(feed, currency, skipped):
    if not feed:
        return None
    for key in ('price', 'price_trading'):
        node = feed.get(key)
        if node and node.get('value') is not None and _same_currency(node, currency):
            return node
    node = feed.get('price')
    if node and node.get('value') is not None:
        skipped.append({
            'source': 'yahoo', 'field': 'price',
            'reason': 'in %s, but the baseline is in %s' % (node.get('currency'), currency),
        })
    return None


def _load_sec(ticker, user_agent, period, layers):
    try:
        result = sec_baseline(ticker, user_agent, period_end=period)
        if not result.get('ok') and period:
            near = _nearest(result.get('periods') or [], period)
            if near:
                result = sec_baseline(ticker, user_agent, period_end=near)
    except LookupError as exc:
        layers['10k'] = {'status': 'no_data', 'detail': str(exc)}
        return None
    except Exception as exc:
        layers['10k'] = {'status': 'error', 'detail': '%s: %s' % (type(exc).__name__, exc)}
        return None
    if not result.get('ok'):
        layers['10k'] = {
            'status': 'no_data_for_year',
            'detail': result.get('error'),
            'latest_period_end': result.get('latest_period_end'),
        }
        return None
    layers['10k'] = {'status': 'loaded'}
    return result


def _load_yahoo(ticker, period, client, layers):
    try:
        result = yahoo_baseline(ticker, period_end=period, client=client)
    except ImportError:
        layers['yahoo'] = {
            'status': 'not_installed',
            'detail': 'pip install -e ".[yahoo]" to add yfinance.',
        }
        return None
    except Exception as exc:
        layers['yahoo'] = {'status': 'error', 'detail': '%s: %s' % (type(exc).__name__, exc)}
        return None
    if not result.get('ok'):
        layers['yahoo'] = {
            'status': 'no_data_for_year',
            'detail': result.get('error'),
            'latest_period_end': result.get('latest_period_end'),
        }
        return None
    layers['yahoo'] = {'status': 'loaded'}
    return result


def _nearest(periods, period):
    best = None
    for text in periods:
        days = _days(period, text)
        if days <= PERIOD_TOLERANCE_DAYS and (best is None or days < best[0]):
            best = (days, text)
    return best[1] if best else None


def _days(first, second):
    try:
        one = datetime.strptime(str(first)[:10], '%Y-%m-%d')
        two = datetime.strptime(str(second)[:10], '%Y-%m-%d')
    except (TypeError, ValueError):
        return _NO_MATCH
    return abs((two - one).days)

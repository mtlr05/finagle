'''Fill a case from several sources in a fixed priority.

``build_baseline`` takes the leaves the bot read from the prompt, the
primary article, and any other attachments, then fills the remaining
statement fields from the SEC and from yfinance. The first source with a
value for the chosen fiscal year wins each field. Canadian companies never
call the SEC. Inputs that are still empty are filled by ``apply_defaults``.
'''

import copy
from datetime import datetime

from finagle.agent.baseline import sec_baseline
from finagle.agent.defaults import apply_defaults
from finagle.agent.schema import SOURCES, is_leaf
from finagle.agent.yahoo import PERIOD_TOLERANCE_DAYS, yahoo_baseline

CANADIAN_SUFFIXES = ('.TO', '.V', '.CN', '.NE')
FIELDS = (
    'date', 'revenue', 'ebitda', 'capex', 'da', 'tax', 'interest', 'sbc',
    'debt', 'cash', 'nol', 'noa', 'gross_margin',
)
_OPTIONAL = ('gross_margin',)
_UNITLESS = ('date', 'ebitda_definition', 'gross_margin')
_NO_MATCH = 10 ** 6
_NESTED = ('baseline', 'market', 'rates', 'forecast', 'debt', 'distribution', 'leaves', 'name')
_RATE_FIELDS = ('re', 'rd', 't', 'te', 'gt', 'roict', 'risk_free')
_DEBT_FIELDS = ('mode', 'leverage', 'start_year', 'values')
_DISTRIBUTION_FIELDS = ('mode', 'schedule', 'price', 'price_path')
_FORECAST_FIELDS = ('ebitda_growth', 'sales_growth', 'ebitda', 'capex', 'sbc', 'sbc_rate_terminal')


def build_baseline(ticker, article=None, attachment=None, attachments=None,
                   prompt=None, user_agent=None, country=None, period_end=None,
                   yahoo_client=None):
    '''Merge inputs, then fill the remaining default assumptions.

    Priority is a figure stated in ``prompt``, then the primary ``article``,
    then ``attachments`` in the order given, then the SEC 10-K, then
    yfinance. ``attachment`` is the older single-filing argument and is
    treated as a one-item list. Each attachment is ``{name, leaves}`` or a
    flat leaf map. ``sources_used`` names the winning attachment, such as
    ``attachment:investor presentation``.

    Reported baseline leaves need ``basis: "fiscal_year"`` and a
    ``period_end``. Trailing-twelve-month and quarterly leaves are left out
    and listed in ``skipped``. The baseline year is ``period_end`` when
    given, else the year of the highest source that states one. Lower
    sources only fill that year. ``user_agent`` is needed for the SEC. A
    ticker ending in .TO, .V, .CN or .NE, or ``country="CA"``, is Canadian
    and skips the SEC.
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
    sections = {}
    order = []
    for label, source_name, payload in _inputs(prompt, article, attachment, attachments):
        if not payload:
            layers[label] = {'status': 'not_given'}
            order.append(label)
            continue
        sections[label] = _section(source_name, payload)
        order.append(label)

    period = str(period_end) if period_end else None
    period_source = 'argument' if period else None
    for label in order:
        if period is None and label in sections:
            period = _source_period(sections[label]['baseline'])
            if period:
                period_source = label

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
    order.append('10k')

    feeds['yahoo'] = _load_yahoo(ticker, period, yahoo_client, layers)
    if period is None and feeds['yahoo']:
        period = feeds['yahoo']['period_end']
        period_source = 'yahoo'
    order.append('yahoo')

    if period is None:
        return {
            'ok': False,
            'ticker': ticker,
            'error': 'No source has an annual baseline year for %s.' % ticker,
            'skipped': skipped,
            'layers': layers,
        }

    candidates = {}
    for label in order:
        if label in sections:
            candidates[label] = _filter_reported(label, sections[label]['baseline'], period, skipped)
        elif feeds.get(label):
            candidates[label] = _normalize(label, feeds[label]['baseline'], period)

    currency = _reporting_currency(candidates, order)
    baseline = {}
    sources_used = {}
    for field in FIELDS:
        for label in order:
            node = candidates.get(label, {}).get(field)
            if node is None or node.get('value') is None:
                continue
            if not _same_currency(node, currency):
                skipped.append({
                    'source': label, 'field': field,
                    'reason': 'in %s, but the baseline is in %s' % (node.get('currency'), currency),
                })
                continue
            baseline[field] = node
            sources_used[field] = label
            break

    if 'date' not in baseline:
        baseline['date'] = _date_leaf(period, period_source)
        sources_used['date'] = baseline['date']['source']
    else:
        baseline['date']['value'] = period

    ebitda_source = sources_used.get('ebitda')
    if ebitda_source:
        definition = candidates.get(ebitda_source, {}).get('ebitda_definition')
        if definition is None and ebitda_source.split(':')[0] == '10k':
            definition = {
                'value': 'before_sbc',
                'source': '10k',
                'evidence': 'last_10k adds stock-based compensation back to EBITDA',
            }
        if definition is not None:
            baseline['ebitda_definition'] = definition
            sources_used['ebitda_definition'] = ebitda_source

    picked = _pick_sections(sections, order)
    shares, shares_source = _first_shares(sections, feeds, order)
    if shares_source:
        sources_used['shares'] = shares_source
    price, price_source = _first_price(sections, feeds, currency, skipped, order)
    if price_source:
        sources_used['price'] = price_source

    for label in order:
        if label in layers and layers[label]['status'] != 'loaded':
            continue
        layers[label] = {'status': 'used' if label in sources_used.values() else 'unused'}
    layers = {label: layers[label] for label in order if label in layers}

    for name in ('10k', 'yahoo'):
        latest = (feeds.get(name) or {}).get('latest_period_end') or layers.get(name, {}).get('latest_period_end')
        if latest and str(latest) > period and _days(period, latest) > PERIOD_TOLERANCE_DAYS:
            warnings.append(
                'A newer fiscal year ending %s exists in the %s data. The baseline '
                'stays on the year ending %s, so growth rates must be measured from it.'
                % (latest, 'SEC' if name == '10k' else 'Yahoo', period)
            )

    case = {
        'ticker': ticker,
        'baseline': baseline,
        'market': {},
        'rates': picked['rates'],
        'forecast': picked['forecast'],
        'debt': picked['debt'],
        'distribution': picked['distribution'],
    }
    if shares:
        case['market']['shares'] = shares
    if price:
        case['market']['price'] = price
    for key, source in _section_sources(picked).items():
        sources_used[key] = source
    defaults = apply_defaults(case, currency=currency, yahoo_client=yahoo_client)
    warnings.extend(defaults['warnings'])
    for path in defaults['filled']:
        sources_used.setdefault(path, 'assumption')

    missing = [field for field in FIELDS if field not in _OPTIONAL and field not in baseline]
    if 'ebitda_definition' not in baseline:
        missing.append('ebitda_definition')
    return {
        'ok': True,
        'ticker': ticker,
        'canadian': canadian,
        'period_end': period,
        'period_source': period_source,
        'currency': currency,
        'baseline': case['baseline'],
        'shares': case.get('market', {}).get('shares'),
        'price': case.get('market', {}).get('price'),
        'rates': case.get('rates') or {},
        'forecast': case.get('forecast') or {},
        'debt': case.get('debt') or {},
        'distribution': case.get('distribution') or {},
        'case': case,
        'assumptions': defaults['filled'],
        'sources_used': sources_used,
        'missing': missing,
        'skipped': skipped,
        'warnings': warnings,
        'layers': layers,
        'note': (
            'Year 0 is the fiscal year ending %s. case holds the merged leaves and the '
            'computed assumptions. Copy it, fill what is still missing, then validate_case. '
            'nol and noa are never in the data feeds.' % period
        ),
    }


def _inputs(prompt, article, attachment, attachments):
    '''(label, leaf source, payload) in priority order.'''
    rows = []
    if prompt:
        rows.append(('prompt', 'prompt', prompt))
    rows.append(('article', 'article', article))
    items = _attachment_items(attachment, attachments)
    if not items:
        rows.append(('attachment', 'attachment', None))
    seen = {}
    for item in items:
        name = item.get('name')
        label = 'attachment:%s' % name if name else 'attachment'
        if label in seen:
            seen[label] += 1
            label = '%s:%d' % (label, seen[label])
        else:
            seen[label] = 1
        rows.append((label, 'attachment', item.get('leaves')))
    return rows


def _attachment_items(attachment, attachments):
    items = []
    if attachment:
        if isinstance(attachment, dict) and 'leaves' in attachment and 'name' in attachment:
            items.append(attachment)
        else:
            items.append({'name': None, 'leaves': attachment})
    for item in attachments or []:
        if isinstance(item, dict) and 'leaves' in item:
            items.append({'name': item.get('name'), 'leaves': item.get('leaves')})
        elif item:
            items.append({'name': None, 'leaves': item})
    return items


def _section(source_name, payload):
    '''Split a flat leaf map or a partial case into the case sections.'''
    payload = payload or {}
    baseline = {}
    if isinstance(payload.get('baseline'), dict) and not is_leaf(payload.get('baseline')):
        baseline.update(_clean(source_name, payload['baseline']))
    nested = set(_NESTED)
    if is_leaf(payload.get('debt')):
        nested.discard('debt')
    flat = {key: value for key, value in payload.items() if key not in nested}
    baseline.update(_clean(source_name, flat))
    market = _clean(source_name, payload.get('market') or {})
    shares = baseline.pop('shares', None)
    price = baseline.pop('price', None)
    if shares and 'shares' not in market:
        market['shares'] = shares
    if price and 'price' not in market:
        market['price'] = price
    rates = _clean(source_name, payload.get('rates') or {})
    for name in _RATE_FIELDS:
        if name in baseline and name not in rates:
            rates[name] = baseline.pop(name)
    debt_node = payload.get('debt')
    debt = {} if is_leaf(debt_node) else _clean(source_name, debt_node or {})
    distribution = _clean(source_name, payload.get('distribution') or {})
    forecast = _clean_forecast(source_name, payload.get('forecast') or {})
    for name in _FORECAST_FIELDS:
        if name in baseline and name not in forecast and name not in FIELDS:
            forecast[name] = baseline.pop(name)
    return {
        'baseline': baseline,
        'market': market,
        'rates': rates,
        'forecast': forecast,
        'debt': debt,
        'distribution': distribution,
    }


def _clean_forecast(source_name, forecast):
    cleaned = {}
    for key, node in forecast.items():
        if key == 'margin_path' and isinstance(node, dict) and 'value' not in node:
            cleaned['margin_path'] = _clean(source_name, node)
        elif key == 'dwc' and isinstance(node, dict) and 'value' not in node:
            cleaned['dwc'] = _clean(source_name, node)
        elif isinstance(node, dict) and 'value' in node:
            node = copy.deepcopy(node)
            node.setdefault('source', source_name)
            cleaned[key] = node
    return cleaned


def _pick_sections(sections, order):
    picked = {'rates': {}, 'forecast': {}, 'debt': {}, 'distribution': {}}
    origins = {'rates': {}, 'forecast': {}, 'debt': {}, 'distribution': {}}
    for label in order:
        block = sections.get(label)
        if not block:
            continue
        for name in _RATE_FIELDS:
            _take(picked['rates'], origins['rates'], name, block['rates'].get(name), label)
        for name in _DEBT_FIELDS:
            _take(picked['debt'], origins['debt'], name, block['debt'].get(name), label)
        for name in _DISTRIBUTION_FIELDS:
            _take(picked['distribution'], origins['distribution'], name, block['distribution'].get(name), label)
        for name in _FORECAST_FIELDS:
            _take(picked['forecast'], origins['forecast'], name, block['forecast'].get(name), label)
        margin = block['forecast'].get('margin_path') or {}
        dest = picked['forecast'].setdefault('margin_path', {})
        dest_from = origins['forecast'].setdefault('margin_path', {})
        for name in ('me', 'mc', 'gsnext'):
            _take(dest, dest_from, name, margin.get(name), label)
        dwc = block['forecast'].get('dwc') or {}
        if dwc:
            dest = picked['forecast'].setdefault('dwc', {})
            dest_from = origins['forecast'].setdefault('dwc', {})
            for name in ('mode', 'values'):
                _take(dest, dest_from, name, dwc.get(name), label)
    if not picked['forecast'].get('margin_path'):
        picked['forecast'].pop('margin_path', None)
    if not picked['forecast'].get('dwc'):
        picked['forecast'].pop('dwc', None)
    picked['_from'] = origins
    return picked


def _take(dest, origins, name, node, label):
    if name in dest or not isinstance(node, dict) or node.get('value') is None:
        return
    dest[name] = node
    origins[name] = label


def _section_sources(picked):
    origins = picked.pop('_from', {})
    found = {}
    for section, fields in origins.items():
        for name, label in fields.items():
            if isinstance(label, dict):
                for child, child_label in label.items():
                    found['%s.%s.%s' % (section, name, child)] = child_label
            else:
                found['%s.%s' % (section, name)] = label
    return found


def _first_price(sections, feeds, currency, skipped, order):
    for label in order:
        node = (sections.get(label) or {}).get('market', {}).get('price')
        if node and node.get('value') is not None:
            if _same_currency(node, currency):
                return node, label
            skipped.append({
                'source': label, 'field': 'price',
                'reason': 'in %s, but the baseline is in %s' % (node.get('currency'), currency),
            })
    yahoo = _yahoo_price(feeds.get('yahoo'), currency, skipped)
    if yahoo is not None:
        return yahoo, 'yahoo'
    return None, None


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


def _reporting_currency(candidates, order):
    for name in order:
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
    source = str(period_source).split(':')[0]
    if source not in SOURCES:
        source = 'assumption'
    return {
        'value': period,
        'source': source,
        'evidence': 'period_end of the %s fiscal-year figures' % period_source,
        'period_end': period,
        'basis': 'fiscal_year',
    }


def _first_shares(sections, feeds, order):
    for label in order:
        if label in sections:
            node = sections[label]['market'].get('shares')
        else:
            node = (feeds.get(label) or {}).get('shares')
        if node and node.get('value') is not None:
            return node, label
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

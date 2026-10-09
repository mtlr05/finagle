'''Map the latest 10-K into case leaves.

``last_10k`` is called as it stands. Share count is read from the same
company-facts filing and returned beside the dollar baseline, because the
company constructor takes shares itself and would ignore a shares column
inside the financials dict.
'''

from datetime import datetime

from finagle.sec import (
    FACTS_URL,
    TICKER_URL,
    _cik_for_ticker,
    _get_json,
    last_10k,
)

_SHARE_TAG = 'EntityCommonStockSharesOutstanding'
# Cover-page share counts are dated after the fiscal period, often near the filing.
_COVER_WINDOW_DAYS = 150


def get_10k_baseline(ticker, user_agent, scale=1_000_000):
    '''Year-0 leaves from the latest 10-K, plus shares outstanding.

    Dollar amounts and the share count are divided by ``scale`` (one million
    by default). ``baseline`` matches the keys ``last_10k`` resolved. Debt is
    the baseline year only.
    '''
    baseline = last_10k(ticker, user_agent, scale=scale)
    leaves = leaves_from_baseline(baseline)
    period_end = None
    if 'date' in leaves:
        period_end = leaves['date']['value']
    cik = _cik_for_ticker(_get_json(TICKER_URL, user_agent), ticker)
    facts = _get_json(FACTS_URL.format(cik=cik), user_agent)
    shares = shares_from_facts(facts, period_end, scale=scale)
    return {
        'ok': True,
        'ticker': str(ticker).strip(),
        'period_end': period_end,
        'scale': scale,
        'units': 'millions of dollars; shares in millions',
        'baseline': leaves,
        'shares': shares,
        'not_in_10k': [
            'dwc', 'nol', 'noa', 're', 'rd', 't', 'gt', 'roict', 'price', 'dividend',
        ],
        'note': (
            'Figures are the latest 10-K fiscal year, not a newer trailing '
            'twelve months. List fields are year 0 only. Copy the leaves you '
            'accept into the case; this does not load a company.'
        ),
    }


def leaves_from_baseline(baseline):
    '''Turn a ``last_10k`` Baseline into provenance leaves.

    One-element lists become scalars so they match year-0 case fields.
    Missing facts stay in the result with a null value and the helper's note.
    '''
    leaves = {}
    frame = baseline.provenance
    financials = baseline.financials or {}
    for row in frame.to_dict(orient='records'):
        key = row['key']
        if key in financials:
            value = financials[key]
            if isinstance(value, list) and len(value) == 1:
                value = value[0]
        else:
            value = None
        tags = row.get('tags') or ''
        note = row.get('note') or ''
        evidence = tags or note or 'latest 10-K'
        leaf = {
            'value': value,
            'source': '10k',
            'evidence': evidence,
            'period_end': row.get('period end'),
        }
        if note:
            leaf['note'] = note
        leaves[key] = leaf
    return leaves


def shares_from_facts(facts, period_end, scale=1_000_000):
    '''Shares outstanding, in millions, from ``dei:EntityCommonStockSharesOutstanding``.

    Prefer an undimensioned 10-K instant on ``period_end``. Otherwise use the
    cover-page count dated within 150 days after that period. A dimensioned
    fact is used only when no total is present, and the note says it may be
    one class.
    '''
    missing = {
        'value': None,
        'source': '10k',
        'evidence': 'dei:' + _SHARE_TAG,
        'period_end': period_end,
        'note': _SHARE_TAG + ' was not found in the latest 10-K.',
    }
    if not scale:
        raise ValueError('scale must be non-zero')
    dei = (facts or {}).get('facts', {}).get('dei', {})
    concept = dei.get(_SHARE_TAG) or {}
    rows = []
    for unit_rows in (concept.get('units') or {}).values():
        for fact in unit_rows:
            if fact.get('form') == '10-K' and not fact.get('start'):
                rows.append(fact)
    if not rows:
        return missing

    on_period = [fact for fact in rows if fact.get('end') == period_end]
    pool = on_period or _cover_facts(rows, period_end)
    if not pool:
        return missing
    undimensioned = [fact for fact in pool if not _dimensioned(fact)]
    chosen_pool = undimensioned or pool
    chosen = max(chosen_pool, key=lambda fact: (fact.get('filed') or '', fact.get('accn') or ''))
    raw = chosen.get('val')
    if raw is None:
        return missing
    note_parts = []
    if chosen.get('end') != period_end:
        note_parts.append(
            'Share count is as of %s, the 10-K cover date, not the fiscal period end.'
            % chosen.get('end')
        )
    if _dimensioned(chosen):
        note_parts.append('This fact has dimensions; it may be one share class rather than the total.')
    leaf = {
        'value': raw / scale,
        'source': '10k',
        'evidence': 'dei:' + _SHARE_TAG,
        'period_end': chosen.get('end') or period_end,
    }
    if note_parts:
        leaf['note'] = ' '.join(note_parts)
    return leaf


def _cover_facts(rows, period_end):
    found = []
    for fact in rows:
        days = _days_after(period_end, fact.get('end'))
        if days is not None and 0 < days <= _COVER_WINDOW_DAYS:
            found.append(fact)
    return found


def _dimensioned(fact):
    for key in ('segment', 'dimensions'):
        if fact.get(key):
            return True
    return False


def _days_after(period_end, end):
    if not period_end or not end:
        return None
    try:
        start = datetime.strptime(period_end, '%Y-%m-%d')
        stop = datetime.strptime(end, '%Y-%m-%d')
    except (TypeError, ValueError):
        return None
    return (stop - start).days

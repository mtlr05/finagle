'''Check a case before it is compiled.

Errors and missing fields stop a run. Warnings are returned with the result
and do not stop it.
'''

import re

from finagle.agent.schema import (
    ACQUISITION_DISCLOSURES,
    BASES,
    DEBT_MODES,
    DISTRIBUTION_MODES,
    DWC_MODES,
    EBITDA_BASES,
    EBITDA_DEFINITIONS,
    PATHS,
    PRICE_PATHS,
    SOURCES,
    is_leaf,
    iter_leaves,
    unwrap,
)

_ABSENT = object()

# Market cap / EBITDA outside this band usually means the dollars or the
# share count were not entered in millions.
_RATIO_LOW = 0.05
_RATIO_HIGH = 500

# Average EBITDA growth above this, with no acquisitions, usually means the
# path includes EBITDA the case never pays for.
_GROWTH_WARN = 0.10

_ACQUISITION_WORDING = re.compile(
    r'acquired|inorganic|m\s*&\s*a|\bdeals?\b',
    re.IGNORECASE,
)

_TOP_LEVEL = {
    'path', 'ticker', 'year', 'market', 'rates', 'baseline', 'forecast',
    'dividend_per_share', 'debt', 'acquisitions', 'disposals', 'distribution',
    'earnings', 'fcfe', 'article',
}


def validate_case(case):
    '''Return ``errors``, ``warnings``, and ``missing`` for one case.'''
    errors = []
    warnings = []
    missing = []
    if not isinstance(case, dict):
        return {
            'errors': ['case must be an object'],
            'warnings': [],
            'missing': [],
        }

    _check_identity(case, errors, missing)
    _check_unknown(case, errors, warnings)
    if case.get('path') not in (None,) + PATHS:
        errors.append('path must be ebitda, earnings, or fcfe')

    path = case.get('path') or 'ebitda'
    if path == 'ebitda':
        _check_ebitda(case, errors, warnings, missing)
    elif path == 'earnings':
        _check_earnings(case, errors, missing)
    elif path == 'fcfe':
        _check_fcfe(case, errors, missing)

    _check_periods(case, errors)
    _check_currency(case, errors)
    _check_stale_price(case, warnings)
    return {'errors': errors, 'warnings': warnings, 'missing': missing}


def _check_identity(case, errors, missing):
    if 'ticker' not in case or not str(case.get('ticker') or '').strip():
        missing.append('ticker')
    elif not isinstance(case.get('ticker'), str):
        errors.append('ticker must be a string')
    year = _whole_year(case.get('year', _ABSENT))
    if 'year' not in case:
        missing.append('year')
    elif year is None:
        errors.append('year must be an integer of 1 or more')


def _check_unknown(case, errors, warnings):
    if 'debt_target' in case or 'fcf_to_debt' in case or isinstance(case.get('debt'), list):
        errors.append('only one debt target is allowed')
    for key in case:
        if key not in _TOP_LEVEL:
            warnings.append('unknown field %s is ignored' % key)


def _check_ebitda(case, errors, warnings, missing):
    year = _whole_year(case.get('year', _ABSENT))
    _read(case, errors, missing, True, 'market', 'shares')
    _read(case, errors, missing, False, 'market', 'price')
    for name in ('re', 'rd', 't', 'gt', 'roict'):
        _read(case, errors, missing, True, 'rates', name)
    _read(case, errors, missing, False, 'rates', 'te')
    for name in (
        'date', 'ebitda', 'capex', 'da', 'tax', 'interest', 'sbc', 'debt',
        'cash', 'nol', 'noa',
    ):
        _read(case, errors, missing, True, 'baseline', name)
    _read(case, errors, missing, False, 'baseline', 'ebitda_definition')
    _read(case, errors, missing, False, 'baseline', 'revenue')
    _read(case, errors, missing, True, 'forecast', 'ebitda_basis')
    _read_acquisition_disclosure(case, errors, missing)
    _read(case, errors, missing, False, 'forecast', 'ebitda_growth')
    _read(case, errors, missing, False, 'forecast', 'ebitda')
    _read(case, errors, missing, False, 'forecast', 'capex')
    _read(case, errors, missing, False, 'forecast', 'sbc')
    _read(case, errors, missing, False, 'forecast', 'sbc_rate_terminal')
    _read_margin_block(case, errors, missing)
    _read(case, errors, missing, True, 'forecast', 'dwc', 'mode')
    _read(case, errors, missing, False, 'forecast', 'dwc', 'values')
    _read(case, errors, missing, False, 'dividend_per_share')
    _read(case, errors, missing, True, 'debt', 'mode')
    _read(case, errors, missing, False, 'debt', 'leverage')
    _read(case, errors, missing, False, 'debt', 'start_year')
    _read(case, errors, missing, False, 'debt', 'values')
    _read(case, errors, missing, True, 'distribution', 'mode')
    _read(case, errors, missing, False, 'distribution', 'schedule')
    _read(case, errors, missing, False, 'distribution', 'price')
    _read(case, errors, missing, False, 'distribution', 'price_path')
    _read(case, errors, missing, False, 'article', 'ebitda_total')
    _read_article(case, errors, missing)
    _read_deals(case, errors, missing)
    _reject_other_paths(case, 'ebitda', errors)

    data = unwrap(case)
    _check_definition(data, errors, warnings)
    _check_rates(data, errors)
    _check_units(data, warnings)
    if year is None:
        return
    _check_ebitda_forecast(data, year, errors, warnings, missing)
    _check_acquisition_consistency(case, data, year, errors, warnings, missing)
    _check_debt_intent(data, year, errors, missing)
    _check_distribution_intent(data, errors, missing)
    _check_series(data, year, errors)


def _check_earnings(case, errors, missing):
    _read(case, errors, missing, True, 'rates', 're')
    _read(case, errors, missing, True, 'rates', 'gt')
    _read(case, errors, missing, True, 'baseline', 'date')
    for name in ('e', 'payout', 'gf', 'roe'):
        _read(case, errors, missing, True, 'earnings', name)
    _read(case, errors, missing, False, 'market', 'shares')
    _reject_other_paths(case, 'earnings', errors)
    data = unwrap(case)
    _check_rates(data, errors)
    year = _whole_year(case.get('year', _ABSENT))
    growth = (data.get('earnings') or {}).get('gf')
    if year is not None:
        _check_length('earnings.gf', growth, year, errors, exact=False)


def _check_fcfe(case, errors, missing):
    _read(case, errors, missing, True, 'rates', 're')
    _read(case, errors, missing, True, 'rates', 'gt')
    _read(case, errors, missing, True, 'baseline', 'date')
    _read(case, errors, missing, True, 'fcfe')
    _read(case, errors, missing, False, 'market', 'shares')
    _reject_other_paths(case, 'fcfe', errors)
    data = unwrap(case)
    _check_rates(data, errors)
    year = _whole_year(case.get('year', _ABSENT))
    if year is not None:
        _check_length('fcfe', data.get('fcfe'), year, errors, exact=True)


def _check_definition(data, errors, warnings):
    definition = (data.get('baseline') or {}).get('ebitda_definition')
    if definition is None:
        warnings.append(
            'baseline.ebitda_definition is missing. EBITDA must be before '
            'stock-based compensation, because the model subtracts sbc.'
        )
        return
    if definition not in EBITDA_DEFINITIONS:
        errors.append('ebitda_definition must be before_sbc or after_sbc')
    elif definition == 'after_sbc':
        errors.append(
            'EBITDA must be before stock-based compensation. The model '
            'subtracts sbc, and the 10-K baseline already adds sbc back. '
            'Add sbc back to the article EBITDA and set ebitda_definition '
            'to before_sbc.'
        )


def _check_rates(data, errors):
    rates = data.get('rates') or {}
    re = rates.get('re')
    gt = rates.get('gt')
    if _is_number(re) and _is_number(gt) and re <= gt:
        errors.append('re must be greater than gt')


def _check_units(data, warnings):
    market = data.get('market') or {}
    price = market.get('price')
    shares = market.get('shares')
    ebitda = (data.get('baseline') or {}).get('ebitda')
    if isinstance(ebitda, list) and ebitda:
        ebitda = ebitda[0]
    if not (_is_number(price) and _is_number(shares) and _is_number(ebitda)):
        return
    if price <= 0 or shares <= 0 or ebitda <= 0:
        return
    ratio = price * shares / ebitda
    if ratio < _RATIO_LOW or ratio > _RATIO_HIGH:
        warnings.append(
            'market cap divided by EBITDA is %.4g. Money is in millions of '
            'the reporting currency and shares are in millions of shares. A '
            'ratio this far from a normal multiple usually means raw amounts, '
            'billions, or a price in another currency were entered.' % ratio
        )


def _check_ebitda_forecast(data, year, errors, warnings, missing):
    forecast = data.get('forecast') or {}
    explicit = forecast.get('ebitda')
    growth = forecast.get('ebitda_growth')
    if explicit is not None and growth is not None:
        errors.append('forecast.ebitda and forecast.ebitda_growth cannot both be set')
    if explicit is None and growth is None:
        missing.append('forecast.ebitda_growth or forecast.ebitda')
    if explicit is not None:
        _check_length('forecast.ebitda', explicit, year, errors, exact=True)
    if growth is not None:
        _check_length('forecast.ebitda_growth', growth, year, errors, exact=False)
    _check_length('forecast.capex', forecast.get('capex'), year, errors, exact=False)
    _check_length('forecast.sbc', forecast.get('sbc'), year, errors, exact=False)
    sbc = forecast.get('sbc')
    sbc_len = len(sbc) if isinstance(sbc, list) else (1 if sbc is not None else None)
    if forecast.get('sbc_rate_terminal') is not None and sbc_len == year + 1:
        warnings.append(
            'sbc_rate_terminal is ignored when forecast.sbc already covers every year'
        )

    margin = forecast.get('margin_path')
    if isinstance(margin, dict):
        if growth is None:
            errors.append('margin_path requires ebitda_growth')
        for name in ('me', 'mc'):
            if _is_number(margin.get(name)) and margin[name] == 0:
                errors.append(
                    'margin_path %s must be non-zero; the model treats 0 as unset' % name
                )

    dwc = forecast.get('dwc') or {}
    mode = dwc.get('mode')
    if mode is not None and mode not in DWC_MODES:
        errors.append('forecast.dwc.mode must be zero or path')
    if mode == 'path':
        if dwc.get('values') is None:
            missing.append('forecast.dwc.values')
        else:
            _check_length('forecast.dwc.values', dwc.get('values'), year, errors, exact=True)


def _read_margin_block(case, errors, missing):
    margin = _dig(case, 'forecast', 'margin_path')
    if margin is _ABSENT or margin is None:
        return
    if not isinstance(margin, dict) or is_leaf(margin):
        errors.append('forecast.margin_path must be an object')
        return
    for name in ('me', 'mc', 'gsnext'):
        _read_at(margin, errors, missing, True, 'forecast.margin_path.%s' % name, name)


def _check_debt_intent(data, year, errors, missing):
    debt = data.get('debt') or {}
    mode = debt.get('mode')
    if mode is not None and mode not in DEBT_MODES:
        errors.append('debt.mode must be hold, path, or target')
    if mode == 'path':
        if debt.get('values') is None:
            missing.append('debt.values')
        else:
            _check_length('debt.values', debt.get('values'), year, errors, exact=True)
    if mode == 'target':
        if debt.get('leverage') is None:
            missing.append('debt.leverage')
        if debt.get('start_year') is None:
            missing.append('debt.start_year')
        elif not _year_in_horizon(debt.get('start_year'), year, low=1):
            errors.append('debt.start_year is outside the forecast horizon')
        if debt.get('values') is not None:
            _check_length('debt.values', debt.get('values'), year, errors, exact=True)


def _check_distribution_intent(data, errors, missing):
    distribution = data.get('distribution') or {}
    mode = distribution.get('mode')
    if mode is not None and mode not in DISTRIBUTION_MODES:
        errors.append('distribution.mode must be none, max_buybacks, buyback_schedule, or retain')
    price_path = distribution.get('price_path')
    if price_path is not None and price_path not in PRICE_PATHS:
        errors.append('distribution.price_path must be constant or proportional')
    needs_price = mode in ('max_buybacks', 'buyback_schedule')
    market_price = (data.get('market') or {}).get('price')
    if needs_price and distribution.get('price') is None and market_price is None:
        missing.append('distribution.price')
    if mode == 'buyback_schedule':
        if distribution.get('schedule') is None:
            missing.append('distribution.schedule')
        else:
            year = _whole_year((data.get('year')))
            if year is not None:
                _check_length(
                    'distribution.schedule', distribution.get('schedule'),
                    year, errors, exact=False,
                )


def _check_series(data, year, errors):
    deals = data.get('acquisitions') or []
    if not isinstance(deals, list):
        errors.append('acquisitions must be a list')
        return
    for index, deal in enumerate(deals):
        if not isinstance(deal, dict):
            errors.append('acquisitions[%d] must be an object' % index)
            continue
        deal_year = deal.get('year')
        if _is_last_explicit_year(deal_year, year):
            errors.append(
                'acquisitions[%d].year is the last forecast year, so the acquired '
                'EBITDA would fall past the horizon' % index
            )
        elif not _year_in_horizon(deal_year, year, low=0):
            errors.append('acquisitions[%d].year is outside the forecast horizon' % index)
    for index, disposal in enumerate(data.get('disposals') or []):
        if isinstance(disposal, dict) and not _year_in_horizon(disposal.get('year'), year, low=0):
            errors.append('disposals[%d].year is outside the forecast horizon' % index)


def _read_deals(case, errors, missing):
    deals = case.get('acquisitions')
    if deals is None:
        return
    if not isinstance(deals, list):
        errors.append('acquisitions must be a list')
        return
    size_fields = ('ebitda_frac', 'acquired_ebitda', 'deal_value')
    required = ('year', 'multiple', 'leverage', 'next_growth', 'capex_frac', 'pay_from_cash')
    for index, deal in enumerate(deals):
        prefix = 'acquisitions[%d]' % index
        if not isinstance(deal, dict):
            errors.append('%s must be an object' % prefix)
            continue
        for name in required:
            _read_at(deal, errors, missing, True, '%s.%s' % (prefix, name), name)
        present = []
        for name in size_fields:
            value = _read_at(deal, errors, missing, False, '%s.%s' % (prefix, name), name)
            if value is not None:
                present.append(name)
        if len(present) == 0:
            missing.append('%s size (ebitda_frac, acquired_ebitda, or deal_value)' % prefix)
        elif len(present) > 1:
            errors.append(
                '%s size must be only one of ebitda_frac, acquired_ebitda, or deal_value' % prefix
            )
        pay = deal.get('pay_from_cash')
        if is_leaf(pay) and not isinstance(pay.get('value'), bool):
            errors.append('%s.pay_from_cash must be true or false' % prefix)

    disposals = case.get('disposals')
    if disposals is None:
        return
    if not isinstance(disposals, list):
        errors.append('disposals must be a list')
        return
    for index, disposal in enumerate(disposals):
        prefix = 'disposals[%d]' % index
        if not isinstance(disposal, dict):
            errors.append('%s must be an object' % prefix)
            continue
        _read_at(disposal, errors, missing, True, '%s.year' % prefix, 'year')
        _read_at(disposal, errors, missing, True, '%s.amount' % prefix, 'amount')
        _read_at(disposal, errors, missing, False, '%s.tax' % prefix, 'tax')


def _read_acquisition_disclosure(case, errors, missing):
    node = _dig(case, 'forecast', 'acquisition_disclosure')
    if node is _ABSENT or node is None:
        missing.append('forecast.acquisition_disclosure')
        return
    if not isinstance(node, dict) or is_leaf(node):
        errors.append('forecast.acquisition_disclosure must be an object')
        return
    _read_at(node, errors, missing, True, 'forecast.acquisition_disclosure.status', 'status')
    _read_at(node, errors, missing, False, 'forecast.acquisition_disclosure.deal_spend', 'deal_spend')


def _check_acquisition_consistency(case, data, year, errors, warnings, missing):
    '''Acquired EBITDA and the price paid for it have to be the same story.'''
    forecast = data.get('forecast') or {}
    basis = forecast.get('ebitda_basis')
    if basis is not None and basis not in EBITDA_BASES:
        errors.append('forecast.ebitda_basis must be organic or total')

    disclosure = forecast.get('acquisition_disclosure') or {}
    if not isinstance(disclosure, dict):
        disclosure = {}
    status = disclosure.get('status')
    if status is not None and status not in ACQUISITION_DISCLOSURES:
        errors.append(
            'forecast.acquisition_disclosure.status must be future_deals or none'
        )
    spend = disclosure.get('deal_spend')
    if status == 'future_deals':
        if spend is None:
            missing.append('forecast.acquisition_disclosure.deal_spend')
        elif not isinstance(spend, list) or not all(_is_number(item) for item in spend):
            errors.append(
                'forecast.acquisition_disclosure.deal_spend must be a list of numbers'
            )
        else:
            _check_length(
                'forecast.acquisition_disclosure.deal_spend', spend, year, errors, exact=False,
            )

    deals = data.get('acquisitions') or []
    has_deals = isinstance(deals, list) and len(deals) > 0
    if basis == 'total' and has_deals:
        errors.append(
            'forecast.ebitda_basis is total and acquisitions are set, so acquired '
            'EBITDA would be counted twice'
        )
    if status == 'future_deals' and not has_deals:
        errors.append(
            'acquisition_disclosure is future_deals but there are no acquisitions'
        )
    if status == 'none' and has_deals:
        errors.append(
            'acquisition_disclosure is none but acquisitions are set'
        )
    if has_deals:
        return

    average = _average_ebitda_growth(forecast)
    if average is not None and average > _GROWTH_WARN:
        warnings.append(
            'average EBITDA growth is %.1f%% and there are no acquisitions. '
            'If that growth includes acquired EBITDA, use an organic path and '
            'book the purchase price with acquisitions, or MnA stays zero.'
            % (average * 100)
        )
    if _forecast_mentions_acquisitions(case):
        warnings.append(
            'forecast evidence mentions acquisitions but the case has none, so MnA stays zero'
        )


def _average_ebitda_growth(forecast):
    growth = forecast.get('ebitda_growth')
    if isinstance(growth, bool):
        return None
    if isinstance(growth, list) and growth and all(_is_number(item) for item in growth):
        return sum(growth) / float(len(growth))
    if _is_number(growth):
        return float(growth)
    path = forecast.get('ebitda')
    if not isinstance(path, list) or len(path) < 2:
        return None
    rates = []
    for previous, following in zip(path, path[1:]):
        if _is_number(previous) and _is_number(following) and previous > 0:
            rates.append((following - previous) / float(previous))
    if not rates:
        return None
    return sum(rates) / float(len(rates))


def _forecast_mentions_acquisitions(case):
    forecast = case.get('forecast')
    if not isinstance(forecast, dict):
        return False
    for path, node in iter_leaves(forecast, 'forecast'):
        if path.startswith('forecast.acquisition_disclosure'):
            continue
        evidence = node.get('evidence') or ''
        if _ACQUISITION_WORDING.search(evidence):
            return True
    return False


def _read_article(case, errors, missing):
    article = case.get('article')
    if article is None:
        return
    if not isinstance(article, dict):
        errors.append('article must be an object')
        return
    for name in ('title', 'url', 'published', 'price_target'):
        _read_at(article, errors, missing, False, 'article.%s' % name, name)


def _reject_other_paths(case, path, errors):
    if path != 'earnings' and case.get('earnings') is not None:
        errors.append('earnings is only used when path is earnings')
    if path != 'fcfe' and case.get('fcfe') is not None:
        errors.append('fcfe is only used when path is fcfe')


def _check_periods(case, errors):
    '''Reported year-0 figures are the last fiscal year and all end on baseline.date.'''
    baseline = case.get('baseline')
    if not isinstance(baseline, dict):
        return
    date = baseline.get('date')
    year_end = str(date['value']) if is_leaf(date) and date.get('value') else None
    ends = set()
    for key, node in baseline.items():
        if key == 'ebitda_definition' or not is_leaf(node):
            continue
        path = 'baseline.%s' % key
        reported = node.get('source') != 'assumption'
        basis = node.get('basis')
        period_end = node.get('period_end')
        if basis is not None and basis not in BASES:
            errors.append('%s basis must be fiscal_year, ttm, or quarter' % path)
        elif basis in ('ttm', 'quarter'):
            errors.append(
                '%s is %s. Year 0 must be the last completed fiscal year, because '
                'EBITDA growth is measured from it.' % (path, basis)
            )
        elif reported and basis is None:
            errors.append('%s needs basis fiscal_year' % path)
        if reported and not period_end:
            errors.append('%s needs period_end' % path)
        if period_end:
            ends.add(str(period_end))
            if year_end and str(period_end) != year_end:
                errors.append(
                    '%s is for %s, not the baseline year ending %s'
                    % (path, period_end, year_end)
                )
    if not year_end and len(ends) > 1:
        errors.append(
            'year-0 figures come from different period ends: ' + ', '.join(sorted(ends))
        )


def _check_currency(case, errors):
    '''One reporting currency for year 0, and a price in that currency.'''
    baseline = case.get('baseline')
    currencies = set()
    if isinstance(baseline, dict):
        for node in baseline.values():
            if is_leaf(node) and node.get('currency'):
                currencies.add(str(node['currency']).upper())
    if len(currencies) > 1:
        errors.append(
            'baseline figures are in more than one currency: ' + ', '.join(sorted(currencies))
        )
        return
    if not currencies:
        return
    reporting = next(iter(currencies))
    for path in (('market', 'price'), ('distribution', 'price')):
        node = _dig(case, *path)
        if is_leaf(node) and node.get('currency'):
            currency = str(node['currency']).upper()
            if currency != reporting:
                errors.append(
                    '%s is in %s but the financials are in %s. Use the price converted '
                    'to %s.' % ('.'.join(path), currency, reporting, reporting)
                )


def _check_stale_price(case, warnings):
    price = _dig(case, 'market', 'price')
    published = _dig(case, 'article', 'published')
    if not is_leaf(price) or not is_leaf(published):
        return
    as_of = price.get('as_of')
    published_on = published.get('value')
    if as_of and published_on and str(as_of) < str(published_on):
        warnings.append('article price is older than the article date')


def _read(case, errors, missing, required, *keys):
    path = '.'.join(keys)
    node = _dig(case, *keys)
    return _consume(node, path, errors, missing, required)


def _read_at(parent, errors, missing, required, path, key):
    if not isinstance(parent, dict) or key not in parent:
        node = _ABSENT
    else:
        node = parent[key]
    return _consume(node, path, errors, missing, required)


def _consume(node, path, errors, missing, required):
    if node is _ABSENT or node is None:
        if required:
            missing.append(path)
        return None
    if not is_leaf(node):
        errors.append('%s is missing provenance' % path)
        return None
    _check_leaf_meta(path, node, errors)
    if required and node.get('value') is None:
        missing.append(path)
        return None
    return node.get('value')


def _check_leaf_meta(path, node, errors):
    if node.get('source') not in SOURCES:
        errors.append(
            '%s source must be article, attachment, 10k, yahoo, or assumption' % path
        )
    evidence = node.get('evidence')
    if not isinstance(evidence, str) or not evidence.strip():
        errors.append('%s is missing provenance' % path)


def _dig(node, *keys):
    current = node
    for key in keys:
        if not isinstance(current, dict) or key not in current:
            return _ABSENT
        current = current[key]
    return current


def _check_length(path, value, year, errors, exact):
    if value is None or isinstance(value, bool) or not isinstance(value, (list, float, int)):
        return
    length = len(value) if isinstance(value, list) else 1
    if length > year + 1:
        errors.append(
            '%s is longer than the forecast horizon (year + 1 = %d)' % (path, year + 1)
        )
    elif exact and length != year + 1:
        errors.append(
            '%s is short: expected %d values, got %d' % (path, year + 1, length)
        )


def _is_last_explicit_year(value, year):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return False
    if isinstance(value, float) and not value.is_integer():
        return False
    return int(value) == year


def _year_in_horizon(value, year, low):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return False
    if isinstance(value, float) and not value.is_integer():
        return False
    return low <= int(value) <= year


def _whole_year(value):
    if value is _ABSENT or isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    if isinstance(value, float) and not value.is_integer():
        return None
    number = int(value)
    if number < 1:
        return None
    return number


def _is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)

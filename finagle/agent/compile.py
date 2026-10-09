'''Turn a validated case into constructor arguments and an ordered method list.

The company class replays capital actions in a fixed order and keeps only the
last debt target and the last distribution. The case therefore carries intents,
and this module is the only place those intents become method calls.
'''

from finagle.agent.schema import unwrap


class CompileError(ValueError):
    '''The case cannot be compiled into company method calls.'''


def compile_case(case):
    '''Return financials, constructor kwargs, and steps for ``run_case``.

    Steps that need the financials dict set ``pass_financials``. Acquisition
    steps that still need a fraction set ``resolve_frac``; the runner fills
    ``ebitda_frac`` from EBITDA in the deal year just before that deal.
    '''
    if not isinstance(case, dict):
        raise CompileError('case must be an object')
    data = unwrap(case)
    path = data.get('path') or 'ebitda'
    if path == 'earnings':
        return _compile_earnings(data)
    if path == 'fcfe':
        return _compile_fcfe(data)
    if path != 'ebitda':
        raise CompileError('path must be ebitda, earnings, or fcfe')
    return _compile_ebitda(data)


def _require(data, *keys):
    current = data
    walked = []
    for key in keys:
        walked.append(key)
        if not isinstance(current, dict) or key not in current or current[key] is None:
            raise CompileError('case is missing %s' % '.'.join(walked))
        current = current[key]
    return current


def _as_list(value):
    if isinstance(value, list):
        return list(value)
    return [value]


def _compile_ebitda(data):
    year = int(_require(data, 'year'))
    baseline = _require(data, 'baseline')
    forecast = data.get('forecast') or {}
    if not isinstance(forecast, dict):
        raise CompileError('forecast must be an object')

    explicit = forecast.get('ebitda')
    growth = forecast.get('ebitda_growth')
    if explicit is not None and growth is not None:
        raise CompileError('forecast.ebitda and forecast.ebitda_growth cannot both be set')
    if explicit is None and growth is None:
        raise CompileError('case is missing forecast.ebitda_growth or forecast.ebitda')

    if explicit is not None:
        ebitda_series = _as_list(explicit)
        forecast_ebitda = False
    else:
        ebitda_series = [_require(baseline, 'ebitda')]
        forecast_ebitda = True

    if forecast.get('capex') is not None:
        capex_series = _as_list(forecast['capex'])
    else:
        capex_series = [_require(baseline, 'capex')]
    forecast_capex = len(capex_series) < year + 1

    sbc_rate = forecast.get('sbc_rate_terminal')
    if forecast.get('sbc') is not None:
        sbc_series = _as_list(forecast['sbc'])
    else:
        sbc_series = [_require(baseline, 'sbc')]
    forecast_sbc = len(sbc_series) < year + 1

    financials = {
        'date': _require(baseline, 'date'),
        'revenue': [0],
        'ebitda': ebitda_series,
        'capex': capex_series,
        'sbc': sbc_series,
        'dwc': _dwc_series(forecast, year),
        'tax': _as_list(_require(baseline, 'tax')),
        'da': _as_list(_require(baseline, 'da')),
        'debt': _debt_series(data, baseline, year),
        'interest': _as_list(_require(baseline, 'interest')),
        'cash': _require(baseline, 'cash'),
        'nol': _require(baseline, 'nol'),
        'noa': _require(baseline, 'noa'),
    }

    steps = []
    if forecast_ebitda:
        args = {
            'ebitda_ttm': _require(baseline, 'ebitda'),
            'gf': growth if not isinstance(growth, list) else list(growth),
        }
        margin = forecast.get('margin_path')
        if margin:
            args['me'] = _require(margin, 'me')
            args['mc'] = _require(margin, 'mc')
            args['gsnext'] = _require(margin, 'gsnext')
        steps.append({
            'method': 'forecast_ebitda',
            'args': args,
            'pass_financials': True,
        })
    if forecast_capex:
        steps.append({
            'method': 'forecast_capex',
            'args': {'capex_f': list(capex_series)},
            'pass_financials': True,
        })
    if forecast_sbc:
        args = {'sbc_f': list(sbc_series)}
        if sbc_rate is not None:
            args['sbc_rate_t'] = sbc_rate
        steps.append({
            'method': 'forecast_sbc',
            'args': args,
            'pass_financials': True,
        })

    load_in_constructor = not steps
    if not load_in_constructor:
        steps.append({
            'method': 'load_financials',
            'args': {},
            'pass_financials': True,
        })

    steps.append({'method': 'fcf_from_ebitda', 'args': {}})
    steps.extend(_acquisition_steps(data.get('acquisitions') or []))
    steps.extend(_disposal_steps(data.get('disposals') or []))

    debt = _require(data, 'debt')
    if debt.get('mode') == 'target':
        steps.append({
            'method': 'fcf_to_debt',
            'args': {
                'leverage': _require(debt, 'leverage'),
                'year_d': int(_require(debt, 'start_year')),
            },
        })

    steps.extend(_distribution_steps(data.get('distribution') or {'mode': 'none'}, data.get('market') or {}))
    steps.append({'method': 'value', 'args': {}})

    constructor = _common_constructor(data)
    constructor['rd'] = _require(data, 'rates', 'rd')
    constructor['t'] = _require(data, 'rates', 't')
    constructor['roict'] = _require(data, 'rates', 'roict')
    te = (data.get('rates') or {}).get('te')
    if te is not None:
        constructor['te'] = te
    if load_in_constructor:
        constructor['financials'] = financials

    return {
        'path': 'ebitda',
        'financials': financials,
        'constructor': constructor,
        'steps': steps,
    }


def _dwc_series(forecast, year):
    dwc = forecast.get('dwc') or {}
    mode = dwc.get('mode')
    if mode == 'zero':
        return [0] * (year + 1)
    if mode == 'path':
        return _as_list(_require(dwc, 'values'))
    raise CompileError('forecast.dwc.mode must be zero or path')


def _debt_series(data, baseline, year):
    debt = _require(data, 'debt')
    mode = debt.get('mode')
    values = debt.get('values')
    if mode == 'hold':
        return [_require(baseline, 'debt')] * (year + 1)
    if mode == 'path':
        return _as_list(_require(debt, 'values'))
    if mode == 'target':
        if values is None:
            return [_require(baseline, 'debt')] * (year + 1)
        return _as_list(values)
    raise CompileError('debt.mode must be hold, path, or target')


def _acquisition_steps(deals):
    steps = []
    for deal in deals:
        args = {
            'adjust_cash': bool(_require(deal, 'pay_from_cash')),
            'year_a': int(_require(deal, 'year')),
            'multiple': _require(deal, 'multiple'),
            'leverage': _require(deal, 'leverage'),
            'gnext': _require(deal, 'next_growth'),
            'cap_frac': _require(deal, 'capex_frac'),
        }
        step = {'method': 'fcf_to_acquire', 'args': args}
        if deal.get('ebitda_frac') is not None:
            args['ebitda_frac'] = deal['ebitda_frac']
        elif deal.get('acquired_ebitda') is not None:
            step['resolve_frac'] = {
                'kind': 'acquired_ebitda',
                'acquired_ebitda': deal['acquired_ebitda'],
                'year': int(deal['year']),
            }
        elif deal.get('deal_value') is not None:
            step['resolve_frac'] = {
                'kind': 'deal_value',
                'deal_value': deal['deal_value'],
                'multiple': deal['multiple'],
                'year': int(deal['year']),
            }
        else:
            raise CompileError(
                'each acquisition needs ebitda_frac, acquired_ebitda, or deal_value'
            )
        steps.append(step)
    return steps


def _disposal_steps(disposals):
    steps = []
    for disposal in disposals:
        steps.append({
            'method': 'noa_to_dispose',
            'args': {
                'dnoa': _require(disposal, 'amount'),
                'tax': 0 if disposal.get('tax') is None else disposal['tax'],
                'year_dis': int(_require(disposal, 'year')),
            },
        })
    return steps


def _distribution_steps(distribution, market):
    mode = distribution.get('mode') or 'none'
    if mode == 'none':
        return []
    price = distribution.get('price')
    if price is None:
        price = market.get('price')
    price_path = distribution.get('price_path') or 'proportional'
    if mode == 'max_buybacks':
        return [{
            'method': 'fcf_to_buyback',
            'args': {'price': _required_price(price), 'dp': price_path},
        }]
    if mode == 'buyback_schedule':
        return [{
            'method': 'fcf_to_allocate',
            'args': {
                'price': _required_price(price),
                'dp': price_path,
                'buybacks': _as_list(_require(distribution, 'schedule')),
            },
        }]
    if mode == 'retain':
        return [{'method': 'fcf_to_bs', 'args': {}}]
    raise CompileError('distribution.mode is not recognized')


def _required_price(price):
    if price is None:
        raise CompileError('distribution.price is missing')
    return price


def _common_constructor(data):
    constructor = {
        'ticker': _require(data, 'ticker'),
        're': _require(data, 'rates', 're'),
        'gt': _require(data, 'rates', 'gt'),
        'year': int(_require(data, 'year')),
    }
    market = data.get('market') or {}
    if market.get('shares') is not None:
        constructor['shares'] = market['shares']
    if market.get('price') is not None:
        constructor['price'] = market['price']
    if data.get('dividend_per_share') is not None:
        constructor['dividend'] = data['dividend_per_share']
    return constructor


def _compile_earnings(data):
    earnings = _require(data, 'earnings')
    financials = {
        'date': _require(data, 'baseline', 'date'),
        'e': _require(earnings, 'e'),
    }
    constructor = _common_constructor(data)
    constructor['financials'] = financials
    steps = [
        {
            'method': 'fcf_from_earnings',
            'args': {
                'payout': _require(earnings, 'payout'),
                'gf': _require(earnings, 'gf'),
                'ROE': _require(earnings, 'roe'),
            },
        },
        {'method': 'value', 'args': {}},
    ]
    return {
        'path': 'earnings',
        'financials': financials,
        'constructor': constructor,
        'steps': steps,
    }


def _compile_fcfe(data):
    financials = {'date': _require(data, 'baseline', 'date')}
    constructor = _common_constructor(data)
    constructor['financials'] = financials
    constructor['fcfe'] = _as_list(_require(data, 'fcfe'))
    return {
        'path': 'fcfe',
        'financials': financials,
        'constructor': constructor,
        'steps': [{'method': 'value', 'args': {}}],
    }

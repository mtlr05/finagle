'''Run a case on the existing company class and return plain numbers.

The run happens in a temporary directory so the log file company writes does
not land next to a person's notebooks. display_fin is not called.
'''

import os
import tempfile
import threading

from finagle.agent.compile import CompileError, compile_case, normalize_assumed_paths
from finagle.agent.schema import provenance, unwrap
from finagle.agent.validate import validate_case
from finagle.company import company

_LOCK = threading.Lock()

_YEAR_COLUMNS = (
    'revenue', 'ebitda', 'capex', 'MnA', 'dDebt', 'fcf', 'fcfe', 'fcff', 'debt', 'cash',
    'cashBS', 'shares', 'price', 'dividend', 'equity', 'firm',
)

# Model EBITDA more than this far from the article's total, or MnA more than
# this far from the disclosed deal spend, is called out beside the value.
_EBITDA_GAP = 0.05
_MNA_GAP = 0.10

_SCALAR_ORDER = (
    'ticker', 'rd', 're', 't', 'te', 'shares', 'price', 'gt', 'roict',
    'year', 'dividend',
)
_SCALAR_DEFAULTS = {
    'ticker': 'REPLACE',
    'rd': None,
    're': None,
    't': None,
    'te': None,
    'shares': 1,
    'price': 0,
    'gt': 0,
    'roict': 0.15,
    'year': 6,
    'dividend': 0,
}
_FINANCIAL_KEYS = (
    'date', 'revenue', 'ebitda', 'capex', 'dwc', 'sbc', 'tax', 'da',
    'debt', 'interest', 'cash', 'nol', 'noa',
)
_FINANCIAL_COMMENTS = {
    'revenue': 'leave [0] if forecast_ebitda is commented; that call overwrites it',
    'ebitda': 'TTM only, or a full list of year + 1 values',
}
_ACQUIRE_STUB = (
    '# co.fcf_to_acquire(adjust_cash=False, year_a=%s, ebitda_frac=0.10, '
    'multiple=10, leverage=3, gnext=0.10, cap_frac=0.20)'
)
_DISTRIBUTION_STUBS = (
    ('fcf_to_allocate', "# co.fcf_to_allocate(price=price, dp='proportional', buybacks=None)"),
    ('fcf_to_buyback', "# co.fcf_to_buyback(price=price, dp='proportional')"),
    ('fcf_to_bs', '# co.fcf_to_bs()'),
)

_ARG_ORDER = {
    'forecast_ebitda': ('ebitda_ttm', 'gf', 'me', 'mc', 'gsnext'),
    'forecast_capex': ('capex_f', 'method'),
    'forecast_sbc': ('sbc_f', 'sbc_rate_t'),
    'fcf_from_earnings': ('payout', 'gf', 'ROE'),
    'fcf_to_debt': ('leverage', 'year_d'),
    'fcf_to_buyback': ('price', 'dp'),
    'fcf_to_allocate': ('price', 'dp', 'buybacks'),
    'fcf_to_acquire': (
        'adjust_cash', 'year_a', 'ebitda_frac', 'multiple', 'leverage',
        'gnext', 'cap_frac',
    ),
    'noa_to_dispose': ('dnoa', 'tax', 'year_dis'),
}


def run_case(case):
    '''Validate ``case``, run it, and return a JSON-ready result.

    Validation errors and missing fields skip the model. A log error from
    the model leaves the numbers in place and sets ``ok`` false.

    An assumed full capex or stock-based-compensation path, or one that
    repeats the year-0 amount, is cut back to year 0 before validation.
    The returned case is that copy, and the cut is a validation warning.
    '''
    path_warnings = []
    if isinstance(case, dict):
        case, path_warnings = normalize_assumed_paths(case)
    validation = validate_case(case)
    validation['warnings'].extend(path_warnings)
    if validation['errors'] or validation['missing']:
        return _blocked(case, validation)
    try:
        compiled = compile_case(case)
    except CompileError as exc:
        validation['errors'].append(str(exc))
        return _blocked(case, validation)

    with _LOCK:
        previous = os.getcwd()
        try:
            with tempfile.TemporaryDirectory(prefix='finagle-') as tmp:
                os.chdir(tmp)
                try:
                    return _execute(case, compiled, validation)
                finally:
                    os.chdir(previous)
        except BaseException:
            if os.getcwd() != previous:
                os.chdir(previous)
            raise


def sanity_flags(re, gt, cash, fcfe, terminal_wacc, reconciliation=None,
                 mna_total=None, deal_spend_total=None):
    '''Flags a person should read beside the value. They do not change it.'''
    flags = []
    if _any_below(cash, 0):
        flags.append('negative cash')
    if _any_below((fcfe or [])[2:], 0):
        flags.append('negative FCFE after year 1')
    if re is not None and gt is not None and re <= gt:
        flags.append('re <= gt')
    if terminal_wacc is not None and gt is not None and terminal_wacc <= gt:
        flags.append('wacc <= gt')
    for row in reconciliation or []:
        gap = row.get('gap_pct')
        if gap is not None and abs(gap) > _EBITDA_GAP:
            flags.append(
                'model EBITDA differs from the article total by more than 5%'
            )
            break
    if mna_total is not None and deal_spend_total is not None:
        scale = max(abs(mna_total), abs(deal_spend_total))
        if scale and abs(mna_total - deal_spend_total) / scale > _MNA_GAP:
            flags.append('MnA differs from disclosed deal spend by more than 10%')
    return flags


def resolve_ebitda_frac(model, spec):
    '''Convert a dollar deal into the fraction company expects.

    The fraction is acquired EBITDA divided by EBITDA in the deal year as it
    stands just before that deal, which includes earlier deals.
    '''
    year = int(spec['year'])
    organic = float(model.fin['ebitda'].iloc[year])
    if organic == 0:
        raise ValueError(
            'EBITDA in year %d is 0, so acquired EBITDA cannot be converted '
            'to ebitda_frac' % year
        )
    if spec['kind'] == 'acquired_ebitda':
        acquired = float(spec['acquired_ebitda'])
    else:
        multiple = float(spec['multiple'])
        if multiple == 0:
            raise ValueError(
                'acquisition multiple is 0, so deal value cannot be converted to EBITDA'
            )
        acquired = float(spec['deal_value']) / multiple
    return acquired / organic


def _execute(case, compiled, validation):
    model = None
    executed = []
    error = None
    try:
        model = company(**compiled['constructor'])
        for step in compiled['steps']:
            executed.append(_invoke(model, step, compiled['financials']))
    except Exception as exc:
        error = '%s: %s' % (type(exc).__name__, exc)

    log_text = _read_log(model)
    records = _log_records(log_text)
    result = _blocked(case, validation)
    result['path'] = compiled['path']
    result['steps'] = executed
    result['notebook_code'] = notebook_code(compiled, executed)
    if error:
        result['error'] = error
    if model is None or not hasattr(model, 'fin'):
        result['log'] = records
        result['ok'] = False
        return result

    fin = model.fin
    result['value_per_share'] = _at(fin, 'value_per_share', 0)
    result['value_per_share_DDM'] = _at(fin, 'value_per_share_DDM', 0)
    result['vpsbb'] = _number(getattr(model, 'vpsbb', None))
    result['equity'] = _at(fin, 'equity', 0)
    result['firm'] = _at(fin, 'firm', 0)
    result['fcfet'] = _number(getattr(model, 'fcfet', None))
    result['terminal_wacc'] = _at(fin, 'wacc', -1)
    result['years'] = _years(fin)
    result['acquisitions_summary'] = _acquisitions_summary(model)
    result['ebitda_reconciliation'] = _ebitda_reconciliation(result['years'], case)
    result['log'] = records
    rates = unwrap(case).get('rates') or {}
    cash = [row.get('cash') for row in result['years']]
    fcfe = [row.get('fcfe') for row in result['years']]
    mna_total, deal_spend_total = _deal_spend_totals(result['years'], case)
    result['sanity_flags'] = sanity_flags(
        rates.get('re'), rates.get('gt'), cash, fcfe, result['terminal_wacc'],
        result['ebitda_reconciliation'], mna_total, deal_spend_total,
    )
    result['ok'] = error is None and not any(row['level'] == 'ERROR' for row in records)
    return result


def _invoke(model, step, financials):
    args = dict(step.get('args') or {})
    if step.get('resolve_frac'):
        args['ebitda_frac'] = resolve_ebitda_frac(model, step['resolve_frac'])
    method = getattr(model, step['method'])
    if step['method'] == 'load_financials':
        method(financials)
    elif step.get('pass_financials'):
        method(financials=financials, **args)
    else:
        method(**args)
    recorded = {'method': step['method'], 'args': _plain(args)}
    if step.get('pass_financials'):
        recorded['pass_financials'] = True
    return recorded


def notebook_code(compiled, executed):
    '''Python in the shape of the template's scenario-copy cell.

    Methods this run called stay live. The template's other methods stay
    in the cell and are commented. The instance name is co.
    '''
    financials = compiled['financials']
    constructor = compiled['constructor']
    direct_fcfe = 'fcfe' in constructor
    earnings = _named(executed, 'fcf_from_earnings')
    ebitda_path = not direct_fcfe and not earnings
    ebitda_step = _first(executed, 'forecast_ebitda')
    margin = bool(ebitda_step and {'me', 'mc', 'gsnext'} <= set(ebitda_step['args']))

    lines = [
        'import finagle as cmp',
        '',
        '# Scenario copy. Edit this cell; leave the cell above as the base case.',
        '# First variant: uncomment the margin-path forecast_ebitda and comment the growth-only call.',
        '',
    ]
    for key in _SCALAR_ORDER:
        if key in constructor:
            value = constructor[key]
        else:
            value = _SCALAR_DEFAULTS[key]
        lines.append('%s = %r' % (key, value))
    if direct_fcfe:
        lines.append('fcfe = %r' % (constructor['fcfe'],))
    lines.append('')
    if direct_fcfe:
        lines.extend(_financials_lines(financials))
        lines.append('')
    lines.extend(_company_call(direct_fcfe))
    lines.append('')
    if not direct_fcfe:
        lines.append('n = year + 1')
        lines.extend(_financials_lines(financials))
        lines.append('')
    lines.extend(_forecast_lines(executed, ebitda_step, margin))
    lines.append(_comment(
        'co.load_financials(financials=financials.copy())', not direct_fcfe,
    ))
    lines.append(_comment('co.fcf_from_ebitda()', ebitda_path))
    if earnings:
        lines.append(_call_line(earnings[0]))
    else:
        lines.append(
            '# co.fcf_from_earnings(payout=1, gf=0, ROE=1)'
            '  # alternate path; do not combine with EBITDA'
        )
    lines.append('')
    lines.extend(_action_lines(executed))
    lines.append('')
    lines.append('co.value()')
    lines.append(_comment('co.display_fin()', ebitda_path))
    return '\n'.join(lines) + '\n'


def _company_call(direct_fcfe):
    lines = [
        'co = cmp.company(',
        '    ticker=ticker, rd=rd, re=re, t=t, te=te,',
        '    shares=shares, price=price, gt=gt, roict=roict,',
        '    year=year, dividend=dividend,',
    ]
    if direct_fcfe:
        lines.append('    fcfe=fcfe, financials=financials,')
    lines.append(')')
    return lines


def _financials_lines(financials):
    lines = ['financials = {']
    keys = [key for key in _FINANCIAL_KEYS if key in financials]
    keys.extend(key for key in financials if key not in keys)
    for key in keys:
        comment = _FINANCIAL_COMMENTS.get(key)
        suffix = '  # %s' % comment if comment else ''
        lines.append('    %r: %r,%s' % (key, financials[key], suffix))
    lines.append('}')
    return lines


def _forecast_lines(executed, ebitda_step, margin):
    lines = [
        '# gf: explicit EBITDA growth. Interpolated toward gt. Non-NaN prefix = explicit years.',
        '# Growth-only: ebitda[i+1] = ebitda[i] * (1 + g[i]); revenue is 0.',
        '# Margin path (me, mc, and gsnext all set), after the explicit gf years:',
        '#   me = EBITDA/sales at the handoff; revenue = ebitda/me * (1+gs)',
        '#   mc = incremental EBITDA / incremental sales',
        '#   gsnext = sales growth in the first year after explicit EBITDA growth, then faded to gt',
        '#   ebitda[i+1] = ebitda[i] * (1 + (mc/me) * gs[i]); me is updated to the new margin',
    ]
    if ebitda_step:
        args = ebitda_step['args']
        lines.append('ebitda_ttm = %r' % (args['ebitda_ttm'],))
        lines.append('gf = %r' % (args['gf'],))
    else:
        lines.append("# ebitda_ttm = financials['ebitda'][0]")
        lines.append('# gf = [0.10, 0.05]')
    lines.append('')
    lines.append(
        '# Comment this out when ebitda is already a full list. Keep revenue at [0] in that case.'
    )
    lines.append(_comment(
        'co.forecast_ebitda(ebitda_ttm, gf, financials)',
        ebitda_step is not None and not margin,
    ))
    lines.append('')
    if margin:
        args = ebitda_step['args']
        lines.append('me = %r' % (args['me'],))
        lines.append('mc = %r' % (args['mc'],))
        lines.append('gsnext = %r' % (args['gsnext'],))
        lines.append(
            'co.forecast_ebitda(ebitda_ttm, gf, financials, me=me, mc=mc, gsnext=gsnext)'
        )
    else:
        lines.append('# me = 0.25')
        lines.append('# mc = 0.40')
        lines.append('# gsnext = 0.10')
        lines.append(
            '# co.forecast_ebitda(ebitda_ttm, gf, financials, me=me, mc=mc, gsnext=gsnext)'
        )
    lines.append('')
    sbc = _first(executed, 'forecast_sbc')
    if sbc:
        lines.append(_call_line(sbc))
    else:
        lines.append(
            "# co.forecast_sbc(sbc_f=financials['sbc'], sbc_rate_t=None, financials=financials)"
        )
    lines.append('# method omitted: 1 when revenue was forecast, 2 when revenue stays 0.')
    capex = _first(executed, 'forecast_capex')
    if capex:
        lines.append(_call_line(capex))
    else:
        lines.append("# co.forecast_capex(financials['capex'], financials)")
    lines.append('')
    return lines


def _action_lines(executed):
    lines = []
    deals = _named(executed, 'fcf_to_acquire')
    if deals:
        lines.extend(_call_line(step) for step in deals)
    else:
        lines.append(_ACQUIRE_STUB % 1)
        lines.append(_ACQUIRE_STUB % 2)
    disposals = _named(executed, 'noa_to_dispose')
    if disposals:
        lines.extend(_call_line(step) for step in disposals)
    else:
        lines.append('# co.noa_to_dispose(dnoa=0, tax=0, year_dis=1)')
    debt = _first(executed, 'fcf_to_debt')
    if debt:
        lines.append(_call_line(debt))
    else:
        lines.append('# co.fcf_to_debt(leverage=3, year_d=1)')
    lines.append('')
    lines.append('# Exactly one distribution method:')
    for name, stub in _DISTRIBUTION_STUBS:
        step = _first(executed, name)
        lines.append(_call_line(step) if step else stub)
    return lines


def _named(executed, name):
    return [step for step in executed if step['method'] == name]


def _first(executed, name):
    found = _named(executed, name)
    if found:
        return found[0]
    return None


def _comment(line, live):
    if live:
        return line
    return '# ' + line


def _call_line(step):
    name = step['method']
    if name == 'value':
        return 'co.value()'
    if name == 'load_financials':
        return 'co.load_financials(financials=financials.copy())'
    args = step.get('args') or {}
    ordered = []
    for key in _ARG_ORDER.get(name, ()):
        if key in args:
            ordered.append('%s=%r' % (key, args[key]))
    for key in args:
        if key not in _ARG_ORDER.get(name, ()):
            ordered.append('%s=%r' % (key, args[key]))
    if step.get('pass_financials'):
        ordered.append('financials=financials')
    return 'co.%s(%s)' % (name, ', '.join(ordered))


def _acquisitions_summary(model):
    '''One row per deal: the year it is paid, the EBITDA it adds, and the multiple.'''
    deals = getattr(model, '_deals', None) or []
    series = getattr(model, '_deal_ebitdas', None) or []
    rows = []
    for deal, acquired_path in zip(deals, series):
        year = int(deal['year_a'])
        acquired = None
        if year + 1 < len(acquired_path):
            acquired = _number(acquired_path[year + 1])
        multiple = _number(deal.get('multiple'))
        mna = None
        if acquired is not None and multiple is not None:
            mna = multiple * acquired
        implied = None
        if mna is not None and acquired:
            implied = mna / acquired
        rows.append({
            'year': year,
            'mna': mna,
            'acquired_ebitda': acquired,
            'multiple': implied,
        })
    return rows


def _ebitda_reconciliation(years, case):
    '''Compare model EBITDA with the article's total, when the article states one.'''
    article = unwrap(case).get('article') or {}
    stated = article.get('ebitda_total')
    if not isinstance(stated, list):
        return []
    rows = []
    for index, article_ebitda in enumerate(stated):
        if index >= len(years):
            break
        model_ebitda = years[index].get('ebitda')
        gap = None
        gap_pct = None
        if _number(model_ebitda) is not None and _number(article_ebitda) is not None:
            gap = model_ebitda - article_ebitda
            if article_ebitda:
                gap_pct = gap / float(article_ebitda)
        rows.append({
            'year': index,
            'model': model_ebitda,
            'article': article_ebitda,
            'gap': gap,
            'gap_pct': gap_pct,
        })
    return rows


def _deal_spend_totals(years, case):
    '''Return total MnA and disclosed deal spend, or (None, None) when undisclosed.'''
    forecast = unwrap(case).get('forecast') or {}
    disclosure = forecast.get('acquisition_disclosure') or {}
    if not isinstance(disclosure, dict) or disclosure.get('status') != 'future_deals':
        return None, None
    spend = disclosure.get('deal_spend')
    if not isinstance(spend, list):
        return None, None
    spend_total = 0
    for item in spend:
        number = _number(item)
        if number is not None:
            spend_total += number
    mna_total = 0
    for row in years:
        number = _number(row.get('MnA'))
        if number is not None:
            mna_total += number
    return mna_total, spend_total


def _years(fin):
    rows = []
    for index, label in enumerate(list(fin.index)):
        row = {'year': index, 'date': _date(label)}
        for name in _YEAR_COLUMNS:
            if name in fin.columns:
                row[name] = _number(fin[name].iloc[index])
        rows.append(row)
    return rows


def _blocked(case, validation):
    target = None
    if isinstance(case, dict):
        article = unwrap(case).get('article') or {}
        target = article.get('price_target')
        ticker = case.get('ticker')
        path = case.get('path') or 'ebitda'
    else:
        ticker = None
        path = None
    return {
        'ok': False,
        'ticker': ticker,
        'path': path,
        'value_per_share': None,
        'value_per_share_DDM': None,
        'vpsbb': None,
        'equity': None,
        'firm': None,
        'fcfet': None,
        'terminal_wacc': None,
        'article_price_target': target,
        'years': [],
        'acquisitions_summary': [],
        'ebitda_reconciliation': [],
        'log': [],
        'sanity_flags': [],
        'validation': validation,
        'case': case,
        'steps': [],
        'notebook_code': '',
        'provenance': provenance(case),
    }


def _read_log(model):
    if model is None:
        return ''
    path = getattr(model, 'logfile', None)
    if not path or not os.path.exists(path):
        return ''
    with open(path, encoding='utf-8') as handle:
        return handle.read()


def _log_records(text):
    records = []
    for line in text.splitlines():
        for level in ('WARNING', 'ERROR'):
            token = ' %s:' % level
            index = line.find(token)
            if index != -1:
                records.append({
                    'level': level,
                    'message': line[index + len(token):].strip(),
                })
                break
    return records


def _at(fin, column, index):
    if column not in getattr(fin, 'columns', ()):
        return None
    try:
        return _number(fin[column].iloc[index])
    except (TypeError, IndexError, KeyError, AttributeError):
        return None


def _any_below(values, limit):
    for value in values or []:
        if value is not None and value < limit:
            return True
    return False


def _date(value):
    if hasattr(value, 'strftime'):
        return value.strftime('%Y-%m-%d')
    return str(value)


def _number(value):
    if value is None or isinstance(value, bool):
        return value
    try:
        if value != value:
            return None
    except TypeError:
        return value
    if hasattr(value, 'item'):
        try:
            value = value.item()
        except ValueError:
            pass
    if isinstance(value, float):
        return float(value)
    if isinstance(value, int):
        return int(value)
    return value


def _plain(value):
    if isinstance(value, dict):
        return {key: _plain(item) for key, item in value.items()}
    if isinstance(value, list):
        return [_plain(item) for item in value]
    return _number(value)

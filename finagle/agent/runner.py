'''Run a case on the existing company class and return plain numbers.

The run happens in a temporary directory so the log file company writes does
not land next to a person's notebooks. display_fin is not called.
'''

import os
import pprint
import tempfile
import threading

from finagle.agent.compile import CompileError, compile_case
from finagle.agent.schema import provenance, unwrap
from finagle.agent.validate import validate_case
from finagle.company import company

_LOCK = threading.Lock()

_YEAR_COLUMNS = (
    'ebitda', 'capex', 'fcf', 'fcfe', 'fcff', 'debt', 'cash', 'cashBS',
    'shares', 'price', 'dividend', 'equity', 'firm',
)

_CTOR_ORDER = (
    'ticker', 'rd', 're', 't', 'te', 'shares', 'price', 'gt', 'roict',
    'year', 'dividend', 'fcfe',
)

_ARG_ORDER = {
    'forecast_ebitda': ('ebitda_ttm', 'gf', 'me', 'mc', 'gsnext'),
    'forecast_capex': ('capex_f',),
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
    '''
    validation = validate_case(case)
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


def sanity_flags(re, gt, cash, fcfe, terminal_wacc):
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
    result['log'] = records
    rates = unwrap(case).get('rates') or {}
    cash = [row.get('cash') for row in result['years']]
    fcfe = [row.get('fcfe') for row in result['years']]
    result['sanity_flags'] = sanity_flags(
        rates.get('re'), rates.get('gt'), cash, fcfe, result['terminal_wacc'],
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
    '''Python a person can paste into a notebook to rerun the same case.'''
    financials = compiled['financials']
    lines = [
        'import finagle as cmp',
        '',
        'financials = ' + pprint.pformat(financials, sort_dicts=False),
        '',
    ]
    constructor = compiled['constructor']
    parts = []
    for key in _CTOR_ORDER:
        if key in constructor:
            parts.append('%s=%r' % (key, constructor[key]))
    if 'financials' in constructor:
        parts.append('financials=financials')
    lines.append('model = cmp.company(%s)' % ', '.join(parts))
    for step in executed:
        lines.append(_call_line(step))
    return '\n'.join(lines) + '\n'


def _call_line(step):
    name = step['method']
    if name == 'value':
        return 'model.value()'
    if name == 'load_financials':
        return 'model.load_financials(financials=financials.copy())'
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
    return 'model.%s(%s)' % (name, ', '.join(ordered))


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
        'log': [],
        'sanity_flags': [],
        'validation': validation,
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

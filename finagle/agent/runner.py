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
    'ebitda', 'capex', 'MnA', 'dDebt', 'fcf', 'fcfe', 'fcff', 'debt', 'cash',
    'cashBS', 'shares', 'price', 'dividend', 'equity', 'firm',
)

# Model EBITDA more than this far from the article's total, or MnA more than
# this far from the disclosed deal spend, is called out beside the value.
_EBITDA_GAP = 0.05
_MNA_GAP = 0.10

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

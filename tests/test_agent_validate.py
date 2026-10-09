from finagle.agent.catalog import VALUE_FROM_ARTICLE, describe_inputs
from finagle.agent.schema import CASE_JSON_SCHEMA, case_tool_schema
from finagle.agent.validate import validate_case


def leaf(value, **extra):
    item = {
        'value': value,
        'source': 'assumption',
        'evidence': 'stated for this test',
    }
    item.update(extra)
    return item


def minimal_case():
    return {
        'path': 'ebitda',
        'ticker': 'TEST',
        'year': 2,
        'market': {'shares': leaf(10), 'price': leaf(20)},
        'rates': {
            're': leaf(0.10),
            'rd': leaf(0.05),
            't': leaf(0.21),
            'gt': leaf(0.02),
            'roict': leaf(0.15),
        },
        'baseline': {
            'date': leaf('2021-12-31'),
            'ebitda': leaf(10),
            'ebitda_definition': leaf('before_sbc'),
            'capex': leaf(1),
            'da': leaf(1),
            'tax': leaf(1),
            'interest': leaf(1),
            'sbc': leaf(0),
            'debt': leaf(5),
            'cash': leaf(2),
            'nol': leaf(0),
            'noa': leaf(0),
        },
        'forecast': {
            'ebitda_growth': leaf([0.10]),
            'dwc': {'mode': leaf('zero')},
        },
        'debt': {'mode': leaf('hold')},
        'distribution': {'mode': leaf('none')},
    }


def test_minimal_case_is_complete():
    report = validate_case(minimal_case())
    assert report['errors'] == []
    assert report['missing'] == []
    assert report['warnings'] == []


def test_duplicate_debt_target_is_rejected():
    case = minimal_case()
    case['debt_target'] = {'mode': leaf('target'), 'leverage': leaf(2)}
    report = validate_case(case)
    assert any('only one debt target' in error for error in report['errors'])


def test_short_dwc_path_is_rejected():
    case = minimal_case()
    case['forecast']['dwc'] = {
        'mode': leaf('path'),
        'values': leaf([0, 1]),
    }
    report = validate_case(case)
    assert any('short' in error and 'dwc' in error for error in report['errors'])


def test_bare_number_is_missing_provenance():
    case = minimal_case()
    case['rates']['re'] = 0.10
    report = validate_case(case)
    assert any('rates.re' in error and 'missing provenance' in error for error in report['errors'])


def test_units_far_from_a_normal_multiple_warn():
    case = minimal_case()
    case['market']['shares'] = leaf(10_000_000)
    report = validate_case(case)
    assert report['errors'] == []
    assert any('millions' in warning for warning in report['warnings'])


def test_cost_of_equity_must_exceed_terminal_growth():
    case = minimal_case()
    case['rates']['re'] = leaf(0.02)
    case['rates']['gt'] = leaf(0.02)
    report = validate_case(case)
    assert any('re must be greater than gt' in error for error in report['errors'])


def test_mixed_period_ends_and_a_stale_price_warn():
    case = minimal_case()
    case['baseline']['ebitda'] = leaf(10, period_end='2020-12-31')
    case['baseline']['capex'] = leaf(1, period_end='2021-12-31')
    case['market']['price'] = leaf(20, as_of='2020-01-01')
    case['article'] = {'published': leaf('2024-06-01')}
    report = validate_case(case)
    assert any('different period ends' in warning for warning in report['warnings'])
    assert any('older than the article date' in warning for warning in report['warnings'])


def test_ebitda_after_sbc_is_rejected():
    case = minimal_case()
    case['baseline']['ebitda_definition'] = leaf('after_sbc')
    report = validate_case(case)
    assert any('before stock-based compensation' in error for error in report['errors'])


def test_full_ebitda_path_cannot_also_carry_growth():
    case = minimal_case()
    case['forecast']['ebitda'] = leaf([10, 11, 12])
    report = validate_case(case)
    assert any('cannot both be set' in error for error in report['errors'])


def test_schema_names_the_case_fields():
    assert CASE_JSON_SCHEMA['properties']['ticker']['type'] == 'string'
    assert 'leaf' in CASE_JSON_SCHEMA['$defs']
    tool_schema = case_tool_schema()
    assert 'ticker' in tool_schema['properties']['case']['properties']
    assert 'leaf' in tool_schema['$defs']


def test_catalog_and_prompt_tell_the_bot_what_to_supply():
    catalog = describe_inputs()
    paths = [field['path'] for field in catalog['fields']]
    assert 'baseline.ebitda' in paths
    assert 'distribution' in paths
    assert 'article.price_target' in paths
    assert 'millions of dollars' in catalog['units']['money']
    assert 'before stock-based compensation' in catalog['ebitda_definition']
    assert 'describe_inputs' in VALUE_FROM_ARTICLE
    assert 'get_10k_baseline' in VALUE_FROM_ARTICLE
    assert 'validate_case' in VALUE_FROM_ARTICLE
    assert 'run_case' in VALUE_FROM_ARTICLE
    assert 'price target' in VALUE_FROM_ARTICLE.lower()


def test_package_import_does_not_pull_in_the_agent():
    import finagle
    text = open(finagle.__file__, encoding='utf-8').read()
    assert 'agent' not in text

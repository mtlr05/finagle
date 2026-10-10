'''Valuation case: provenance leaves, intents, and the JSON Schema an agent fills.'''

import copy

SOURCES = ('article', 'attachment', '10k', 'yahoo', 'assumption')
BASELINE_PRIORITY = ('article', 'attachment', '10k', 'yahoo')
BASES = ('fiscal_year', 'ttm', 'quarter')
PATHS = ('ebitda', 'earnings', 'fcfe')
DEBT_MODES = ('hold', 'path', 'target')
DWC_MODES = ('zero', 'path')
DISTRIBUTION_MODES = ('none', 'max_buybacks', 'buyback_schedule', 'retain')
PRICE_PATHS = ('constant', 'proportional')
EBITDA_DEFINITIONS = ('before_sbc', 'after_sbc')
EBITDA_BASES = ('organic', 'total')
ACQUISITION_DISCLOSURES = ('future_deals', 'none')

_LEAF = {
    'type': 'object',
    'additionalProperties': False,
    'required': ['value', 'source', 'evidence'],
    'properties': {
        'value': {
            'description': (
                'The figure itself: a number, a string, a boolean, null, '
                'or an array of numbers. Rates are decimals (0.09 is 9%). '
                'Money is millions in the reporting currency. Shares are millions of shares. '
                'Price is per share, in the same currency as the financials.'
            ),
        },
        'source': {
            'type': 'string',
            'enum': list(SOURCES),
            'description': (
                'article, attachment (a filing the user attached), 10k (SEC), '
                'yahoo (yfinance), or assumption.'
            ),
        },
        'evidence': {
            'type': 'string',
            'description': (
                'A short quote from the article, a page and quote from the attachment, '
                'an XBRL tag, a Yahoo row name, or the reason for an assumption.'
            ),
        },
        'period_end': {
            'type': 'string',
            'description': 'YYYY-MM-DD period the figure belongs to, when it is a reported number.',
        },
        'basis': {
            'type': 'string',
            'enum': list(BASES),
            'description': (
                'fiscal_year, ttm, or quarter. Reported baseline figures must be '
                'fiscal_year: the last completed fiscal year, not trailing twelve months.'
            ),
        },
        'currency': {
            'type': 'string',
            'description': 'ISO currency code of a money figure or price, such as USD or CAD.',
        },
        'as_of': {
            'type': 'string',
            'description': 'YYYY-MM-DD date the figure was observed. Use this on market.price.',
        },
        'note': {'type': 'string'},
    },
}

_LEAF_REF = {'$ref': '#/$defs/leaf'}


def _leaf_prop(description):
    return {
        'description': description,
        'allOf': [_LEAF_REF],
    }


CASE_JSON_SCHEMA = {
    '$schema': 'https://json-schema.org/draft/2020-12/schema',
    '$defs': {'leaf': _LEAF},
    'type': 'object',
    'additionalProperties': False,
    'required': ['ticker', 'year'],
    'description': (
        'One valuation. Year index 0 is the baseline period. '
        'Every figure is a leaf with value, source, and evidence. '
        'ticker, year, and path are plain. Do not send a list of method calls.'
    ),
    'properties': {
        'path': {
            'type': 'string',
            'enum': list(PATHS),
            'description': 'ebitda (default), earnings, or fcfe. fcfe discounts a cash-flow series directly.',
        },
        'ticker': {'type': 'string', 'description': 'Ticker symbol. Required. Used only as a label.'},
        'year': {
            'type': 'integer',
            'minimum': 1,
            'description': 'Last explicit forecast year. A full series has year + 1 values, starting at the baseline.',
        },
        'market': {
            'type': 'object',
            'additionalProperties': False,
            'properties': {
                'shares': _leaf_prop('Shares outstanding, all classes, in millions.'),
                'price': _leaf_prop(
                    'Current price per share, in the same currency as the baseline. '
                    'This is not the article\'s price target.'
                ),
            },
        },
        'rates': {
            'type': 'object',
            'additionalProperties': False,
            'description': 'Discount and tax assumptions. re must be greater than gt.',
            'properties': {
                're': _leaf_prop('Cost of equity, decimal.'),
                'rd': _leaf_prop('Cost of debt, decimal.'),
                't': _leaf_prop('Marginal tax rate, decimal.'),
                'te': _leaf_prop('Optional effective tax rate for year 1. Omit to let the model derive it.'),
                'gt': _leaf_prop('Terminal growth rate, decimal. Must be below re.'),
                'roict': _leaf_prop('Terminal ROIC, decimal. Used to set terminal depreciation.'),
            },
        },
        'baseline': {
            'type': 'object',
            'additionalProperties': False,
            'description': (
                'Year-0 figures, in millions of the reporting currency. Year 0 is the last '
                'completed fiscal year, not trailing twelve months, because EBITDA growth is '
                'measured from it. Fill each figure from the article, then an attached filing, '
                'then the SEC (not for Canadian companies), then yfinance; build_baseline does this. '
                'Reported figures need period_end equal to date, basis fiscal_year, and one currency. '
                'EBITDA must be before stock-based compensation, because the model subtracts sbc. '
                'The SEC and yfinance helpers already add sbc back.'
            ),
            'properties': {
                'date': _leaf_prop('Baseline period end, YYYY-MM-DD.'),
                'ebitda': _leaf_prop('Year-0 EBITDA before stock-based compensation, millions.'),
                'ebitda_definition': _leaf_prop('before_sbc or after_sbc. The model can only run before_sbc.'),
                'revenue': _leaf_prop('Informational. The model fills the revenue column itself.'),
                'capex': _leaf_prop('Year-0 capital expenditure, millions. Positive is a cash outflow.'),
                'da': _leaf_prop('Year-0 depreciation, or a partial list. Later years are calculated.'),
                'tax': _leaf_prop('Year-0 cash tax, millions.'),
                'interest': _leaf_prop('Year-0 interest expense, millions.'),
                'sbc': _leaf_prop('Year-0 stock-based compensation, millions. Use 0 only as an explicit assumption.'),
                'debt': _leaf_prop('Year-0 short-term plus long-term debt, millions. Not a full forecast.'),
                'cash': _leaf_prop('Year-0 excess cash and short-term investments, millions.'),
                'nol': _leaf_prop('Year-0 net operating loss, millions. Use 0 only as an explicit assumption.'),
                'noa': _leaf_prop('Year-0 non-operating assets, millions. Use 0 only as an explicit assumption.'),
            },
        },
        'forecast': {
            'type': 'object',
            'additionalProperties': False,
            'description': (
                'How the baseline becomes a full horizon. '
                'Set ebitda_growth or a full ebitda path, not both. '
                'That path is organic unless ebitda_basis is total. '
                'Deals that close after year 0 are acquisitions, not extra EBITDA in the path. '
                'capex and sbc may be shorter than the horizon; they are extended. '
                'dwc is not extended: zero, or one value per year.'
            ),
            'properties': {
                'ebitda_basis': _leaf_prop(
                    'organic or total. organic excludes deals that close after year 0. '
                    'total already includes acquired EBITDA and is only valid when there are no future deals.'
                ),
                'acquisition_disclosure': {
                    'type': 'object',
                    'additionalProperties': False,
                    'description': (
                        'Required on the EBITDA path. status future_deals means the forecast buys EBITDA '
                        'after year 0 and deal_spend is the enterprise value paid in each year, the dollars '
                        'that become MnA. status none means the forecast has no future deals; the evidence '
                        'must say so.'
                    ),
                    'properties': {
                        'status': _leaf_prop('future_deals or none.'),
                        'deal_spend': _leaf_prop(
                            'Enterprise value paid by year, millions. Index 0 is the baseline. '
                            'Required when status is future_deals. Same dollars as deal_value, which become MnA.'
                        ),
                    },
                },
                'ebitda_growth': _leaf_prop(
                    'EBITDA growth rates after year 0, for the organic business and deals already paid for. '
                    'A short list fades toward gt. Do not include EBITDA from deals that close after year 0.'
                ),
                'ebitda': _leaf_prop(
                    'Explicit EBITDA for every year, including year 0, on the basis in ebitda_basis. '
                    'Do not also set ebitda_growth.'
                ),
                'margin_path': {
                    'description': (
                        'Optional. Requires ebitda_growth. me is the ending EBITDA margin, '
                        'mc is the contribution margin, gsnext is the next sales growth. '
                        'me and mc must be non-zero.'
                    ),
                    'type': 'object',
                    'additionalProperties': False,
                    'properties': {
                        'me': _leaf_prop('Ending EBITDA margin, decimal.'),
                        'mc': _leaf_prop('Contribution margin, decimal. Must be non-zero.'),
                        'gsnext': _leaf_prop('Sales growth in the year after the explicit EBITDA growth.'),
                    },
                },
                'capex': _leaf_prop(
                    'Capex from year 0. A short list is extended at the last capex-to-EBITDA ratio.'
                ),
                'sbc': _leaf_prop('SBC from year 0. A short list is extended toward sbc_rate_terminal.'),
                'sbc_rate_terminal': _leaf_prop(
                    'Terminal SBC as a fraction of EBITDA. Used only when sbc does not already cover every year.'
                ),
                'dwc': {
                    'type': 'object',
                    'additionalProperties': False,
                    'description': 'Change in non-cash working capital. zero, or an explicit path covering every year.',
                    'properties': {
                        'mode': _leaf_prop('zero or path.'),
                        'values': _leaf_prop('One value per year, millions, when mode is path. Index 0 is the baseline.'),
                    },
                },
            },
        },
        'dividend_per_share': _leaf_prop(
            'Dividend per share. A number, or a list whose index 0 is the baseline year. '
            'After the list ends, the total dividend grows with free cash flow and is never cut. Omit for 0.'
        ),
        'debt': {
            'type': 'object',
            'additionalProperties': False,
            'description': (
                'One debt policy. hold keeps baseline debt flat. '
                'path uses values for every year. '
                'target starts from values, or from flat baseline debt, then moves toward leverage '
                'from start_year. Do not send a second target.'
            ),
            'properties': {
                'mode': _leaf_prop('hold, path, or target.'),
                'leverage': _leaf_prop('Debt/EBITDA target. Required when mode is target.'),
                'start_year': _leaf_prop(
                    'First forecast year the target applies. 1 is the year after the baseline. Required for target.'
                ),
                'values': _leaf_prop(
                    'Debt by year, millions. Required for path. Optional for target.'
                ),
            },
        },
        'acquisitions': {
            'type': 'array',
            'description': (
                'Bolt-on deals, in order. Give acquired_ebitda, or deal_value together with multiple. '
                'ebitda_frac is the acquired EBITDA divided by total EBITDA in the deal year '
                'after earlier deals; leave it unset unless that fraction is already known.'
            ),
            'items': {
                'type': 'object',
                'additionalProperties': False,
                'properties': {
                    'year': _leaf_prop('Year index of the deal. 0 is the baseline year.'),
                    'ebitda_frac': _leaf_prop('Acquired EBITDA as a fraction of EBITDA in the deal year.'),
                    'acquired_ebitda': _leaf_prop('Acquired EBITDA, millions.'),
                    'deal_value': _leaf_prop('Enterprise value paid, millions. Divided by multiple to get EBITDA.'),
                    'multiple': _leaf_prop('EV/EBITDA paid.'),
                    'leverage': _leaf_prop('Debt/EBITDA added with the deal.'),
                    'next_growth': _leaf_prop('Growth of the acquired EBITDA in the following year.'),
                    'capex_frac': _leaf_prop('Capex added, as a fraction of acquired EBITDA.'),
                    'pay_from_cash': _leaf_prop(
                        'True reduces year-0 cash when the deal is in year 0. Otherwise false.'
                    ),
                },
            },
        },
        'disposals': {
            'type': 'array',
            'description': 'Non-operating asset sales.',
            'items': {
                'type': 'object',
                'additionalProperties': False,
                'properties': {
                    'year': _leaf_prop('Year index of the sale.'),
                    'amount': _leaf_prop('Gross proceeds before tax, millions.'),
                    'tax': _leaf_prop('Tax rate on the sale, decimal. Omit for 0.'),
                },
            },
        },
        'distribution': {
            'type': 'object',
            'additionalProperties': False,
            'description': (
                'One payout policy. none pays all free cash flow to equity out as dividends '
                'in the year it is earned, with no buybacks. '
                'max_buybacks spends available cash on repurchases. '
                'buyback_schedule uses the dollar amounts given; index 0 is the baseline year, '
                'and 0 means no repurchase that year with the rest of the cash retained. '
                'After the list ends the last amount repeats in proportion to free cash flow, '
                'so end the list with 0 when the program stops. '
                'retain keeps undistributed cash on the balance sheet and pays it out in the last year. '
                'price_path constant keeps the given price; proportional holds the year-1 EV/EBITDA multiple.'
            ),
            'properties': {
                'mode': _leaf_prop('none, max_buybacks, buyback_schedule, or retain.'),
                'schedule': _leaf_prop(
                    'Buyback dollars by year, millions. Index 0 is the baseline. '
                    'End with 0 when the program stops; otherwise the last amount continues.'
                ),
                'price': _leaf_prop('Price used to turn buyback dollars into shares.'),
                'price_path': _leaf_prop('constant or proportional.'),
            },
        },
        'earnings': {
            'type': 'object',
            'additionalProperties': False,
            'description': 'Used only when path is earnings.',
            'properties': {
                'e': _leaf_prop('Year-0 earnings.'),
                'payout': _leaf_prop('Payout ratio, or a list of payout ratios.'),
                'gf': _leaf_prop('Earnings growth rates.'),
                'roe': _leaf_prop('Terminal return on equity, decimal.'),
            },
        },
        'fcfe': _leaf_prop(
            'Free cash flow to equity for every year, including a baseline value that is not discounted. '
            'Used only when path is fcfe.'
        ),
        'article': {
            'type': 'object',
            'additionalProperties': False,
            'description': 'Where the case came from. price_target is reported beside the result and is not an input.',
            'properties': {
                'title': _leaf_prop('Article title.'),
                'url': _leaf_prop('Article URL.'),
                'published': _leaf_prop('Publication date, YYYY-MM-DD.'),
                'price_target': _leaf_prop('The article\'s price target, if it states one. Not a model input.'),
                'ebitda_total': _leaf_prop(
                    'The article\'s total EBITDA by year, including acquired EBITDA, if it states one. '
                    'Reported beside the model EBITDA. Not a model input.'
                ),
            },
        },
    },
}


def case_tool_schema():
    '''Tool input schema whose ``case`` property is the valuation case.

    ``$ref`` pointers are resolved from the document root, so the leaf
    definition is lifted to that root when the case sits under ``case``.
    '''
    case = copy.deepcopy(CASE_JSON_SCHEMA)
    definitions = case.pop('$defs')
    case.pop('$schema', None)
    return {
        'type': 'object',
        'properties': {
            'case': case,
        },
        'required': ['case'],
        '$defs': definitions,
    }


def is_leaf(node):
    return (
        isinstance(node, dict)
        and 'value' in node
        and 'source' in node
        and 'evidence' in node
    )


def leaf(value, source, evidence, period_end=None, as_of=None, note=None,
         basis=None, currency=None):
    '''Build one provenance leaf.'''
    item = {'value': value, 'source': source, 'evidence': evidence}
    if period_end is not None:
        item['period_end'] = period_end
    if basis is not None:
        item['basis'] = basis
    if currency is not None:
        item['currency'] = currency
    if as_of is not None:
        item['as_of'] = as_of
    if note:
        item['note'] = note
    return item


def unwrap(node):
    '''Replace every leaf with its value. Structural objects stay nested.'''
    if is_leaf(node):
        return copy.deepcopy(node.get('value'))
    if isinstance(node, dict):
        return {key: unwrap(value) for key, value in node.items()}
    if isinstance(node, list):
        return [unwrap(value) for value in node]
    return node


def iter_leaves(node, prefix=''):
    '''Yield ``(path, leaf)`` for every provenance leaf under ``node``.'''
    if is_leaf(node):
        if prefix:
            yield prefix, node
        return
    if isinstance(node, dict):
        for key, child in node.items():
            if child is None:
                continue
            path = '%s.%s' % (prefix, key) if prefix else str(key)
            if is_leaf(child):
                yield path, child
            else:
                for item in iter_leaves(child, path):
                    yield item
    elif isinstance(node, list):
        for index, child in enumerate(node):
            if child is None:
                continue
            path = '%s[%d]' % (prefix, index)
            if is_leaf(child):
                yield path, child
            else:
                for item in iter_leaves(child, path):
                    yield item


def provenance(case):
    '''Echo the leaves on a case, in walk order.'''
    if not isinstance(case, dict):
        return []
    rows = []
    for path, node in iter_leaves(case):
        rows.append({
            'path': path,
            'value': node.get('value'),
            'source': node.get('source'),
            'evidence': node.get('evidence'),
            'period_end': node.get('period_end'),
            'basis': node.get('basis'),
            'currency': node.get('currency'),
            'as_of': node.get('as_of'),
        })
    return rows

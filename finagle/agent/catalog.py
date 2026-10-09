'''What an agent needs to know before it fills a case.'''

VALUE_FROM_ARTICLE = '''You are valuing one publicly traded company from an article. The article is already in your context. Do not fetch it. Seeking Alpha and Substack posts are often paywalled, and no tool downloads them.

Use the finagle tools. Do not invent a list of Python method calls. Fill a case, which is a set of intents. The runner turns those intents into the existing company methods.

1. Extract every figure the article actually states: baseline results, guidance, growth, margins, capex, buybacks, dividends, leverage, acquisitions, asset sales, share count, and the current price. Keep a short quote for each one.
2. Call describe_inputs. Match what you extracted to those fields. Note the units: money in millions of dollars, shares in millions, price in dollars per share, rates as decimals.
3. Call get_10k_baseline for any baseline field you still do not have. The 10-K figures are the latest fiscal year, in millions. That matches trailing twelve months only while that 10-K is still the latest report. If the article uses a newer quarter, prefer the article and record both period ends. Debt from the 10-K is one year, not a forecast.
4. Anything still empty is an assumption. Set source to assumption and write the reason in evidence. Do not leave a required field out, and do not silently use zero for working capital, net operating losses, non-operating assets, or stock-based compensation.
5. Call validate_case. Fix every error and every missing field. Warnings stay in the result; read them before you run.
6. Call run_case.
7. Report value per share from the free-cash-flow-to-equity model and from the dividend discount model, the terminal free cash flow to equity, and the assumptions (every leaf whose source is assumption). If the result has a log error, or ok is false, the run failed: report the error and do not present the number as a valuation. Place the article's price target beside the result when the article states one.

EBITDA must be before stock-based compensation. The model subtracts stock-based compensation, and the 10-K baseline already adds it back. If the article's EBITDA is after stock-based compensation, add it back and set ebitda_definition to before_sbc.

Do not tune any input so the model matches the article's price target. The price target is not an input. It belongs only in article.price_target.
'''


def describe_inputs():
    '''Field catalog, units, and the article cues that map onto intents.'''
    return {
        'units': {
            'money': 'millions of dollars',
            'shares': 'millions of shares',
            'price': 'dollars per share',
            'rates': 'decimals; 0.09 means 9 percent',
        },
        'year_index': (
            'Index 0 is the baseline period. Index 1 is the first forecast year. '
            'year is the last explicit year, so a full series has year + 1 values. '
            'The baseline is the latest 10-K fiscal year, or the article\'s trailing '
            'twelve months when that is newer.'
        ),
        'ebitda_definition': (
            'EBITDA must be before stock-based compensation. The model subtracts sbc. '
            'The 10-K helper adds sbc back into EBITDA. If an article\'s EBITDA is '
            'after stock-based compensation, add sbc back and record '
            'ebitda_definition as before_sbc. after_sbc is rejected.'
        ),
        'paths': {
            'ebitda': 'Operating forecast, free cash flow, then value. This is the usual path.',
            'earnings': 'Grow earnings and apply a payout. The result is equity only, with no value per share.',
            'fcfe': 'Discount a free-cash-flow-to-equity series directly. The result is equity only.',
        },
        'debt_modes': {
            'hold': 'Keep year-0 debt flat for the whole horizon.',
            'path': 'Use an explicit debt balance for every year.',
            'target': (
                'Start from an explicit path, or from flat year-0 debt, then move '
                'toward leverage (Debt/EBITDA) beginning at start_year.'
            ),
        },
        'dwc_modes': {
            'zero': 'No change in non-cash working capital.',
            'path': 'One change per year, including the baseline. This series is not extended for you.',
        },
        'distribution_modes': {
            'none': 'Do not allocate free cash flow. Dividends and buybacks stay at the dividend policy.',
            'max_buybacks': 'Spend available cash on repurchases.',
            'buyback_schedule': (
                'Repurchase the dollar amounts given, in millions. Index 0 is the '
                'baseline year and is usually 0 when the repurchase is in the future. '
                'A 0 keeps that year\'s cash on the balance sheet. This is not '
                '"buy back as much as possible".'
            ),
            'retain': 'Keep undistributed cash on the balance sheet and pay it out in the last year.',
        },
        'acquisition_size': (
            'Prefer acquired_ebitda in millions, or deal_value and multiple '
            '(acquired EBITDA = deal value / multiple). ebitda_frac is that EBITDA '
            'divided by total EBITDA in the deal year after earlier deals. Leave '
            'the fraction unset unless you already have it.'
        ),
        'price_path': {
            'constant': 'Hold the buyback price fixed.',
            'proportional': 'Hold the year-1 EV/EBITDA multiple, which is the default.',
        },
        'fields': _fields(),
        'article_cues': [
            'Guidance, a consensus forecast, or the author\'s revenue and profit path becomes ebitda_growth, or a full ebitda path when the article gives the dollars.',
            'A margin target plus sales growth becomes margin_path: me is the ending EBITDA margin, mc is the contribution margin, and gsnext is the next year\'s sales growth. me and mc must be non-zero.',
            'A capex guide or maintenance capex becomes forecast.capex. A short list is extended at the last capex-to-EBITDA ratio.',
            'Stock-based compensation as a share of profit becomes forecast.sbc and sbc_rate_terminal.',
            'A working-capital build or release becomes forecast.dwc mode path. If the article is silent, mode zero is an assumption and must be labeled as one.',
            'A leverage target or a debt paydown becomes debt mode target. A stated debt schedule becomes path. Silence about a change becomes hold.',
            'A buyback authorization or a dollar repurchase program becomes distribution mode buyback_schedule. Language that all free cash flow will be used for buybacks becomes max_buybacks.',
            'A dividend per share becomes dividend_per_share.',
            'A bolt-on acquisition becomes one acquisitions entry. Use the deal value and the multiple, or the acquired EBITDA.',
            'An asset sale or a stake sale becomes one disposals entry.',
            'Excess cash, net operating losses, and non-operating assets become cash, nol, and noa. Use 0 only as an explicit assumption.',
            'The current share price becomes market.price, with as_of set to the date of that price. The author\'s price target becomes article.price_target and nothing else.',
            'If the article does not give a discount rate, terminal growth, or terminal ROIC, those are assumptions. Say what you used and why.',
        ],
        'do_not': [
            'Do not tune inputs so value_per_share matches the article\'s price target.',
            'Do not put the price target anywhere except article.price_target.',
            'Do not send a list of company method calls. Send the intents in the case.',
            'Do not treat 10-K debt as a full-horizon series. It is year 0 only.',
            'Do not treat the latest 10-K as trailing twelve months once a later 10-Q is the source of the article\'s numbers.',
            'Do not call display_fin. run_case returns the numbers.',
        ],
    }


def _fields():
    return [
        _field('ticker', 'string', 'Article or the user', 'The company the article is about.', True),
        _field('year', 'integer', 'assumption', 'How many years to forecast before the terminal year. 10 is a common assumption.', True),
        _field('path', 'ebitda, earnings, or fcfe', 'assumption', 'Omit to use the EBITDA path.', False),
        _field('market.shares', 'millions of shares', '10-K share count or the article', 'All classes. The 10-K helper returns this beside the dollar baseline, not inside it.', True),
        _field('market.price', 'dollars per share', 'article or a quoted market price', 'Current price, with as_of. Required for buybacks when distribution.price is omitted.', False),
        _field('rates.re', 'decimal', 'assumption, unless the article builds a discount rate', 'Cost of equity. Must be greater than gt.', True),
        _field('rates.rd', 'decimal', 'assumption or the article\'s stated cost of debt', 'Cost of debt. Required on the EBITDA path.', True),
        _field('rates.t', 'decimal', 'assumption; 0.21 is the US federal statutory rate', 'Marginal tax rate.', True),
        _field('rates.te', 'decimal', 'article, when year-1 cash tax is not the marginal rate', 'Optional effective tax rate for year 1.', False),
        _field('rates.gt', 'decimal', 'assumption', 'Terminal growth. Must be below re.', True),
        _field('rates.roict', 'decimal', 'assumption', 'Terminal return on invested capital, used for terminal depreciation.', True),
        _field('baseline.date', 'YYYY-MM-DD', 'article period or 10-K period end', 'The baseline period, year index 0.', True),
        _field('baseline.ebitda', 'millions', 'article or 10-K', 'Year-0 EBITDA before stock-based compensation.', True),
        _field('baseline.ebitda_definition', 'before_sbc or after_sbc', 'article wording or 10-K', 'Record which EBITDA you used. after_sbc is rejected.', False),
        _field('baseline.capex', 'millions', 'article or 10-K', 'Year-0 capital expenditure. Positive is a cash outflow.', True),
        _field('baseline.da', 'millions, or a short list', 'article or 10-K', 'Depreciation. Later years are calculated from capex and ROIC, so a full horizon is not required.', True),
        _field('baseline.tax', 'millions', 'article or 10-K', 'Year-0 tax expense.', True),
        _field('baseline.interest', 'millions', 'article or 10-K', 'Year-0 interest expense. Later years are rd times prior debt.', True),
        _field('baseline.sbc', 'millions', 'article or 10-K', 'Year-0 stock-based compensation. Required even when it is zero.', True),
        _field('baseline.debt', 'millions', 'article or 10-K', 'Year-0 debt only. The debt intent builds the rest of the horizon.', True),
        _field('baseline.cash', 'millions', 'article or 10-K', 'Excess cash. The 10-K helper adds short-term investments when that fact exists; drop them when you want a tighter figure.', True),
        _field('baseline.nol', 'millions', 'article, otherwise an assumption', 'Net operating loss. Not in the 10-K helper.', True),
        _field('baseline.noa', 'millions', 'article, otherwise an assumption', 'Non-operating assets. Not in the 10-K helper.', True),
        _field('forecast.ebitda_growth', 'list of decimals', 'article guidance or forecast', 'Growth after year 0. Do not also send a full EBITDA path.', False),
        _field('forecast.ebitda', 'millions for every year', 'article forecast table', 'Explicit EBITDA including year 0. Must cover the whole horizon.', False),
        _field('forecast.margin_path', 'me, mc, gsnext', 'article margin and sales outlook', 'Optional companion to ebitda_growth. me and mc must be non-zero.', False),
        _field('forecast.capex', 'millions, short or full', 'article capex guide', 'From year 0. A short list is extended. Omit to start from baseline capex and extend that.', False),
        _field('forecast.sbc', 'millions, short or full', 'article', 'From year 0. A short list is extended toward sbc_rate_terminal, or at the last ratio.', False),
        _field('forecast.sbc_rate_terminal', 'fraction of EBITDA', 'assumption or article', 'Ignored when forecast.sbc already covers every year.', False),
        _field('forecast.dwc', 'zero or a full path', 'article working-capital comment, otherwise an assumption', 'Not extended. mode path must have year + 1 values.', True),
        _field('dividend_per_share', 'dollars, or a list', 'article dividend', 'Index 0 is the baseline year. Omit for no dividend.', False),
        _field('debt', 'hold, path, or target', 'article leverage or debt plan', 'Exactly one policy. See debt_modes.', True),
        _field('acquisitions', 'list of deals', 'article M&A', 'Empty when the article has no deals. See acquisition_size.', False),
        _field('disposals', 'list of sales', 'article asset sales', 'amount is gross proceeds in millions.', False),
        _field('distribution', 'one mode', 'article capital return', 'Exactly one policy. See distribution_modes.', True),
        _field('article.published', 'YYYY-MM-DD', 'article', 'Compared with market.price as_of. An older price is a warning.', False),
        _field('article.price_target', 'dollars per share', 'article', 'Reported next to the model value. Not used in the calculation.', False),
        _field('earnings.e', 'year-0 earnings', 'article', 'Only when path is earnings.', False),
        _field('earnings.payout', 'decimal or list', 'article', 'Payout ratio. Terminal payout is 1 - gt/roe.', False),
        _field('earnings.gf', 'growth rates', 'article', 'Earnings growth.', False),
        _field('earnings.roe', 'decimal', 'assumption or article', 'Terminal return on equity.', False),
        _field('fcfe', 'millions for every year', 'article cash-flow forecast', 'Only when path is fcfe. The year-0 value is not discounted.', False),
    ]


def _field(path, unit, source, article, required_on_ebitda):
    return {
        'path': path,
        'unit': unit,
        'source': source,
        'article': article,
        'required_on_ebitda_path': required_on_ebitda,
    }

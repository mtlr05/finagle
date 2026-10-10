'''What an agent needs to know before it fills a case.'''

from finagle.agent.schema import BASELINE_PRIORITY

VALUE_FROM_ARTICLE = '''You are valuing one publicly traded company from an article. The article is already in your context. Do not fetch it. Seeking Alpha and Substack posts are often paywalled, and no tool downloads them. The user may also attach a filing, such as an annual report or a 10-K; read it the same way.

Use the finagle tools. Do not invent a list of Python method calls. Fill a case, which is a set of intents. The runner turns those intents into the existing company methods.

1. Extract every figure the article actually states: baseline results, guidance, growth, margins, capex, buybacks, dividends, leverage, acquisitions, asset sales, share count, and the current price. Keep a short quote for each one.
2. Call describe_inputs. Match what you extracted to those fields. Note the units: money in millions of the reporting currency, shares in millions, price per share in that same currency, rates as decimals.
3. Year 0 is the last completed fiscal year, never trailing twelve months or a quarter, because EBITDA growth is measured from it. Give every reported baseline leaf basis fiscal_year, its period_end, and its currency. Leave out trailing-twelve-month figures. Growth rates from the article must be relative to that fiscal year.
4. Fill the baseline in this order: the article, then the attached filing (source attachment, evidence as page plus quote, such as "p. 47: Total revenue 1,200.4"), then the SEC 10-K, then yfinance. Call build_baseline with the article leaves and the attachment leaves; it fills the rest for the same fiscal year and reports what it skipped. Canadian companies skip the SEC: pass the TSX symbol, such as SHOP.TO or a .V symbol for the TSX Venture. Use the price build_baseline returns in the reporting currency. Read its warnings: a newer fiscal year in the data does not move year 0. Debt from any source is one year, not a forecast.
5. Keep acquired EBITDA and its purchase price together. Forecast EBITDA is one of three things: the organic business, the full-year run-rate of deals that closed in or before year 0 (already paid for, so they stay in the EBITDA path and get no MnA), or deals that close after year 0. Put only the first two in forecast.ebitda or forecast.ebitda_growth, and set forecast.ebitda_basis to organic. Each later deal is one acquisitions entry. deal_value is the enterprise value paid and multiple is the EV/EBITDA paid, so the purchase is booked as MnA in the deal year and the EBITDA starts the next year. If the article has an organic growth row, use it; acquired EBITDA is total EBITDA minus organic. If it gives total EBITDA and a deal-spend row, infer the multiple as cumulative spend divided by acquired EBITDA. If it implies future M&A but states no spend, assume the spend and the multiple and label both as assumptions. A deal that closes during year 1 mixes the full-year run-rate of deals already paid for with new deal EBITDA: keep the run-rate in the organic path and put only the new deal in acquisitions. Copy every year of a deal-spend row that falls inside the horizon into deal_spend, including year 1. A positive number in that row is a future deal, not the run-rate of a deal that already closed. The run-rate is only the EBITDA that shows up with no spend attached. Do not leave that year at zero. Set forecast.acquisition_disclosure status to future_deals and give deal_spend, in millions by year, as the same dollars as deal_value (those dollars become MnA). Set status to none, with a quote showing the forecast has no future deals, only when that is true. Do not put acquired EBITDA in the path and also list the deal. Set ebitda_basis to total only when the article's EBITDA forecast already includes acquired EBITDA and there are no future deals. If the article states a total EBITDA path, copy it to article.ebitda_total so the result can reconcile it. That path is not an input.
6. Anything still empty is an assumption. Set source to assumption and write the reason in evidence. Do not leave a required field out, and do not silently use zero for working capital, net operating losses, non-operating assets, or stock-based compensation.
7. Call validate_case. Fix every error and every missing field. Warnings stay in the result; read them before you run.
8. Call run_case.
9. Report value per share from the free-cash-flow-to-equity model and from the dividend discount model, the terminal free cash flow to equity, and the assumptions (every leaf whose source is assumption). If the result has a log error, or ok is false, the run failed: report the error and do not present the number as a valuation. Place the article's price target beside the result when the article states one. Read acquisitions_summary, the MnA column, and ebitda_reconciliation before you report the value. If any reconciliation gap is wider than 5%, the organic path is wrong: change it and run again. A deal paid in year t adds EBITDA starting in year t+1, so organic EBITDA in that later year is the article's total minus the EBITDA those deals add.

EBITDA must be before stock-based compensation. The model subtracts stock-based compensation, and the SEC and yfinance baselines already add it back. If the article's EBITDA is after stock-based compensation, add it back and set ebitda_definition to before_sbc.

Do not tune any input so the model matches the article's price target. The price target is not an input. It belongs only in article.price_target.
'''


def describe_inputs():
    '''Field catalog, units, and the article cues that map onto intents.'''
    return {
        'units': {
            'money': 'millions of the reporting currency (the currency of the financial statements)',
            'shares': 'millions of shares',
            'price': 'per share, in the reporting currency',
            'rates': 'decimals; 0.09 means 9 percent',
        },
        'year_index': (
            'Index 0 is the baseline period. Index 1 is the first forecast year. '
            'year is the last explicit year, so a full series has year + 1 values. '
            'The baseline is the last completed fiscal year, never trailing twelve '
            'months, because EBITDA growth is measured from it.'
        ),
        'baseline_sources': {
            'priority': list(BASELINE_PRIORITY),
            'rule': (
                'Each year-0 field comes from the first source that has it for the '
                'baseline fiscal year: the article, then an attached filing, then the '
                'SEC 10-K, then yfinance. build_baseline does the merge.'
            ),
            'year': (
                'The baseline year is the article\'s fiscal year when it states one, '
                'else the attachment\'s, else the latest 10-K, else the latest yfinance '
                'annual statement. Lower sources only fill that same year. A newer '
                'fiscal year in the data is a warning, not a reason to move year 0.'
            ),
            'reported_leaves': (
                'Every reported baseline leaf needs basis fiscal_year, a period_end '
                'equal to baseline.date, and the reporting currency. ttm and quarter '
                'are rejected.'
            ),
            'attachment': (
                'A filing the user attached. Use source attachment and evidence as '
                'page plus quote, for example "p. 47: Total revenue 1,200.4".'
            ),
            'canadian': (
                'A ticker ending in .TO, .V, .CN, or .NE, or country CA, never calls '
                'the SEC. Pass the TSX symbol yfinance uses, such as SHOP.TO.'
            ),
            'yahoo': (
                'Optional: pip install -e ".[yahoo]". Unofficial and for personal use. '
                'EBITDA is rebuilt as operating income plus D&A plus SBC. Total Debt can '
                'include leases. A listing that trades in another currency gets a '
                'converted price.'
            ),
        },
        'ebitda_definition': (
            'EBITDA must be before stock-based compensation. The model subtracts sbc. '
            'The SEC and yfinance helpers add sbc back into EBITDA. If an article\'s EBITDA is '
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
            'none': (
                'Pay all free cash flow to equity out as dividends in the year it is '
                'earned, with no buybacks. The two values per share come out equal.'
            ),
            'max_buybacks': 'Spend available cash on repurchases.',
            'buyback_schedule': (
                'Repurchase the dollar amounts given, in millions. Index 0 is the '
                'baseline year and is usually 0 when the repurchase is in the future. '
                'A 0 keeps that year\'s cash on the balance sheet. After the list ends, '
                'the last amount is repeated in proportion to free cash flow, so end '
                'the list with 0 when the program stops: three years of $100 million '
                'is [0, 100, 100, 100, 0]. This is not "buy back as much as possible".'
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
            'A buyback authorization or a dollar repurchase program becomes distribution mode buyback_schedule. End the schedule with 0 when the program has an end date; otherwise the last amount continues every year. Language that all free cash flow will be used for buybacks becomes max_buybacks.',
            'A dividend per share becomes dividend_per_share. After the list ends, the total dividend grows with free cash flow and is never cut.',
            'A bolt-on acquisition becomes one acquisitions entry. Use the deal value and the multiple, or the acquired EBITDA. The EBITDA path stays organic: pay in the deal year, and the acquired EBITDA starts the next year.',
            'An organic growth row, or total EBITDA minus organic, is the EBITDA path. A deal-spend or acquisition-consideration row is forecast.acquisition_disclosure.deal_spend and the acquisitions entries. Do not leave MnA at zero while that EBITDA is in the path.',
            'An asset sale or a stake sale becomes one disposals entry.',
            'Excess cash, net operating losses, and non-operating assets become cash, nol, and noa. Use 0 only as an explicit assumption.',
            'Results for the last fiscal year become the baseline, with basis fiscal_year. Trailing-twelve-month or quarterly results are not year 0; leave them out and let a lower source fill the fiscal year.',
            'The current share price becomes market.price, with as_of set to the date of that price and the reporting currency. The author\'s price target becomes article.price_target and nothing else.',
            'If the article does not give a discount rate, terminal growth, or terminal ROIC, those are assumptions. Say what you used and why.',
        ],
        'do_not': [
            'Do not tune inputs so value_per_share matches the article\'s price target.',
            'Do not put the price target anywhere except article.price_target.',
            'Do not send a list of company method calls. Send the intents in the case.',
            'Do not treat baseline debt as a full-horizon series. It is year 0 only.',
            'Do not use trailing twelve months or a quarter as year 0. Use the last completed fiscal year.',
            'Do not mix fiscal years or currencies in the baseline.',
            'Do not put EBITDA that includes future acquisitions into the forecast and leave MnA at zero. Split organic EBITDA from the deals and book the purchase price with acquisitions.',
            'Do not leave a year at zero in deal_spend when the article\'s deal-spend row is positive that year, including year 1.',
            'Do not set ebitda_basis to total and also list acquisitions. That counts the acquired EBITDA twice.',
            'Do not call the SEC for a Canadian company.',
            'Do not call display_fin. run_case returns the numbers.',
        ],
    }


def _fields():
    return [
        _field('ticker', 'string', 'Article or the user', 'The company the article is about.', True),
        _field('year', 'integer', 'assumption', 'How many years to forecast before the terminal year. 10 is a common assumption.', True),
        _field('path', 'ebitda, earnings, or fcfe', 'assumption', 'Omit to use the EBITDA path.', False),
        _field('market.shares', 'millions of shares', 'article, attachment, 10-K, or yfinance', 'All classes. build_baseline and the 10-K helper return this beside the baseline, not inside it.', True),
        _field('market.price', 'per share, reporting currency', 'article or build_baseline (yfinance)', 'Current price, with as_of and currency. A listing in another currency must be converted. Required for buybacks when distribution.price is omitted.', False),
        _field('rates.re', 'decimal', 'assumption, unless the article builds a discount rate', 'Cost of equity. Must be greater than gt.', True),
        _field('rates.rd', 'decimal', 'assumption or the article\'s stated cost of debt', 'Cost of debt. Required on the EBITDA path.', True),
        _field('rates.t', 'decimal', 'assumption; 0.21 is the US federal statutory rate', 'Marginal tax rate.', True),
        _field('rates.te', 'decimal', 'article, when year-1 cash tax is not the marginal rate', 'Optional effective tax rate for year 1.', False),
        _field('rates.gt', 'decimal', 'assumption', 'Terminal growth. Must be below re.', True),
        _field('rates.roict', 'decimal', 'assumption', 'Terminal return on invested capital, used for terminal depreciation.', True),
        _field('baseline.date', 'YYYY-MM-DD', 'fiscal year end from the article, attachment, 10-K, or yfinance', 'The last completed fiscal year, year index 0. Every reported baseline leaf has this period_end.', True),
        _field('baseline.ebitda', 'millions', 'article, attachment, 10-K, then yfinance', 'Year-0 EBITDA before stock-based compensation.', True),
        _field('baseline.ebitda_definition', 'before_sbc or after_sbc', 'same source as baseline.ebitda', 'Record which EBITDA you used. after_sbc is rejected.', False),
        _field('baseline.capex', 'millions', 'article, attachment, 10-K, then yfinance', 'Year-0 capital expenditure. Positive is a cash outflow.', True),
        _field('baseline.da', 'millions, or a short list', 'article, attachment, 10-K, then yfinance', 'Depreciation. Later years are calculated from capex and ROIC, so a full horizon is not required.', True),
        _field('baseline.tax', 'millions', 'article, attachment, 10-K, then yfinance', 'Year-0 tax expense.', True),
        _field('baseline.interest', 'millions', 'article, attachment, 10-K, then yfinance', 'Year-0 interest expense. Later years are rd times prior debt.', True),
        _field('baseline.sbc', 'millions', 'article, attachment, 10-K, then yfinance', 'Year-0 stock-based compensation. Required even when it is zero.', True),
        _field('baseline.debt', 'millions', 'article, attachment, 10-K, then yfinance', 'Year-0 debt only. The debt intent builds the rest of the horizon. Yahoo Total Debt can include leases.', True),
        _field('baseline.cash', 'millions', 'article, attachment, 10-K, then yfinance', 'Excess cash. The 10-K and yfinance helpers include short-term investments when they exist; drop them when you want a tighter figure.', True),
        _field('baseline.nol', 'millions', 'article or attachment, otherwise an assumption', 'Net operating loss. Not in the SEC or yfinance helpers.', True),
        _field('baseline.noa', 'millions', 'article or attachment, otherwise an assumption', 'Non-operating assets. Not in the SEC or yfinance helpers.', True),
        _field('forecast.ebitda_basis', 'organic or total', 'article forecast', 'organic excludes deals after year 0. total is only valid when the forecast has no future deals.', True),
        _field('forecast.acquisition_disclosure.status', 'future_deals or none', 'article M&A forecast', 'future_deals when the forecast buys EBITDA after year 0. none only with a quote that it does not.', True),
        _field('forecast.acquisition_disclosure.deal_spend', 'millions by year', 'article deal spend, or an assumption', 'Enterprise value paid, same dollars as deal_value. Those dollars become MnA. Required when status is future_deals.', False),
        _field('forecast.ebitda_growth', 'list of decimals', 'article guidance or forecast', 'Organic growth after year 0, plus the run-rate of deals already paid for. Do not also send a full EBITDA path.', False),
        _field('forecast.ebitda', 'millions for every year', 'article forecast table', 'Explicit EBITDA including year 0, on the basis in ebitda_basis. Must cover the whole horizon.', False),
        _field('forecast.margin_path', 'me, mc, gsnext', 'article margin and sales outlook', 'Optional companion to ebitda_growth. me and mc must be non-zero.', False),
        _field('forecast.capex', 'millions, short or full', 'article capex guide', 'From year 0. A short list is extended. Omit to start from baseline capex and extend that.', False),
        _field('forecast.sbc', 'millions, short or full', 'article', 'From year 0. A short list is extended toward sbc_rate_terminal, or at the last ratio.', False),
        _field('forecast.sbc_rate_terminal', 'fraction of EBITDA', 'assumption or article', 'Ignored when forecast.sbc already covers every year.', False),
        _field('forecast.dwc', 'zero or a full path', 'article working-capital comment, otherwise an assumption', 'Not extended. mode path must have year + 1 values.', True),
        _field('dividend_per_share', 'dollars, or a list', 'article dividend', 'Index 0 is the baseline year. After the list ends, the total dividend grows with free cash flow and is never cut. Omit for no dividend.', False),
        _field('debt', 'hold, path, or target', 'article leverage or debt plan', 'Exactly one policy. See debt_modes.', True),
        _field('acquisitions', 'list of deals', 'article M&A', 'One entry per deal that closes after the organic path. Empty only when acquisition_disclosure status is none. See acquisition_size.', False),
        _field('disposals', 'list of sales', 'article asset sales', 'amount is gross proceeds in millions.', False),
        _field('distribution', 'one mode', 'article capital return', 'Exactly one policy. See distribution_modes.', True),
        _field('article.published', 'YYYY-MM-DD', 'article', 'Compared with market.price as_of. An older price is a warning.', False),
        _field('article.price_target', 'per share', 'article', 'Reported next to the model value. Not used in the calculation.', False),
        _field('article.ebitda_total', 'millions by year', 'article forecast table', 'The article\'s total EBITDA, including acquired EBITDA. Reconciled with the model. Not an input.', False),
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

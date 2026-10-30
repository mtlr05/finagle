# Agent instructions

This repository values publicly traded companies. People use the `company` class in `finagle/company.py` from notebooks. Agents use `finagle.agent`, which runs that same class without changing it.

## Valuing a company from an article

The workflow is `VALUE_FROM_ARTICLE` in [`finagle/agent/catalog.py`](finagle/agent/catalog.py). Read it first and follow it. This file only points to it.

Set up from the repository root:

```bash
pip install -e .
pip install -e ".[yahoo]"   # optional: yfinance, needed for Canadian companies
```

Then use these functions from `finagle.agent`, in this order:

1. `describe_inputs()` lists every field, its unit, and the article wording that maps onto it.
2. `read_pdf(path, out_dir, name=...)` on every PDF. Use `name="article"` for the primary article and the attachment name for each other file. The extract includes each page's text and, for every image that covers at least 8 percent of the page, a chart block: OCR digits, or an `UNREAD CHART` line with the png to open. Fill leaves from that extract, including the chart block.
3. `build_baseline(ticker, prompt=..., article=..., attachments=..., documents=..., user_agent=...)` fills the case. Pass the leaves the user stated in the request, the leaves you read from the primary article, the other attachments in the order the request gives, and the `read_pdf` results as `documents`. It fills the remaining statement fields from the SEC 10-K, then yfinance, for the same fiscal year, then fills the default assumptions. It returns a partial case, which source supplied each field, what it skipped, and warnings. If `ok` is false and `unread_charts` is set, cite each page (`p. N` plus `chart`, `OCR`, `read from axis`, or `no numeric labels`) and call it again. Do not continue while `unread_charts` is set. The SEC requires a user agent with a name and an email address. Use the one the user gives you, or the `FINAGLE_SEC_USER_AGENT` environment variable.
4. `validate_case(case)` returns errors, missing fields, and warnings. Fix every error and missing field.
5. `run_case(case)` compiles the case into `company` methods, runs `value()`, and returns the result, including the case.
6. `update_case(case, updates)` applies a follow-up. Pass only the leaves that changed. It does not reread the documents or call the SEC or yfinance.

Build each figure with `finagle.agent.schema.leaf(value, source, evidence, period_end=..., basis=..., currency=...)`. `source` is `prompt` (the user's own figure), `article` (with a short quote), `attachment` (page plus quote, such as `"p. 47: Total revenue 1,200.4"`), `10k`, `yahoo`, or `assumption` (with a reason).

## Year 0

- Year 0 is the last completed fiscal year. Never use trailing twelve months or a quarter, because EBITDA growth is measured from year 0. Article growth rates must be relative to that fiscal year.
- Each field comes from the first of these that has it for that year: a figure stated in the request, the primary article, other attachments in the order the request gives, the SEC 10-K, then yfinance.
- Every reported baseline leaf needs `basis: "fiscal_year"`, a `period_end` equal to `baseline.date`, and the same `currency`. `market.price` must be in that currency too; use the converted price `build_baseline` returns.
- Canadian companies skip the SEC. Use the TSX symbol yfinance uses, such as `SHOP.TO`, or a `.V` symbol for the TSX Venture.
- `nol` and `noa` are not in the SEC or yfinance data. Take them from the article or the attachment, or label them as assumptions.
- yfinance is unofficial, Yahoo limits its data to personal use, and its row names change. Check its figures against a filing when they matter. Its `Total Debt` can include leases.
- `get_10k_baseline(ticker, user_agent)` is still available when you only want the SEC figures.

## Defaults

When no source states them, `build_baseline` fills:

- `me` = year-0 EBITDA / revenue, and `mc` = the average of `me` and the gross margin (`mc` equals `me` when gross margin is missing). Derived EBITDA growth rolls that margin forward: each year's sales growth is multiplied by `mc` divided by that year's EBITDA margin. The `me` passed to `forecast_ebitda` is EBITDA divided by revenue at the end of the explicit growth list when that revenue is known. A stated EBITDA growth rate is kept. Sales growth is turned into EBITDA growth only when no EBITDA growth was stated. When the article's sales forecast runs longer than its EBITDA dollars, leave `forecast.ebitda` unset. Impute `forecast.ebitda_growth` from those dollars through the last printed year, pass sales growth over that same window, and set `gsnext` to the sales growth stated after it. `article.ebitda_total` lists only the printed EBITDA years.
- `re` = 0.055 + the 10-year yield of the reporting currency. `gt` equals that yield, and a higher `gt` is an error. `roict` = 0.15.
- The price is the last trading day's close, in the reporting currency.
- With no debt policy, the target Debt/EBITDA is the year-0 ratio, starting in year 1.
- A buyback uses a proportional price path.
- `year` = 10.
- `t` = 0.21, the US federal statutory rate.
- `rd` = year-0 interest / year-0 debt, when debt is not zero. A blank Yahoo `Interest Expense` uses signed net interest income, flipped so an expense is positive.
- `nol` = 0 and `noa` = 0. Neither is in the SEC or yfinance feeds.
- Working capital does not change (`forecast.dwc.mode` = `zero`).
- Distribution mode is `none` (free cash flow to equity is paid out, with no buybacks).

When no source states a capex or stock-based-compensation path, leave `forecast.capex` and `forecast.sbc` unset. Do not repeat the year-0 amount across the horizon. `run_case` then calls `forecast_sbc` and `forecast_capex`. If a case still assumes a full capex or SBC path, or repeats the year-0 amount, `run_case` cuts it back to year 0, warns, and runs `forecast_capex` and `forecast_sbc`. When revenue is forecast, capex method 1 keeps the last supplied depreciation-to-revenue ratio, solves terminal capex from the ROIC identity, and bridges the open years with revenue. Otherwise capex holds the last capex-to-EBITDA ratio. SBC holds the last SBC-to-EBITDA ratio unless a terminal rate is set. A stated path is still a short or full list, and a full list is kept.

## A follow-up

Use `update_case` with the case `run_case` returned and only the changed leaves. Do not reread the article or the attachments, and do not call `build_baseline` again.

## Negative depreciation

Terminal depreciation or terminal capex can come out negative. The valuation still completes. The log carries a warning. Report the warning and still present the value. Other log errors still mean the run failed.

[`examples/article_case.json`](examples/article_case.json) is a complete case built from [`examples/sample_article.md`](examples/sample_article.md), a fictional article. Use it as the model for the case shape.

## Rules

- Forecast EBITDA is organic unless `forecast.ebitda_basis` says `total`. Deals that close after year 0 go in `acquisitions`, with the purchase price, so `MnA` is not zero. Say which in `forecast.acquisition_disclosure`. Do not put acquired EBITDA in the EBITDA path and also in `acquisitions`. A positive year on a deal-spend row, including the first forecast year, is a future deal: do not leave that year at zero in `deal_spend`.
- The valuation must come from `run_case`. Do not calculate it yourself.
- Do not edit `finagle/company.py`, `finagle/sec.py`, or the notebooks to get a valuation.
- Do not change inputs to match the article's price target. The price target goes only in `article.price_target`.
- If `ok` is false, the run failed. Report the error and do not present a value.

## What to report

1. `value_per_share` (free cash flow to equity) and `value_per_share_DDM` (dividend discount)
2. Every input whose source is `assumption`, with its reason
3. Validation warnings, log errors, and `sanity_flags`
4. The article's price target, next to the model value
5. `notebook_code`, so a person can rerun the same case in a notebook

## Other ways to connect

An agent that cannot run Python here can use the MCP server, which exposes the same functions, plus `get_10k_baseline` and `update_case`, and the workflow as the `value_from_article` prompt. It needs Python 3.10 or later:

```bash
pip install -e ".[agent]"
python -m finagle.agent.mcp_server --transport stdio
```

Use `--transport streamable-http` for a remote client.

## Tests

`pytest` from the repository root. CI runs Python 3.9.12. The MCP tests are skipped when `mcp` is not installed. The source tests use fake SEC and yfinance clients and need no network.

## Collaborative request

The article workflow above is unchanged. This section applies only when the request has `"mode": "collaborative"`, or the user asked for a joint valuation with a Forecaster operating path.

- The inputs are the Forecaster's sales growth, EBITDA growth, `me`, `mc`, `gsnext`, EBITDA basis, and any acquisitions in the forecast, plus `prompt_inputs`. `prompt_inputs` are only leaves the user stated, as explicit values, with the user's words and date. Source is `prompt`.
- Do not treat the request as an article valuation. Do not fill a leaf from a PDF or from a default when the request already has it.
- `build_baseline` still supplies price, shares, the risk-free rate, the cost of equity, terminal growth, the year-0 figures, and the unstated assumptions in Defaults (horizon, tax rate, cost of debt, tax losses, non-operating assets, working capital, distribution, and the debt target), unless a `prompt_inputs` item states one.
- Apply the request with `update_case` after `build_baseline`. Forecaster leaves first, then `prompt_inputs`. `build_baseline` drops a prompt leaf that has no fiscal-year basis, so these leaves go on afterwards.
- Report every leaf filled by default (`source` `assumption`). Do not apply a post-run factor. Do not write the company report. The Forecaster does that from the result file.
- A Grok Bot fast-forwards its own clone to `origin/main` before the run and records the commit. If the fast-forward fails, stop. Do not reset. A Cursor session does not pull. It uses the checkout it was given and records that commit.
- Do not invent a stock-comp path, and do not replace those labeled defaults with a different horizon, tax rate, cost of debt, working-capital rule, or payout rule. If the user stated a formula rather than numbers, the coordinator turns it into explicit values before the run. A full list whose source is `assumption` is cut back to year 0, so a list the user stated uses source `prompt`.

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
2. `build_baseline(ticker, prompt=..., article=..., attachments=..., user_agent=...)` fills the case. Pass the leaves the user stated in the request, the leaves you read from the primary article, and the other attachments in the order the request gives. It fills the remaining statement fields from the SEC 10-K, then yfinance, for the same fiscal year, then fills the default assumptions. It returns a partial case, which source supplied each field, what it skipped, and warnings. The SEC requires a user agent with a name and an email address. Use the one the user gives you, or the `FINAGLE_SEC_USER_AGENT` environment variable.
3. `validate_case(case)` returns errors, missing fields, and warnings. Fix every error and missing field.
4. `run_case(case)` compiles the case into `company` methods, runs `value()`, and returns the result, including the case.
5. `update_case(case, updates)` applies a follow-up. Pass only the leaves that changed. It does not reread the documents or call the SEC or yfinance.

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

- `me` = year-0 EBITDA / revenue, and `mc` = the average of `me` and the gross margin (`mc` equals `me` when gross margin is missing). The forecast uses these so the years after any explicit growth list are driven by sales. A stated EBITDA growth rate is kept. Sales growth is turned into EBITDA growth only when no EBITDA growth was stated.
- `re` = 0.055 + the 10-year yield of the reporting currency. `gt` equals that yield, and a higher `gt` is an error. `roict` = 0.15.
- The price is the last trading day's close, in the reporting currency.
- With no debt policy, the target Debt/EBITDA is the year-0 ratio, starting in year 1.
- A buyback uses a proportional price path.

When no source states a capex or stock-based-compensation path, leave `forecast.capex` and `forecast.sbc` unset. Do not repeat the year-0 amount across the horizon. `run_case` then calls `forecast_capex` and `forecast_sbc`, which extend those series at the last ratio to EBITDA. A stated path is still a short or full list, and a full list is kept.

Cost of debt, the tax rate, working capital, net operating losses, non-operating assets, and the distribution mode are still for you to fill.

## A follow-up

Use `update_case` with the case `run_case` returned and only the changed leaves. Do not reread the article or the attachments, and do not call `build_baseline` again.

## Negative depreciation

Terminal depreciation can come out negative. The valuation still completes. The log carries a warning. Report the warning and still present the value. Other log errors still mean the run failed.

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

# Agent instructions

This repository values publicly traded companies. People use the `company` class in `finagle/company.py` from notebooks. Agents use `finagle.agent`, which runs that same class without changing it.

## Valuing a company from an article

The workflow is `VALUE_FROM_ARTICLE` in [`finagle/agent/catalog.py`](finagle/agent/catalog.py). Read it first and follow it. This file only points to it.

Set up from the repository root:

```bash
pip install -e .
```

Then use these four functions from `finagle.agent`, in this order:

1. `describe_inputs()` lists every field, its unit, and the article wording that maps onto it.
2. `get_10k_baseline(ticker, user_agent)` fills baseline figures the article does not give. The SEC requires a user agent with a name and an email address. Use the one the user gives you, or the `FINAGLE_SEC_USER_AGENT` environment variable.
3. `validate_case(case)` returns errors, missing fields, and warnings. Fix every error and missing field.
4. `run_case(case)` compiles the case into `company` methods, runs `value()`, and returns the result.

Build each figure with `finagle.agent.schema.leaf(value, source, evidence)`. `source` is `article` (with a short quote), `10k`, or `assumption` (with a reason).

[`examples/article_case.json`](examples/article_case.json) is a complete case built from [`examples/sample_article.md`](examples/sample_article.md), a fictional article. Use it as the model for the case shape.

## Rules

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

An agent that cannot run Python here can use the MCP server, which exposes the same four functions and the workflow as the `value_from_article` prompt. It needs Python 3.10 or later:

```bash
pip install -e ".[agent]"
python -m finagle.agent.mcp_server --transport stdio
```

Use `--transport streamable-http` for a remote client.

## Tests

`pytest` from the repository root. CI runs Python 3.9.12. The MCP tests are skipped when `mcp` is not installed.

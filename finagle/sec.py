"""Year-0 baseline figures from the latest 10-K on EDGAR.

This module only reads SEC data. It does not load a company or write
``self.fin``. Copy accepted values into the financials dict yourself.
"""

import json
import urllib.request
from datetime import datetime

import pandas as pd

TICKER_URL = "https://www.sec.gov/files/company_tickers.json"
SUBMISSIONS_URL = "https://data.sec.gov/submissions/CIK{cik}.json"
FACTS_URL = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"

# A fiscal year in the company-facts feed is the number of days between the
# reported start and end. Calendar years land on 364 or 365; 52/53-week years
# land on 363 to 371.
FLOW_MIN_DAYS = 350
FLOW_MAX_DAYS = 380

KEYS = (
    "date",
    "revenue",
    "tax",
    "interest",
    "da",
    "capex",
    "sbc",
    "cash",
    "debt",
    "ebitda",
)

_LIST_KEYS = frozenset(
    ("revenue", "tax", "interest", "da", "capex", "sbc", "ebitda", "debt")
)

_REVENUE_TAGS = (
    "RevenueFromContractWithCustomerExcludingAssessedTax",
    "Revenues",
    "SalesRevenueNet",
)
_INTEREST_TAGS = ("InterestExpense", "InterestExpenseDebt")
_CAPEX_TAGS = (
    "PaymentsToAcquirePropertyPlantAndEquipment",
    "PaymentsToAcquireProductiveAssets",
)
_SBC_TAGS = (
    "ShareBasedCompensation",
    "AllocatedShareBasedCompensationExpense",
)
_SHORT_DEBT_TAGS = (
    "LongTermDebtCurrent",
    "ShortTermBorrowings",
    "CommercialPaper",
)


class Baseline:
    """Year-0 figures from one 10-K, plus the tag each figure came from.

    ``provenance`` has one row per key, including keys that were not found.
    ``financials`` contains only the keys that resolved, shaped for year 0
    of the company model: ``date`` is a string, ``cash`` is a number, and
    the other keys are one-element lists. ``debt`` is that one year only.
    """

    def __init__(self, provenance, financials):
        self.provenance = provenance
        self.financials = financials


def last_10k(ticker, user_agent, scale=1_000_000):
    """Pull year-0 figures from the company's latest 10-K.

    ``user_agent`` is required. The SEC rejects requests that do not
    identify the caller; use a string with a name and an email address.

    ``scale`` divides dollar amounts. The default, one million, matches
    the units used in the valuation notebooks. ``provenance`` keeps the
    unscaled dollar amount beside the scaled value.

    The period is the latest 10-K fiscal year. That matches trailing
    twelve months only while that 10-K is still the latest report.
    """
    if user_agent is None or not str(user_agent).strip():
        raise ValueError("user_agent is required")
    if not scale:
        raise ValueError("scale must be non-zero")
    ticker = str(ticker).strip()
    if not ticker:
        raise ValueError("ticker is required")

    cik = _cik_for_ticker(_get_json(TICKER_URL, user_agent), ticker)
    submissions = _get_json(SUBMISSIONS_URL.format(cik=cik), user_agent)
    period_end = _latest_10k_period(submissions, ticker)
    facts = _get_json(FACTS_URL.format(cik=cik), user_agent)
    return baseline_from_facts(facts, period_end, scale=scale)


def baseline_from_facts(facts, period_end, scale=1_000_000):
    """Map one company-facts document onto year-0 figures.

    ``facts`` is the JSON from the SEC company-facts endpoint. ``period_end``
    is the 10-K report date, ``YYYY-MM-DD``. Flow facts must be a fiscal-year
    duration ending that day. Stock facts must be the instant on that day.
    The first matching tag in each preferred list is the one that is used.
    """
    if not scale:
        raise ValueError("scale must be non-zero")
    usgaap = (facts or {}).get("facts", {}).get("us-gaap", {})

    chosen = {}
    chosen["date"] = {
        "value": period_end,
        "raw": None,
        "tags": "",
        "note": "Period end of the latest 10-K.",
    }

    revenue = _first(usgaap, _REVENUE_TAGS, period_end, "flow")
    chosen["revenue"] = _single(revenue, "Revenue was not found in the latest 10-K.")

    tax = _first(usgaap, ("IncomeTaxExpenseBenefit",), period_end, "flow")
    chosen["tax"] = _single(tax, "Income tax expense was not found in the latest 10-K.")

    interest = _first(usgaap, _INTEREST_TAGS, period_end, "flow")
    chosen["interest"] = _single(
        interest,
        "Interest expense was not found in the latest 10-K. Interest income is not used.",
    )

    da = _depreciation(usgaap, period_end)
    chosen["da"] = da

    capex = _first(usgaap, _CAPEX_TAGS, period_end, "flow")
    if capex is None:
        chosen["capex"] = _missing("Capital expenditure was not found in the latest 10-K.")
    else:
        raw = abs(capex["val"])
        chosen["capex"] = {
            "value": raw / scale,
            "raw": raw,
            "tags": capex["tag"],
            "note": "Positive cash outflow.",
        }

    sbc = _first(usgaap, _SBC_TAGS, period_end, "flow")
    chosen["sbc"] = _single(sbc, "Stock-based compensation was not found in the latest 10-K.")

    chosen["cash"] = _cash(usgaap, period_end)
    chosen["debt"] = _debt(usgaap, period_end)
    chosen["ebitda"] = _ebitda(usgaap, period_end, da, chosen["sbc"])

    rows = []
    financials = {}
    for key in KEYS:
        item = chosen[key]
        raw = item["raw"]
        scaled = None if raw is None else raw / scale
        if key == "date":
            display = period_end
        else:
            display = scaled
        rows.append({
            "key": key,
            "value": display,
            "tags": item["tags"],
            "raw dollars": raw,
            "period end": period_end,
            "note": item["note"],
        })
        if item["raw"] is None and key != "date":
            continue
        if key == "date":
            financials["date"] = period_end
        elif key == "cash":
            financials["cash"] = scaled
        elif key in _LIST_KEYS:
            financials[key] = [scaled]

    provenance = pd.DataFrame(rows, columns=[
        "key", "value", "tags", "raw dollars", "period end", "note",
    ])
    return Baseline(provenance, financials)


def _single(found, missing_note):
    if found is None:
        return _missing(missing_note)
    return {
        "value": found["val"],
        "raw": found["val"],
        "tags": found["tag"],
        "note": "",
    }


def _missing(note):
    return {"value": None, "raw": None, "tags": "", "note": note}


def _depreciation(usgaap, period_end):
    total = _best(
        usgaap, "DepreciationDepletionAndAmortization", period_end, "flow")
    if total is not None:
        return {
            "value": total["val"],
            "raw": total["val"],
            "tags": total["tag"],
            "note": "",
        }

    pieces = []
    combined = _best(usgaap, "DepreciationAndAmortization", period_end, "flow")
    if combined is not None:
        pieces.append(combined)
    else:
        depreciation = _best(usgaap, "Depreciation", period_end, "flow")
        if depreciation is not None:
            pieces.append(depreciation)
    intangible = _best(usgaap, "AmortizationOfIntangibleAssets", period_end, "flow")
    if intangible is not None and (
            combined is None or intangible["val"] > combined["val"]):
        pieces.append(intangible)
    if not pieces:
        return _missing(
            "Depreciation and amortization were not found in the latest 10-K."
        )
    raw = sum(piece["val"] for piece in pieces)
    note = ""
    if len(pieces) > 1:
        note = "Sum of separate depreciation and amortization lines."
    return {
        "value": raw,
        "raw": raw,
        "tags": ", ".join(piece["tag"] for piece in pieces),
        "note": note,
    }


def _cash(usgaap, period_end):
    cash = _best(
        usgaap, "CashAndCashEquivalentsAtCarryingValue", period_end, "instant")
    if cash is None:
        return _missing("Cash was not found in the latest 10-K.")
    investments = _best(usgaap, "ShortTermInvestments", period_end, "instant")
    if investments is None:
        return {
            "value": cash["val"],
            "raw": cash["val"],
            "tags": "CashAndCashEquivalentsAtCarryingValue",
            "note": "Cash and cash equivalents. Short-term investments were not found.",
        }
    raw = cash["val"] + investments["val"]
    return {
        "value": raw,
        "raw": raw,
        "tags": "CashAndCashEquivalentsAtCarryingValue, ShortTermInvestments",
        "note": "Cash plus short-term investments. Drop this if you want a tighter excess-cash figure.",
    }


def _debt(usgaap, period_end):
    noncurrent = _best(usgaap, "LongTermDebtNoncurrent", period_end, "instant")
    pieces = []
    if noncurrent is not None:
        pieces.append(noncurrent)
        current = _best(usgaap, "DebtCurrent", period_end, "instant")
        if current is not None:
            pieces.append(current)
        else:
            pieces.extend(_present(usgaap, _SHORT_DEBT_TAGS, period_end))
        note = "Short-term plus long-term debt."
    else:
        long_term = _best(usgaap, "LongTermDebt", period_end, "instant")
        if long_term is not None:
            pieces.append(long_term)
            # LongTermDebt often already includes the current portion.
            pieces.extend(_present(
                usgaap, ("ShortTermBorrowings", "CommercialPaper"), period_end))
            note = (
                "LongTermDebt plus short-term borrowings and commercial paper. "
                "DebtCurrent and LongTermDebtCurrent were not added, because they "
                "are often already inside LongTermDebt."
            )
        else:
            current = _best(usgaap, "DebtCurrent", period_end, "instant")
            if current is not None:
                pieces.append(current)
            else:
                pieces.extend(_present(usgaap, _SHORT_DEBT_TAGS, period_end))
            note = "Long-term debt was not found. The total is the short-term piece only."
    if not pieces:
        return _missing("Debt was not found in the latest 10-K.")
    raw = sum(piece["val"] for piece in pieces)
    return {
        "value": raw,
        "raw": raw,
        "tags": ", ".join(piece["tag"] for piece in pieces),
        "note": note,
    }


def _ebitda(usgaap, period_end, da, sbc):
    operating = _best(usgaap, "OperatingIncomeLoss", period_end, "flow")
    if operating is None:
        return _missing(
            "Operating income was not found, so EBITDA was not calculated."
        )
    if da["raw"] is None:
        return _missing(
            "Depreciation and amortization were not found, so EBITDA was not calculated."
        )
    raw = operating["val"] + da["raw"]
    tags = ["OperatingIncomeLoss", da["tags"]]
    if sbc["raw"] is None:
        note = (
            "Operating income plus depreciation and amortization. "
            "Stock-based compensation was not found, so it was not added."
        )
    else:
        raw += sbc["raw"]
        tags.append(sbc["tags"])
        note = (
            "Operating income plus depreciation and amortization "
            "plus stock-based compensation."
        )
    return {
        "value": raw,
        "raw": raw,
        "tags": ", ".join(tag for tag in tags if tag),
        "note": note,
    }


def _present(usgaap, tags, period_end):
    found = []
    for tag in tags:
        fact = _best(usgaap, tag, period_end, "instant")
        if fact is not None:
            found.append(fact)
    return found


def _first(usgaap, tags, period_end, kind):
    for tag in tags:
        fact = _best(usgaap, tag, period_end, kind)
        if fact is not None:
            return fact
    return None


def _best(usgaap, tag, period_end, kind):
    concept = usgaap.get(tag)
    if not concept:
        return None
    matches = []
    for fact in concept.get("units", {}).get("USD", []):
        if _matches(fact, period_end, kind):
            matches.append(fact)
    if not matches:
        return None
    chosen = max(matches, key=lambda fact: (fact.get("filed") or "", fact.get("accn") or ""))
    return {"tag": tag, "val": chosen["val"]}


def _matches(fact, period_end, kind):
    if fact.get("form") != "10-K" or fact.get("fp") != "FY":
        return False
    if fact.get("end") != period_end:
        return False
    start = fact.get("start")
    if kind == "instant":
        return not start
    if not start:
        return False
    days = (
        datetime.strptime(fact["end"], "%Y-%m-%d")
        - datetime.strptime(start, "%Y-%m-%d")
    ).days
    return FLOW_MIN_DAYS <= days <= FLOW_MAX_DAYS


def _cik_for_ticker(payload, ticker):
    wanted = ticker.upper()
    for item in payload.values():
        if str(item.get("ticker", "")).upper() == wanted:
            return str(item["cik_str"]).zfill(10)
    raise LookupError(f"no CIK for ticker {ticker}")


def _latest_10k_period(submissions, ticker):
    recent = submissions.get("filings", {}).get("recent", {})
    forms = recent.get("form", [])
    report_dates = recent.get("reportDate", [])
    for form, report_date in zip(forms, report_dates):
        if form == "10-K" and report_date:
            return report_date
    raise LookupError(f"no 10-K filing for ticker {ticker}")


def _get_json(url, user_agent):
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": str(user_agent),
            "Accept": "application/json",
        },
    )
    with urllib.request.urlopen(request, timeout=60) as response:
        return json.loads(response.read().decode("utf-8"))

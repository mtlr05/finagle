import pytest

from finagle.sec import baseline_from_facts, last_10k

PERIOD = "2023-12-31"
START = "2023-01-01"
FILED = "2024-02-20"


def _fact(val, end=PERIOD, start=START, form="10-K", fp="FY", filed=FILED, instant=False):
    fact = {"end": end, "val": val, "form": form, "fp": fp, "filed": filed, "accn": filed}
    if not instant:
        fact["start"] = start
    return fact


def _facts(**concepts):
    usgaap = {}
    for tag, facts in concepts.items():
        usgaap[tag] = {"units": {"USD": facts if isinstance(facts, list) else [facts]}}
    return {"facts": {"us-gaap": usgaap}}


def _clean():
    return _facts(
        RevenueFromContractWithCustomerExcludingAssessedTax=_fact(100_000_000),
        SalesRevenueNet=_fact(999_000_000),
        Revenues=[
            _fact(1_000_000, form="10-Q", fp="Q3", start="2023-07-01"),
            _fact(50_000_000, filed="2023-02-20"),
        ],
        IncomeTaxExpenseBenefit=_fact(10_000_000),
        InterestExpense=_fact(3_000_000),
        DepreciationDepletionAndAmortization=_fact(4_000_000),
        PaymentsToAcquirePropertyPlantAndEquipment=_fact(-2_000_000),
        ShareBasedCompensation=_fact(1_000_000),
        CashAndCashEquivalentsAtCarryingValue=_fact(8_000_000, instant=True),
        ShortTermInvestments=_fact(2_000_000, instant=True),
        LongTermDebtNoncurrent=_fact(40_000_000, instant=True),
        DebtCurrent=_fact(5_000_000, instant=True),
        OperatingIncomeLoss=_fact(20_000_000),
    )


def test_clean_mapping_scale_and_ebitda():
    baseline = baseline_from_facts(_clean(), PERIOD)

    assert baseline.financials["date"] == PERIOD
    assert baseline.financials["revenue"] == [100.0]
    assert baseline.financials["tax"] == [10.0]
    assert baseline.financials["interest"] == [3.0]
    assert baseline.financials["da"] == [4.0]
    assert baseline.financials["capex"] == [2.0]
    assert baseline.financials["sbc"] == [1.0]
    assert baseline.financials["cash"] == 10.0
    assert baseline.financials["debt"] == [45.0]
    assert baseline.financials["ebitda"] == [25.0]

    revenue = baseline.provenance.set_index("key").loc["revenue"]
    assert revenue["tags"] == "RevenueFromContractWithCustomerExcludingAssessedTax"
    assert revenue["raw dollars"] == 100_000_000
    assert revenue["value"] == 100.0

    ebitda = baseline.provenance.set_index("key").loc["ebitda"]
    assert ebitda["raw dollars"] == 25_000_000
    assert "OperatingIncomeLoss" in ebitda["tags"]
    assert "ShareBasedCompensation" in ebitda["tags"]
    assert "stock-based compensation" in ebitda["note"]


def test_debt_from_long_term_total_does_not_add_current_portion():
    payload = _facts(
        LongTermDebt=_fact(500_000_000, instant=True),
        DebtCurrent=_fact(100_000_000, instant=True),
        LongTermDebtCurrent=_fact(80_000_000, instant=True),
        ShortTermBorrowings=_fact(20_000_000, instant=True),
    )
    baseline = baseline_from_facts(payload, PERIOD, scale=1_000_000)

    assert baseline.financials["debt"] == [520.0]
    tags = baseline.provenance.set_index("key").loc["debt"]["tags"]
    assert "LongTermDebt" in tags
    assert "ShortTermBorrowings" in tags
    assert "DebtCurrent" not in tags
    assert "LongTermDebtCurrent" not in tags


def test_missing_interest_is_left_out_of_financials():
    payload = _clean()
    del payload["facts"]["us-gaap"]["InterestExpense"]
    baseline = baseline_from_facts(payload, PERIOD)

    assert "interest" not in baseline.financials
    interest = baseline.provenance.set_index("key").loc["interest"]
    assert interest["tags"] == ""
    assert "not found" in interest["note"]


def test_depreciation_pair_sums_and_builds_ebitda():
    payload = _facts(
        Depreciation=_fact(3_000_000),
        AmortizationOfIntangibleAssets=_fact(1_000_000),
        OperatingIncomeLoss=_fact(10_000_000),
    )
    baseline = baseline_from_facts(payload, PERIOD)

    assert baseline.financials["da"] == [4.0]
    assert baseline.financials["ebitda"] == [14.0]
    da = baseline.provenance.set_index("key").loc["da"]
    assert da["tags"] == "Depreciation, AmortizationOfIntangibleAssets"
    note = baseline.provenance.set_index("key").loc["ebitda"]["note"]
    assert "was not added" in note


def test_hrow_shaped_da_adds_separate_intangible_amortization_and_sbc():
    payload = _facts(
        DepreciationAndAmortization=_fact(1_915_000),
        AmortizationOfIntangibleAssets=_fact(16_991_000),
        OperatingIncomeLoss=_fact(30_515_000),
        ShareBasedCompensation=_fact(12_502_000),
    )
    baseline = baseline_from_facts(payload, PERIOD)
    table = baseline.provenance.set_index("key")

    assert table.loc["da"]["raw dollars"] == 18_906_000
    assert table.loc["da"]["tags"] == (
        "DepreciationAndAmortization, AmortizationOfIntangibleAssets")
    assert table.loc["ebitda"]["raw dollars"] == 61_923_000
    assert "ShareBasedCompensation" in table.loc["ebitda"]["tags"]


def test_combined_da_is_not_added_to_a_smaller_intangible_line():
    payload = _facts(
        DepreciationAndAmortization=_fact(10_000_000),
        AmortizationOfIntangibleAssets=_fact(3_000_000),
    )
    baseline = baseline_from_facts(payload, PERIOD)
    da = baseline.provenance.set_index("key").loc["da"]

    assert baseline.financials["da"] == [10.0]
    assert da["tags"] == "DepreciationAndAmortization"


def test_depreciation_alone_fills_da():
    payload = _facts(Depreciation=_fact(2_000_000))
    baseline = baseline_from_facts(payload, PERIOD)

    assert baseline.financials["da"] == [2.0]
    assert baseline.provenance.set_index("key").loc["da"]["tags"] == "Depreciation"


def test_ebitda_omitted_without_operating_income():
    payload = _facts(DepreciationDepletionAndAmortization=_fact(4_000_000))
    baseline = baseline_from_facts(payload, PERIOD)

    assert "ebitda" not in baseline.financials
    assert "da" in baseline.financials
    note = baseline.provenance.set_index("key").loc["ebitda"]["note"]
    assert "Operating income was not found" in note


def test_user_agent_and_scale_are_required():
    with pytest.raises(ValueError, match="user_agent"):
        last_10k("ATKR", "  ")
    with pytest.raises(ValueError, match="scale"):
        baseline_from_facts(_clean(), PERIOD, scale=0)

"""
Synthetic financial data generator (multi-entity, multi-currency, time-phased).

  - Transaction-level actuals with entity, currency, FX, cost centre, expense
    type, posting date and source journal.
  - Summarised budget + prior year by entity x cost-centre x account x period.
  - A full fiscal year is available; the workbook uses year-to-date actuals
    (through the reporting month) and the full-year budget.

Public functions:
  generate_transactions(period)      one month of transaction-level actuals
  generate_ytd_transactions(period)  Jan..period transactions (GL for the pack)
  generate_budget(period)            one month of summarised budget
  generate_budget_year()             all 12 months of budget
  monthly_detail(period)             entity x cc x account, actual+budget+prior
  generate_month(period)             account-level leaf GL (back-compatible)
"""

from __future__ import annotations
import numpy as np
import pandas as pd

from .coa import CHART_OF_ACCOUNTS, DEPARTMENT_COST_CENTRES


ENTITIES = [
    {"name": "Flux GmbH", "country": "DE", "region": "EMEA", "currency": "EUR", "fx": 1.0, "weight": 0.45},
    {"name": "Flux Inc.", "country": "US", "region": "Americas", "currency": "USD", "fx": 0.92, "weight": 0.35},
    {"name": "Flux Kft.", "country": "HU", "region": "EMEA", "currency": "HUF", "fx": 0.00256, "weight": 0.20},
]

# The demo company's fiscal year. One constant so the sample can be moved
# forward without hunting through defaults, and so nothing hardcodes 2025.
BASE_YEAR = 2025
DEFAULT_PERIOD = f"{BASE_YEAR}-06"

MONTHS = [f"{BASE_YEAR}-{m:02d}" for m in range(1, 13)]

BUDGET_BASELINE: dict[str, float] = {
    "4000": 355_000, "4010": 74_000, "4100": 505_000, "4110": 400_000,
    "4200": 210_000, "4210": 242_000, "4300": 95_000,
    "5000": 160_000, "5010": 95_000, "5100": 120_000, "5110": 85_000,
    "5200": 75_000, "5210": 48_000, "5300": 38_000, "5400": 22_000,
    "6000": 112_000, "6010": 58_000, "6001": 52_000, "6002": 38_000, "6100": 50_000,
    "6110": 76_000, "6120": 33_000, "6130": 21_000,
    "6200": 152_000, "6210": 58_000, "6220": 25_000, "6230": 29_000,
    "6300": 76_000, "6310": 38_000, "6320": 10_000, "6330": 18_000,
    "6340": 15_000, "6350": 8_000, "6360": 12_000, "6370": 8_000,
    "6380": 23_000, "6400": 27_000, "6410": 12_000,
    "7000": 16_000, "7100": 6_000, "7200": 4_000,
}

ACTUAL_SHOCKS: dict[str, float] = {
    "4000": 0.85, "4100": 1.11, "4200": 0.82, "5000": 1.16, "5400": 1.28,
    "6110": 1.42, "6120": 1.30, "6230": 1.38, "6300": 0.94, "6350": 1.55,
    "6340": 0.60, "6001": 1.08,
}

_REV_SEASON = [0.92, 0.90, 1.00, 0.98, 1.00, 1.02, 0.95, 0.93, 1.05, 1.08, 1.12, 1.20]
_MKT_SEASON = [0.90, 0.95, 1.25, 1.00, 1.05, 1.10, 0.85, 0.80, 1.30, 1.05, 1.10, 1.00]


def _season(category, group, m):
    if category == "Revenue":
        return _REV_SEASON[m]
    if group == "Sales & Marketing":
        return _MKT_SEASON[m]
    return 1.0


def _source(category, code):
    if category == "Revenue":
        return "AR"
    if code in ("6000", "6010", "6001", "6002", "6100", "6200", "6210", "6300"):
        return "Payroll"
    if code in ("6400", "6410"):
        return "Manual"
    if code in ("7000", "7100", "7200"):
        return "Bank"
    return "AP"


def _combos():
    """entity x account x department x cost centre, with the split weight.

    An account is assigned to one or more departments; within a department the
    amount is distributed across that department's cost centres.
    """
    for ent in ENTITIES:
        for acc in CHART_OF_ACCOUNTS:
            n_dept = len(acc.departments)
            for dept in acc.departments:
                ccs = DEPARTMENT_COST_CENTRES[dept]
                for cc_code, cc_name in ccs:
                    weight = (1.0 / n_dept) * (1.0 / len(ccs))
                    yield ent, acc, dept, cc_code, cc_name, ent["weight"], weight


def generate_budget(period=DEFAULT_PERIOD, seed=42):
    rng = np.random.default_rng(seed + 1 + MONTHS.index(period))
    m = MONTHS.index(period)
    rows = []
    for ent, acc, dept, cc_code, cc_name, ew, cw in _combos():
        base = float(BUDGET_BASELINE[acc.code]) * _season(acc.category, acc.group, m)
        budget = base * ew * cw
        py_factor = 0.93 if acc.category == "Revenue" else 0.96
        prior = budget * py_factor * (1 + rng.normal(0, 0.04))
        rows.append({
            "period": period, "period_no": int(period[:4]) * 100 + int(period[5:7]),
            "version": f"BUD-{BASE_YEAR}-V2", "entity": ent["name"],
            "region": ent["region"], "currency": ent["currency"],
            "department": dept, "cost_centre": cc_code,
            "owner": BUDGET_OWNERS.get(dept, "Finance"),
            "account_code": acc.code, "account_name": acc.name,
            "category": acc.category, "expense_type": acc.expense_type, "group": acc.group,
            "budget_eur": round(budget, 2), "prior_eur": round(prior, 2),
        })
    return pd.DataFrame(rows)


def generate_budget_year(seed=42):
    return pd.concat([generate_budget(p, seed) for p in MONTHS], ignore_index=True)


BUDGET_OWNERS = {
    "Commercial": "CFO", "Sales": "VP Sales", "Marketing": "CMO",
    "Engineering": "CTO", "Product": "CPO", "Operations": "COO", "G&A": "Finance",
}

DOC_TYPES = {           # source journal -> (document type, doc-type text)
    "AR": ("RV", "Billing document"),
    "AP": ("KR", "Vendor invoice"),
    "Payroll": ("PY", "Payroll posting"),
    "Bank": ("ZP", "Payment / bank"),
    "Manual": ("SA", "Manual journal"),
}

VENDORS = ["Northwind Supplies Ltd", "Aurora Media Group", "Helix Cloud Services",
           "Brightpath Consulting", "Kovacs Logistics Kft.", "Sentinel Insurance plc",
           "Delta Components GmbH", "Meridian Facilities", "Orion Staffing"]
CUSTOMERS = ["Vantage Retail Group", "Kestrel Systems Inc.", "Blue Harbour AS",
             "Danube Industrial Zrt.", "Summit Health Partners", "Ridgeline Motors"]
USERS = ["P.BENEDEK", "A.NOVAK", "J.MEYER", "S.OKONKWO", "BATCHUSER", "L.FISCHER"]
TAX_CODES = {"AR": ["A1", "A2", "AE"], "AP": ["V1", "V2", "VE"],
             "Payroll": ["**"], "Bank": ["**"], "Manual": ["**"]}


def _line_text(acc, source, partner, period):
    """Item text, the way a posting line is usually described."""
    month = period[-2:]
    if acc.category == "Revenue":
        return f"Invoice {partner} - {acc.name.split(' - ')[0]} {period}"
    if source == "Payroll":
        return f"Payroll run {period} - {acc.name.split(' - ')[-1]}"
    if source == "Manual":
        return f"Month-end accrual {month} - {acc.name}"
    if source == "Bank":
        return f"Bank posting {period} - {acc.name}"
    return f"{partner} - {acc.name}"


def generate_transactions(period=DEFAULT_PERIOD, seed=42):
    """Transaction-level actuals as a realistic ERP general-ledger extract.

    Each row is one posting line, carrying the document header fields (number,
    type, dates, reference, header text, created by) and line detail (line no,
    debit/credit indicator, amounts in document and reporting currency, cost
    centre, profit centre, assignment, item text, tax code, partner).
    """
    rng = np.random.default_rng(seed + MONTHS.index(period))
    m = MONTHS.index(period)
    year, month = period.split("-")
    rows, doc_seq = [], 0

    for ent, acc, dept, cc_code, cc_name, ew, cw in _combos():
        base = float(BUDGET_BASELINE[acc.code]) * _season(acc.category, acc.group, m)
        budget = base * ew * cw
        shock = ACTUAL_SHOCKS.get(acc.code, 1.0)
        actual_total = budget * shock * (1 + rng.normal(0, 0.05))
        if actual_total <= 0:
            continue

        source = _source(acc.category, acc.code)
        doc_type, doc_type_text = DOC_TYPES[source]
        k = int(rng.integers(1, 4))
        parts = rng.random(k); parts = parts / parts.sum()
        fx = ent["fx"]

        for p in parts:
            doc_seq += 1
            doc_no = f"{'19' if source in ('AP', 'AR') else '10'}{year[-2:]}{doc_seq:06d}"
            day = int(rng.integers(1, 29))
            post_day = day
            doc_day = max(1, day - int(rng.integers(0, 8)))       # invoice predates posting
            entry_day = min(28, day + int(rng.integers(0, 3)))    # entered on/after posting
            partner = (str(rng.choice(CUSTOMERS)) if acc.category == "Revenue"
                       else str(rng.choice(VENDORS)) if source == "AP" else "")
            amt_eur = actual_total * p
            amt_lcy = amt_eur / fx
            # Revenue is credited, costs are debited.
            dc = "H" if acc.category == "Revenue" else "S"
            reversed_flag = "X" if rng.random() < 0.01 else ""

            rows.append({
                # --- document header ---
                "doc_no": doc_no,
                "doc_type": doc_type,
                "doc_type_text": doc_type_text,
                "fiscal_year": int(year),
                "period": period,
                "period_no": int(year) * 100 + int(month),
                "doc_date": f"{period}-{doc_day:02d}",
                "posting_date": f"{period}-{post_day:02d}",
                "entry_date": f"{period}-{entry_day:02d}",
                "entry_time": f"{int(rng.integers(6,20)):02d}:{int(rng.integers(0,60)):02d}",
                "reference": (f"INV-{int(rng.integers(100000,999999))}"
                              if source in ("AR", "AP") else f"{source}-{period}"),
                "header_text": doc_type_text + f" {period}",
                "created_by": ("BATCHUSER" if source in ("Payroll", "Bank")
                               else str(rng.choice(USERS))),
                "reversed": reversed_flag,
                # --- organisation ---
                "entity": ent["name"],
                "company_code": {"Flux GmbH": "1000", "Flux Inc.": "2000",
                                 "Flux Kft.": "3000"}[ent["name"]],
                "region": ent["region"],
                "department": dept,
                "cost_centre": cc_code,
                "profit_centre": f"PC{cc_code[2:6]}",
                # --- line detail ---
                "line_no": int(rng.integers(1, 20)),
                "account_code": acc.code,
                "account_name": acc.name,
                "category": acc.category,
                "expense_type": acc.expense_type,
                "group": acc.group,
                "dc_indicator": dc,
                "currency": ent["currency"],
                "amount_lcy": round(amt_lcy, 0 if ent["currency"] == "HUF" else 2),
                "fx_rate": fx,
                "amount_eur": round(amt_eur, 2),
                "tax_code": str(rng.choice(TAX_CODES[source])),
                "partner": partner,
                "assignment": f"{cc_code}-{period.replace('-', '')}",
                "line_text": _line_text(acc, source, partner or "Internal", period),
                "source": source,
            })

    df = pd.DataFrame(rows)
    return df.sort_values(["period_no", "posting_date", "doc_no", "entity",
                           "account_code"]).reset_index(drop=True)


def generate_ytd_transactions(period=DEFAULT_PERIOD, seed=42):
    idx = MONTHS.index(period)
    parts = [generate_transactions(MONTHS[i], seed) for i in range(idx + 1)]
    df = pd.concat(parts, ignore_index=True)
    return df.sort_values(["period_no", "posting_date", "doc_no", "entity",
                           "account_code"]).reset_index(drop=True)


def monthly_detail(period=DEFAULT_PERIOD, seed=42):
    txns = generate_transactions(period, seed)
    bud = generate_budget(period, seed)
    act = (txns.groupby(["entity", "region", "department", "cost_centre", "account_code",
                         "account_name", "category", "expense_type", "group"], as_index=False)
           ["amount_eur"].sum().rename(columns={"amount_eur": "actual"}))
    merged = bud.merge(act[["entity", "department", "cost_centre", "account_code", "actual"]],
                       on=["entity", "department", "cost_centre", "account_code"], how="left")
    merged["actual"] = merged["actual"].fillna(0.0)
    return merged.rename(columns={"budget_eur": "budget", "prior_eur": "prior_year"})


def generate_month(period=DEFAULT_PERIOD, seed=42):
    detail = monthly_detail(period, seed)
    agg = (detail.groupby(["account_code", "account_name", "category"], as_index=False)
           [["actual", "budget", "prior_year"]].sum())
    order = {a.code: i for i, a in enumerate(CHART_OF_ACCOUNTS)}
    agg = agg.sort_values("account_code", key=lambda s: s.map(order)).reset_index(drop=True)
    agg["period"] = period
    return agg[["period", "account_code", "account_name", "category",
                "actual", "budget", "prior_year"]]


if __name__ == "__main__":
    ytd = generate_ytd_transactions()
    by = generate_budget_year()
    print(f"YTD transactions (Jan-Jun): {len(ytd)} rows, {len(ytd.columns)} cols")
    print(f"Full-year budget: {len(by)} rows")
    print(f"Month actual (EUR): "
          f"{ytd[ytd.period == DEFAULT_PERIOD]['amount_eur'].sum():,.0f}")
    print(f"YTD actual  (EUR): {ytd['amount_eur'].sum():,.0f}")
    print(f"FY budget   (EUR): {by['budget_eur'].sum():,.0f}")
    print("Expense types present:", sorted(ytd['expense_type'].unique()) != [])

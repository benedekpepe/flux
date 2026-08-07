"""
Chart of accounts, cost-centre structure, expense types and P&L layout.

Dimensions on each account:
  - category  : Revenue | COGS | OpEx | Other  (drives the P&L roll-up and sign)
  - group     : P&L sub-group (functional grouping)
  - expense_type : natural expense classification (Salaries, Marketing, ...),
                   used for the management expense report
  - departments : department(s) the account is split across; each department
                  distributes across its own cost centres
  - favourable : direction in which a positive budget variance is good

Revenue and profit lines are "higher is better"; cost lines "lower is better".
"""

from __future__ import annotations
from dataclasses import dataclass


FAV_HIGHER = "higher"
FAV_LOWER = "lower"

# Organisational hierarchy: department -> cost centres.
# Departments are the reporting roll-up; cost centres are the granular units
# spend is actually booked to.
DEPARTMENT_COST_CENTRES: dict[str, list[tuple[str, str]]] = {
    # Codes follow a typical ERP scheme: a two-letter department prefix plus a
    # five-digit block (department block + sequence). Every company codes cost
    # centres differently, so this is illustrative only - Flux accepts whatever
    # scheme a client uses.
    "Commercial":  [("CM10100", "Commercial")],
    "Sales":       [("SL20100", "Sales - EMEA"),
                    ("SL20200", "Sales - Americas"),
                    ("SL20300", "Sales Operations")],
    "Marketing":   [("MK30100", "Brand & Content"),
                    ("MK30200", "Demand Generation"),
                    ("MK30300", "Events")],
    "Engineering": [("EN40100", "Platform Engineering"),
                    ("EN40200", "Application Engineering"),
                    ("EN40300", "QA & Release")],
    "Product":     [("PD50100", "Product Management"),
                    ("PD50200", "Design")],
    "Operations":  [("OP60100", "Manufacturing"),
                    ("OP60200", "Supply Chain"),
                    ("OP60300", "Customer Support")],
    "G&A":         [("GA70100", "Finance"),
                    ("GA70200", "HR & Legal"),
                    ("GA70300", "IT"),
                    ("GA70400", "Facilities")],
}

DEPARTMENTS = list(DEPARTMENT_COST_CENTRES)
# Departments carrying a spend budget (Commercial books revenue, not spend).
SPEND_DEPARTMENTS = [d for d in DEPARTMENTS if d != "Commercial"]

COST_CENTRES = [(dept, code, name)
                for dept, ccs in DEPARTMENT_COST_CENTRES.items()
                for code, name in ccs]
COST_CENTRE_NAMES = [name for _, _, name in COST_CENTRES]
DEPARTMENT_OF_CC = {name: dept for dept, _, name in COST_CENTRES}


@dataclass(frozen=True)
class Account:
    code: str
    name: str
    category: str
    group: str
    expense_type: str
    favourable: str
    departments: tuple


CATEGORY_FAVOURABLE = {"Revenue": FAV_HIGHER, "COGS": FAV_LOWER, "OpEx": FAV_LOWER, "Other": FAV_LOWER}
CATEGORY_SIGN = {"Revenue": +1, "COGS": -1, "OpEx": -1, "Other": -1}


def _a(code, name, category, group, expense_type, fav, depts):
    return Account(code, name, category, group, expense_type, fav, depts)


CHART_OF_ACCOUNTS: list[Account] = [
    # Revenue
    _a("4000", "Product revenue - hardware", "Revenue", "Revenue", "Revenue", FAV_HIGHER, ("Commercial",)),
    _a("4010", "Product revenue - accessories", "Revenue", "Revenue", "Revenue", FAV_HIGHER, ("Commercial",)),
    _a("4100", "Subscription revenue - monthly", "Revenue", "Revenue", "Revenue", FAV_HIGHER, ("Commercial",)),
    _a("4110", "Subscription revenue - annual", "Revenue", "Revenue", "Revenue", FAV_HIGHER, ("Commercial",)),
    _a("4200", "Service revenue - implementation", "Revenue", "Revenue", "Revenue", FAV_HIGHER, ("Commercial",)),
    _a("4210", "Support contract revenue", "Revenue", "Revenue", "Revenue", FAV_HIGHER, ("Commercial",)),
    _a("4300", "Training revenue", "Revenue", "Revenue", "Revenue", FAV_HIGHER, ("Commercial",)),
    # COGS
    _a("5000", "Direct materials", "COGS", "Cost of sales", "Materials & components", FAV_LOWER, ("Operations",)),
    _a("5010", "Component costs", "COGS", "Cost of sales", "Materials & components", FAV_LOWER, ("Operations",)),
    _a("5100", "Direct labour - manufacturing", "COGS", "Cost of sales", "Direct labour", FAV_LOWER, ("Operations",)),
    _a("5110", "Direct labour - services", "COGS", "Cost of sales", "Direct labour", FAV_LOWER, ("Operations",)),
    _a("5200", "Cloud hosting & infrastructure", "COGS", "Cost of sales", "Hosting & infrastructure", FAV_LOWER, ("Operations", "Engineering")),
    _a("5210", "Third-party licences", "COGS", "Cost of sales", "Third-party licences", FAV_LOWER, ("Operations",)),
    _a("5300", "Shipping & fulfilment", "COGS", "Cost of sales", "Logistics & fees", FAV_LOWER, ("Operations",)),
    _a("5400", "Payment processing fees", "COGS", "Cost of sales", "Logistics & fees", FAV_LOWER, ("Operations",)),
    # OpEx - personnel
    _a("6000", "Salaries - Sales", "OpEx", "Sales & Marketing", "Salaries & wages", FAV_LOWER, ("Sales",)),
    _a("6010", "Salaries - Marketing", "OpEx", "Sales & Marketing", "Salaries & wages", FAV_LOWER, ("Marketing",)),
    _a("6001", "Payroll taxes & social security", "OpEx", "Personnel", "Payroll benefits", FAV_LOWER, ("Sales", "Engineering", "G&A")),
    _a("6002", "Pension & insurance benefits", "OpEx", "Personnel", "Payroll benefits", FAV_LOWER, ("Sales", "Engineering", "G&A")),
    _a("6100", "Sales commissions", "OpEx", "Sales & Marketing", "Employee incentives", FAV_LOWER, ("Sales",)),
    # OpEx - Sales & Marketing
    _a("6110", "Advertising & digital", "OpEx", "Sales & Marketing", "Marketing & advertising", FAV_LOWER, ("Marketing",)),
    _a("6120", "Events & trade shows", "OpEx", "Sales & Marketing", "Marketing & advertising", FAV_LOWER, ("Marketing",)),
    _a("6130", "Content & PR", "OpEx", "Sales & Marketing", "Marketing & advertising", FAV_LOWER, ("Marketing",)),
    # OpEx - R&D
    _a("6200", "Salaries - Engineering", "OpEx", "Research & Development", "Salaries & wages", FAV_LOWER, ("Engineering",)),
    _a("6210", "Salaries - Product", "OpEx", "Research & Development", "Salaries & wages", FAV_LOWER, ("Product",)),
    _a("6220", "Software & tools (R&D)", "OpEx", "Research & Development", "IT & software", FAV_LOWER, ("Engineering", "Product")),
    _a("6230", "Contractors - R&D", "OpEx", "Research & Development", "External staff / contractors", FAV_LOWER, ("Engineering",)),
    # OpEx - G&A
    _a("6300", "Salaries - G&A", "OpEx", "General & Administrative", "Salaries & wages", FAV_LOWER, ("G&A",)),
    _a("6310", "Rent & facilities", "OpEx", "General & Administrative", "Facilities & office", FAV_LOWER, ("G&A",)),
    _a("6320", "Utilities", "OpEx", "General & Administrative", "Facilities & office", FAV_LOWER, ("G&A",)),
    _a("6330", "IT & software (G&A)", "OpEx", "General & Administrative", "IT & software", FAV_LOWER, ("G&A",)),
    _a("6340", "Legal fees", "OpEx", "General & Administrative", "External professional fees", FAV_LOWER, ("G&A",)),
    _a("6350", "Audit & accounting", "OpEx", "General & Administrative", "External professional fees", FAV_LOWER, ("G&A",)),
    _a("6360", "Insurance", "OpEx", "General & Administrative", "Insurance", FAV_LOWER, ("G&A",)),
    _a("6370", "Office & admin", "OpEx", "General & Administrative", "Facilities & office", FAV_LOWER, ("G&A",)),
    _a("6380", "Travel & entertainment", "OpEx", "General & Administrative", "Travel & entertainment", FAV_LOWER, ("Sales", "Engineering", "G&A")),
    _a("6400", "Depreciation", "OpEx", "General & Administrative", "Depreciation & amortisation", FAV_LOWER, ("G&A",)),
    _a("6410", "Amortisation", "OpEx", "General & Administrative", "Depreciation & amortisation", FAV_LOWER, ("G&A",)),
    # Other
    _a("7000", "Interest expense", "Other", "Other", "Financing & bank", FAV_LOWER, ("G&A",)),
    _a("7100", "FX losses", "Other", "Other", "Financing & bank", FAV_LOWER, ("G&A",)),
    _a("7200", "Bank charges", "Other", "Other", "Financing & bank", FAV_LOWER, ("G&A",)),
]

ACCOUNTS_BY_CODE = {a.code: a for a in CHART_OF_ACCOUNTS}

# Expense types grouped the way a management expense report is read: cost of
# sales first (above gross profit), then operating expenses with personnel
# leading, then non-cash and financing items.
EXPENSE_GROUPS: list[tuple[str, list[str]]] = [
    # Personnel leads: it is the largest and most actively managed block, and it
    # is what a cost owner looks at first. Within the group the order follows the
    # pay structure - base pay, statutory add-ons, variable pay, bought-in labour.
    ("Personnel costs", [
        "Salaries & wages",
        "Payroll benefits",
        "Employee incentives",
        "External staff / contractors",
    ]),
    ("Other operating costs", [
        "Marketing & advertising",
        "External professional fees",
        "IT & software",
        "Facilities & office",
        "Travel & entertainment",
        "Insurance",
    ]),
    ("Cost of sales", [
        "Materials & components",
        "Direct labour",
        "Hosting & infrastructure",
        "Third-party licences",
        "Logistics & fees",
    ]),
    ("Depreciation & amortisation", [
        "Depreciation & amortisation",
    ]),
    ("Financing", [
        "Financing & bank",
    ]),
]

EXPENSE_TYPES = [t for _g, types in EXPENSE_GROUPS for t in types]

@dataclass(frozen=True)
class ReportLine:
    label: str
    kind: str
    favourable: str
    category: str | None = None
    components: tuple | None = None


PNL_STRUCTURE: list[ReportLine] = [
    ReportLine("Revenue", "category", FAV_HIGHER, category="Revenue"),
    ReportLine("Cost of goods sold", "category", FAV_LOWER, category="COGS"),
    ReportLine("Gross profit", "computed", FAV_HIGHER,
               components=(("+", "Revenue"), ("-", "Cost of goods sold"))),
    ReportLine("Operating expenses", "category", FAV_LOWER, category="OpEx"),
    ReportLine("Operating income (EBIT)", "computed", FAV_HIGHER,
               components=(("+", "Gross profit"), ("-", "Operating expenses"))),
    ReportLine("Other expenses", "category", FAV_LOWER, category="Other"),
    ReportLine("Net income", "computed", FAV_HIGHER,
               components=(("+", "Operating income (EBIT)"), ("-", "Other expenses"))),
]

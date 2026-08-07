"""
Variance commentary generation.

Turns the engine's structured output into a management-style narrative:
a headline on the bottom line, a revenue/gross-profit paragraph, and a driver
paragraph that names the material movers split into favourable and unfavourable.

Two modes:

  - Template (default): pure Python, no API, no cost, works in any demo. It
    describes magnitude and direction faithfully. It does NOT invent root
    causes ("why" a line moved) because that information is not in the data.
  - LLM (optional): if an ANTHROPIC_API_KEY is present and use_llm is on, the
    template narrative is handed to the API as a base to enrich. Any failure or
    missing key falls back to the template, so a public demo never depends on a
    paid key.
"""

from __future__ import annotations
import os

import pandas as pd

from .engine import leaf_variances, MaterialityRule


# Enabled only when a key is actually present in the environment.
HAS_LLM = bool(os.environ.get("ANTHROPIC_API_KEY"))


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------

def _mag(x: float) -> str:
    """Absolute magnitude in EUR with k/m shorthand for readability."""
    ax = abs(float(x))
    if ax >= 1_000_000:
        return f"\u20ac{ax/1_000_000:.1f}m"
    if ax >= 1_000:
        return f"\u20ac{ax/1_000:.1f}k"
    return f"\u20ac{ax:.0f}"


def _pct(x: float | None) -> str:
    if x is None or pd.isna(x):
        return "n/a"
    return f"{abs(float(x))*100:.1f}%"


def _not_meaningful(x) -> bool:
    """True when a percentage is absent or n/m (pandas turns None into NaN)."""
    return x is None or pd.isna(x)


def _join(items: list[str]) -> str:
    """Join a list into 'a', 'a and b', or 'a, b and c'."""
    items = [i for i in items if i]
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    if len(items) == 2:
        return f"{items[0]} and {items[1]}"
    return ", ".join(items[:-1]) + f" and {items[-1]}"


# A separator inside an item would make the list ambiguous.
_AMBIGUOUS = (" - ", " \u2013 ", ", ")


def _join_names(items: list[str]) -> str:
    """Join account names, keeping the boundary between them visible.

    Chart-of-accounts names frequently carry their own separator - "Product
    revenue - hardware" - so joining them with commas and "and" leaves a reader
    unable to tell where one account ends and the next begins. When any name
    contains a separator the list switches to semicolons, which outrank the
    dash and settle the boundary without adding length.
    """
    items = [i for i in items if i]
    if len(items) < 2 or not any(sep in i for i in items for sep in _AMBIGUOUS):
        return _join(items)
    return "; ".join(items)


def _line(report: pd.DataFrame, label: str) -> dict:
    row = report[report["line"] == label].iloc[0]
    return row.to_dict()


# ---------------------------------------------------------------------------
# Template commentary
# ---------------------------------------------------------------------------

def _budgeted(report: pd.DataFrame) -> bool:
    """Whether this report carries a plan to comment against."""
    if "budgeted" in report.attrs:
        return bool(report.attrs["budgeted"])
    return bool(report["var_bud"].notna().any())


def _actuals_only_commentary(report: pd.DataFrame) -> str:
    """Narrative for a ledger supplied without a plan.

    There is nothing to call favourable or unfavourable, so the commentary
    describes the result and margin rather than inventing a comparison.
    """
    ni = _line(report, "Net income")
    ebit = _line(report, "Operating income (EBIT)")
    rev = _line(report, "Revenue")
    gp = _line(report, "Gross profit")
    opex = _line(report, "Operating expenses")

    ni_term = (f"a net loss of {_mag(ni['actual'])}" if ni["actual"] < 0
               else f"net income of {_mag(ni['actual'])}")
    margin = (f" a gross margin of {gp['actual'] / rev['actual'] * 100:.1f}%"
              if rev["actual"] else " no revenue in the period")

    p1 = (f"Revenue of {_mag(rev['actual'])} produced gross profit of "
          f"{_mag(gp['actual'])},{margin}. After operating expenses of "
          f"{_mag(opex['actual'])} the result is {_mag(ebit['actual'])} at EBIT "
          f"and {ni_term}.")
    p2 = ("No budget was supplied with this file, so no variance, favourable / "
          "unfavourable classification or materiality flagging is reported. "
          "Upload a budget alongside the ledger for the full variance pack.")
    return "\n\n".join([p1, p2])


def _template_commentary(
    report: pd.DataFrame,
    gl: pd.DataFrame,
    materiality: MaterialityRule | None = None,
) -> str:
    materiality = materiality or MaterialityRule()
    if not _budgeted(report):
        return _actuals_only_commentary(report)

    ni = _line(report, "Net income")
    ebit = _line(report, "Operating income (EBIT)")
    rev = _line(report, "Revenue")
    gp = _line(report, "Gross profit")
    cogs = _line(report, "Cost of goods sold")

    def dir_word(var: float) -> str:
        return "above" if var >= 0 else "below"

    def fu_word(fu: str) -> str:
        return "favourable" if fu == "F" else "unfavourable"

    # Paragraph 1 - headline on the bottom line.
    ni_term = (f"A net loss of {_mag(ni['actual'])}" if ni["actual"] < 0
               else f"Net income of {_mag(ni['actual'])}")
    ebit_term = (f"an operating loss of {_mag(ebit['actual'])}" if ebit["actual"] < 0
                 else f"operating income (EBIT) of {_mag(ebit['actual'])}")
    p1 = (
        f"{ni_term} came in {_mag(ni['var_bud'])} "
        f"({_pct(ni['var_bud_pct'])}) {dir_word(ni['var_bud'])} budget, "
        f"an {fu_word(ni['fav_unfav'])} result. This reflects {ebit_term}, "
        f"{_mag(ebit['var_bud'])} ({_pct(ebit['var_bud_pct'])}) "
        f"{dir_word(ebit['var_bud'])} plan."
    )

    # Paragraph 2 - revenue and gross profit.
    if _not_meaningful(rev["var_bud_pct"]) or abs(rev["var_bud_pct"]) < 0.01:
        rev_clause = f"Revenue of {_mag(rev['actual'])} was broadly in line with budget"
    else:
        rev_clause = (
            f"Revenue of {_mag(rev['actual'])} was {_mag(rev['var_bud'])} "
            f"({_pct(rev['var_bud_pct'])}) {dir_word(rev['var_bud'])} budget"
        )
    p2 = (
        f"{rev_clause}. Gross profit of {_mag(gp['actual'])} was "
        f"{_mag(gp['var_bud'])} ({_pct(gp['var_bud_pct'])}) "
        f"{dir_word(gp['var_bud'])} budget, with cost of goods sold "
        f"{dir_word(cogs['var_bud'])} plan by {_mag(cogs['var_bud'])} "
        f"({_pct(cogs['var_bud_pct'])})."
    )

    # Paragraph 3 - material account-level drivers.
    leaves = leaf_variances(gl, materiality)
    material = leaves[leaves["material"]]
    unfav = material[material["fav_unfav"] == "U"]
    fav = material[material["fav_unfav"] == "F"]

    def phrase(r) -> str:
        return (
            f"{r['account_name'].lower()} ({_mag(r['var_bud'])}, "
            f"{_pct(r['var_bud_pct'])} {dir_word(r['var_bud'])} budget)"
        )

    parts = []
    if len(unfav):
        parts.append(
            "The main unfavourable drivers were "
            + _join_names([phrase(r) for _, r in unfav.iterrows()])
            + "."
        )
    if len(fav):
        lead = "This was partly offset by" if len(unfav) else "The main favourable drivers were"
        parts.append(
            f"{lead} "
            + _join_names([phrase(r) for _, r in fav.iterrows()])
            + "."
        )
    p3 = " ".join(parts) if parts else (
        "No account-level variances cleared the materiality thresholds."
    )

    return "\n\n".join([p1, p2, p3])


# ---------------------------------------------------------------------------
# Optional LLM enrichment (lazy, graceful)
# ---------------------------------------------------------------------------

def _llm_commentary(base: str, report: pd.DataFrame) -> str:
    """Enrich the template narrative via the Anthropic API.

    Lazily imports the SDK so the template path never depends on it. Any
    failure is raised to the caller, which falls back to the template.
    """
    import anthropic  # optional dependency

    table = report[["line", "actual", "budget", "var_bud", "var_bud_pct", "fav_unfav", "material"]]
    client = anthropic.Anthropic()  # reads ANTHROPIC_API_KEY from env
    prompt = (
        "You are an FP&A analyst writing management commentary for a monthly "
        "P&L variance report. Below is a draft narrative and the underlying "
        "figures. Tighten the draft into concise, professional commentary. "
        "Do not invent root causes that are not implied by the numbers; "
        "describe magnitude, direction and materiality only.\n\n"
        f"Draft:\n{base}\n\nFigures:\n{table.to_string(index=False)}"
    )
    msg = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=600,
        messages=[{"role": "user", "content": prompt}],
    )
    return "".join(b.text for b in msg.content if getattr(b, "type", None) == "text").strip()


def line_comments(
    report: pd.DataFrame,
    gl: pd.DataFrame,
    materiality: MaterialityRule | None = None,
) -> dict[str, str]:
    """Short per-line commentary for the P&L, keyed by line label.

    Each comment states the budget variance in EUR and %, direction and
    favourability, flags materiality, and for category lines names the material
    account-level drivers sitting inside that line. Generated at build time, so
    it refreshes when the export is re-run rather than live in-cell.
    """
    materiality = materiality or MaterialityRule()
    if not _budgeted(report):
        # Nothing to compare against: state the amount and stop there rather
        # than describing a variance measured from zero.
        return {r["line"]: f"{_mag(r['actual'])} actual; no budget supplied."
                for _, r in report.iterrows()}

    leaves = leaf_variances(gl, materiality)
    material = leaves[leaves["material"]]

    cat_of_line = {
        "Revenue": "Revenue",
        "Cost of goods sold": "COGS",
        "Operating expenses": "OpEx",
        "Other expenses": "Other",
    }

    comments: dict[str, str] = {}
    for _, r in report.iterrows():
        label = r["line"]
        var = r["var_bud"]
        pct = r["var_bud_pct"]
        fu = "favourable" if r["fav_unfav"] == "F" else "unfavourable"
        direction = "above" if var >= 0 else "below"

        if not _not_meaningful(pct) and abs(pct) < 0.005:
            base = "In line with budget"
        elif _not_meaningful(pct):
            base = f"{_mag(var)} {direction} budget, {fu}; percentage not meaningful"
        else:
            base = f"{_mag(var)} ({_pct(pct)}) {direction} budget, {fu}"

        if r["material"]:
            base += "; material"

        if label in cat_of_line:
            drv = material[material["category"] == cat_of_line[label]]
            if len(drv):
                names = _join_names([d["account_name"].lower() for _, d in drv.iterrows()])
                if len(set(drv["fav_unfav"])) > 1:
                    base += f"; offsetting movements in {names}"
                else:
                    base += f"; driven by {names}"

        comments[label] = base + "."
    return comments


def generate_commentary(
    report: pd.DataFrame,
    gl: pd.DataFrame,
    materiality: MaterialityRule | None = None,
    use_llm: bool | None = None,
) -> str:
    """Generate variance commentary. Falls back to the template on any LLM issue."""
    base = _template_commentary(report, gl, materiality)
    use_llm = HAS_LLM if use_llm is None else use_llm
    if use_llm:
        try:
            return _llm_commentary(base, report)
        except Exception:
            return base
    return base


if __name__ == "__main__":
    from .synthetic_data import generate_month
    from .engine import build_report

    gl = generate_month()
    report = build_report(gl)
    print(generate_commentary(report, gl))

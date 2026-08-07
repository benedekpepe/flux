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



# ---------------------------------------------------------------------------
# Two-horizon commentary
# ---------------------------------------------------------------------------
# A comment that describes only the month cannot explain the flag beside it.
# MONTH, YTD and BOTH are three different stories - a one-off, a pattern the
# month happens not to show, and a pattern the month is still feeding - and the
# commentary is the one column with room to say which.

def _clause(var: float, pct, fu_word: str | None = None) -> str:
    """One timeframe, as a phrase: amount, percentage, direction, verdict."""
    direction = "above" if var >= 0 else "below"
    if not _not_meaningful(pct) and abs(pct) < 0.005:
        # "In line with budget, unfavourable" is a contradiction: a line that
        # did not move has no verdict worth printing.
        return "in line with budget"
    if _not_meaningful(pct):
        core = f"{_mag(var)} {direction} budget, percentage not meaningful"
    else:
        core = f"{_mag(var)} ({_pct(pct)}) {direction} budget"
    return f"{core}, {fu_word}" if fu_word else core


_VERDICTS = {
    (True, True): "Material on both timeframes.",
    (False, True): "Material cumulatively; the month alone stays within the floors.",
    (True, False): "Material this month; the year to date is still within the floors.",
    (False, False): "Neither timeframe clears the materiality floors.",
}


def _two_horizon(month: tuple, ytd: tuple, fu_word: str) -> str:
    """The shared sentence shape: year to date, then month, then the verdict.

    Each tuple is (variance, percentage, material). The year to date leads
    because it is the trend; the month follows because it is the latest point.
    Reversing them makes every comment read as news even when the line has been
    drifting since January.
    """
    m_var, m_pct, m_mat = month
    y_var, y_pct, y_mat = ytd
    return (f"YTD {_clause(y_var, y_pct, fu_word)}. "
            f"Month {_clause(m_var, m_pct)}. "
            f"{_VERDICTS[(bool(m_mat), bool(y_mat))]}")


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
    ytd_report: pd.DataFrame | None = None,
    ytd_gl: pd.DataFrame | None = None,
) -> dict[str, str]:
    """Short per-line commentary for the P&L, keyed by line label.

    With a year-to-date report supplied, each comment covers both timeframes and
    says which one clears the materiality floors - the same verdict the Flag
    column carries, spelled out. Without one it falls back to describing the
    month alone, which is what the deck needs.

    For category lines the comment also names the material account-level drivers
    sitting inside that line, taken from whichever timeframe is doing the
    talking. Generated at build time, so it refreshes when the export is re-run
    rather than live in-cell.
    """
    materiality = materiality or MaterialityRule()
    if not _budgeted(report):
        # Nothing to compare against: state the amount and stop there rather
        # than describing a variance measured from zero.
        return {r["line"]: f"{_mag(r['actual'])} actual; no budget supplied."
                for _, r in report.iterrows()}

    cat_of_line = {
        "Revenue": "Revenue",
        "Cost of goods sold": "COGS",
        "Operating expenses": "OpEx",
        "Other expenses": "Other",
    }

    month_material = leaf_variances(gl, materiality)
    month_material = month_material[month_material["material"]]
    ytd_rows = None
    if ytd_report is not None:
        ytd_rows = {r["line"]: r for _, r in ytd_report.iterrows()}
        ytd_material = leaf_variances(ytd_gl if ytd_gl is not None else gl, materiality)
        ytd_material = ytd_material[ytd_material["material"]]
    else:
        ytd_material = month_material

    def drivers(pool, category) -> str:
        drv = pool[pool["category"] == category]
        if not len(drv):
            return ""
        names = _join_names([d["account_name"].lower() for _, d in drv.iterrows()])
        lead = ("offsetting movements in" if len(set(drv["fav_unfav"])) > 1
                else "driven by")
        return f" {lead.capitalize()[0]}{lead[1:]} {names}."

    comments: dict[str, str] = {}
    for _, r in report.iterrows():
        label = r["line"]
        fu = "favourable" if r["fav_unfav"] == "F" else "unfavourable"

        if ytd_rows is None or label not in ytd_rows:
            # No cumulative figures: describe the month and stop there rather
            # than implying a trend the caller has not supplied.
            base = _clause(r["var_bud"], r["var_bud_pct"], fu)
            base = base[0].upper() + base[1:]
            if r["material"]:
                base += "; material"
            base += "." + drivers(month_material, cat_of_line.get(label, ""))
            comments[label] = base.strip()
            continue

        y = ytd_rows[label]
        base = _two_horizon((r["var_bud"], r["var_bud_pct"], r["material"]),
                            (y["var_bud"], y["var_bud_pct"], y["material"]), fu)
        # The drivers come from whichever timeframe is doing the talking: a
        # month-only movement is not explained by the accounts that have been
        # drifting all year.
        pool = month_material if (r["material"] and not y["material"]) else ytd_material
        if r["material"] or y["material"]:
            base += drivers(pool, cat_of_line.get(label, ""))
        comments[label] = base.strip()
    return comments


def rollup_comments(
    detail: pd.DataFrame,
    parent: str,
    child: str,
    materiality: MaterialityRule | None = None,
    higher_is_better: bool = False,
    child_names: dict | None = None,
    ytd_detail: pd.DataFrame | None = None,
) -> dict[str, str]:
    """Commentary for the roll-up rows of any parent/child dimension.

    Written only for rows that have something underneath them - a department
    over its cost centres, an expense group over its types, an entity over its
    departments. On a leaf row the comment could only restate the variance and
    percentage sitting in the columns beside it, which is noise; on a roll-up it
    can name *which* of the lines below moved the total, and that is not on the
    row.

    Every roll-up gets one, not just the material ones. A blank cell is
    ambiguous - nothing moved, or nothing was generated? - and "neither
    timeframe clears the floors" is an answer. The Flag column is what points a
    reader at the rows worth stopping on; this column explains them.

    `child_names` maps a child key to how it should read in prose, for the
    dimensions whose key is a code rather than a name.
    """
    materiality = materiality or MaterialityRule()
    child_names = child_names or {}
    needed = {parent, child, "actual", "budget"}
    if not needed <= set(detail.columns):
        return {}

    def verdict(var, budget):
        pct = None if abs(budget) < 100 else var / abs(budget)
        if pct is not None and abs(pct) > 10:
            pct = None
        fu = "F" if (var >= 0 if higher_is_better else var <= 0) else "U"
        material = (abs(var) >= materiality.abs_threshold
                    and (pct is None or abs(pct) >= materiality.pct_threshold))
        return pct, fu, material

    def totals(frame):
        return (frame.groupby(parent, as_index=False)[["actual", "budget"]].sum(),
                frame.groupby([parent, child], as_index=False)[["actual", "budget"]].sum())

    parents, children = totals(detail)
    if ytd_detail is not None and needed <= set(ytd_detail.columns):
        y_parents, y_children = totals(ytd_detail)
    else:
        y_parents = y_children = None

    def movers(frame, key):
        named = []
        for _, c in frame[frame[parent] == key].iterrows():
            cvar = c["actual"] - c["budget"]
            _cpct, cfu, cmaterial = verdict(cvar, c["budget"])
            if cmaterial:
                named.append((str(child_names.get(c[child], c[child])).lower(), cfu))
        if not named:
            return ""
        names = _join_names([m for m, _ in named])
        mixed = len({fu for _, fu in named}) > 1
        lead = "Offsetting movements in" if mixed else "Driven by"
        return f" {lead} {names}."

    comments: dict[str, str] = {}
    for _, p in parents.iterrows():
        key = p[parent]
        var = p["actual"] - p["budget"]
        pct, fu, material = verdict(var, p["budget"])
        word = "favourable" if fu == "F" else "unfavourable"

        if y_parents is None:
            base = _clause(var, pct, word)
            base = base[0].upper() + base[1:]
            if material:
                base += "; material"
            base += "." + movers(children, key)
            comments[key] = base.strip()
            continue

        row = y_parents[y_parents[parent] == key]
        if row.empty:
            y_var, y_pct, y_material = var, pct, material
        else:
            y_var = float(row["actual"].iloc[0] - row["budget"].iloc[0])
            y_pct, _y_fu, y_material = verdict(y_var, float(row["budget"].iloc[0]))

        base = _two_horizon((var, pct, material), (y_var, y_pct, y_material), word)
        if material or y_material:
            pool = children if (material and not y_material) else y_children
            base += movers(pool if pool is not None else children, key)
        comments[key] = base.strip()
    return comments


def total_comment(
    detail: pd.DataFrame,
    child: str,
    materiality: MaterialityRule | None = None,
    higher_is_better: bool = False,
    child_names: dict | None = None,
    ytd_detail: pd.DataFrame | None = None,
) -> str:
    """The same commentary for a sheet's grand-total row.

    A total is a roll-up like any other - it has the whole sheet underneath it -
    and it is the row a reader looks at first. Leaving it as the one blank cell
    in a commented column reads as a gap rather than a decision.
    """
    label = "\u0000total"          # a key no real dimension value can collide with
    frames = []
    for frame in (detail, ytd_detail):
        if frame is None:
            frames.append(None); continue
        out = frame.copy(); out[label] = label
        frames.append(out)
    return rollup_comments(frames[0], label, child, materiality, higher_is_better,
                           child_names, frames[1]).get(label, "")


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

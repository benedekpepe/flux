"""
Core variance engine.

Turns a flat, leaf-level GL export into a management-ready P&L with:

  - roll-up from accounts to category subtotals and computed subtotals
    (Gross profit, EBIT, Net income) per PNL_STRUCTURE
  - actual vs budget variance in absolute and percentage terms
  - actual vs prior-year variance
  - favourable / unfavourable classification that respects account type
    (revenue up = good, cost up = bad, profit subtotals up = good)
  - a two-condition materiality flag: an item is flagged only when the
    absolute variance clears an EUR floor AND the percentage variance clears
    a % floor, so a big % on a tiny base is not treated as material.

The output is a tidy DataFrame, one row per report line, ordered as the P&L.
"""

from __future__ import annotations
from dataclasses import dataclass
import pandas as pd

from .coa import (
    PNL_STRUCTURE,
    CATEGORY_FAVOURABLE,
    FAV_HIGHER,
    FAV_LOWER,
)


@dataclass(frozen=True)
class MaterialityRule:
    """An item is material only if BOTH thresholds are cleared."""
    abs_threshold: float = 25_000.0   # EUR
    pct_threshold: float = 0.10       # 10%


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

# A percentage stops carrying information once the variance dwarfs the base:
# a EUR 150k swing against a EUR 7k plan is -2,000%, arithmetically true and
# analytically useless. Such lines are marked n/m and judged on the amount.
# The cap is deliberately loose - a small but real base (a EUR 23k plan) still
# deserves a percentage; only near-zero bases are suppressed.
NM_RATIO_CAP = 10.0        # variance more than 1,000% of the base
NM_ABS_FLOOR = 100.0       # base below this is treated as zero


def _pct(numerator: float, base: float, floor: float = 0.0) -> float | None:
    """Variance percentage against |base|, or None when not meaningful."""
    if base is None or abs(base) < max(NM_ABS_FLOOR, 1e-9):
        return None
    ratio = numerator / abs(base)
    return None if abs(ratio) > NM_RATIO_CAP else ratio


def _favourable(variance: float, favourable_direction: str) -> bool:
    """True if the variance is a good outcome given the line's direction."""
    if favourable_direction == FAV_HIGHER:
        return variance >= 0
    return variance <= 0


def _category_totals(gl: pd.DataFrame, category: str) -> tuple[float, float, float]:
    """Sum actual/budget/prior_year for one category (positive magnitudes)."""
    sub = gl[gl["category"] == category]
    return (
        float(sub["actual"].sum()),
        float(sub["budget"].sum()),
        float(sub["prior_year"].sum()),
    )


def has_budget(gl: pd.DataFrame) -> bool:
    """True when the frame actually carries a plan to compare against.

    A ledger extract on its own has no budget column, or an empty one. Treating
    that as "budget = 0" would make every line read as a 100% overspend, so the
    variance columns are suppressed instead of computed against nothing.
    """
    if "budget" not in gl.columns:
        return False
    return bool(pd.to_numeric(gl["budget"], errors="coerce").abs().sum() > 0)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def build_report(
    gl: pd.DataFrame,
    materiality: MaterialityRule | None = None,
    budgeted: bool | None = None,
) -> pd.DataFrame:
    """Build the management P&L with variances and flags from a leaf-level GL.

    Parameters
    ----------
    gl : DataFrame with columns
         account_code, account_name, category, actual, budget, prior_year
    materiality : thresholds for flagging; defaults to 25k EUR AND 10%.
    budgeted : whether a plan exists. Detected from the data when omitted. With
        no budget the variance, F/U and materiality columns come back empty
        rather than measured against zero, which would flag every line.
    """
    materiality = materiality or MaterialityRule()
    budgeted = has_budget(gl) if budgeted is None else bool(budgeted)

    # Running store of computed line values, so computed subtotals can
    # reference earlier lines by label.
    values: dict[str, dict[str, float]] = {}
    records: list[dict] = []

    for line in PNL_STRUCTURE:
        if line.kind == "category":
            actual, budget, prior = _category_totals(gl, line.category)
        elif line.kind == "computed":
            actual = budget = prior = 0.0
            for sign, ref in line.components:
                s = 1.0 if sign == "+" else -1.0
                actual += s * values[ref]["actual"]
                budget += s * values[ref]["budget"]
                prior += s * values[ref]["prior_year"]
        else:  # pragma: no cover - guard
            raise ValueError(f"Unknown line kind: {line.kind}")

        values[line.label] = {"actual": actual, "budget": budget, "prior_year": prior}

        var_bud = actual - budget
        var_py = actual - prior
        var_bud_pct = _pct(var_bud, budget)
        var_py_pct = _pct(var_py, prior)

        favourable = _favourable(var_bud, line.favourable)

        # Two-condition materiality: needs BOTH the EUR and the % floor. When the
        # percentage is not meaningful (immaterial base), the EUR test decides.
        big_enough = abs(var_bud) >= materiality.abs_threshold
        material = big_enough and (
            var_bud_pct is None or abs(var_bud_pct) >= materiality.pct_threshold
        )

        records.append(
            {
                "line": line.label,
                "kind": line.kind,
                "actual": round(actual, 2),
                "budget": round(budget, 2) if budgeted else None,
                "prior_year": round(prior, 2),
                "var_bud": round(var_bud, 2) if budgeted else None,
                "var_bud_pct": (None if var_bud_pct is None or not budgeted
                                else round(var_bud_pct, 4)),
                "var_py": round(var_py, 2),
                "var_py_pct": None if var_py_pct is None else round(var_py_pct, 4),
                "fav_unfav": ("F" if favourable else "U") if budgeted else "",
                "material": bool(material) if budgeted else False,
            }
        )

    out = pd.DataFrame.from_records(records)
    out.attrs["budgeted"] = budgeted
    return out


def leaf_variances(
    gl: pd.DataFrame,
    materiality: MaterialityRule | None = None,
    budgeted: bool | None = None,
) -> pd.DataFrame:
    """Account-level variance detail with F/U and materiality flags.

    Used to drive line-item commentary: these are the individual accounts
    that moved, not the subtotals. With no budget the frame comes back with the
    variance columns present but empty, so callers can rely on the shape.
    """
    materiality = materiality or MaterialityRule()
    budgeted = has_budget(gl) if budgeted is None else bool(budgeted)
    out = gl.copy()
    if not budgeted:
        out["var_bud"] = None
        out["var_bud_pct"] = None
        out["fav_unfav"] = ""
        out["material"] = False
        out.attrs["budgeted"] = False
        return out

    out["var_bud"] = out["actual"] - out["budget"]
    out["var_bud_pct"] = out.apply(
        lambda r: _pct(r["var_bud"], r["budget"]), axis=1
    )
    out["fav_unfav"] = out.apply(
        lambda r: "F"
        if _favourable(r["var_bud"], CATEGORY_FAVOURABLE.get(r["category"], FAV_LOWER))
        else "U",
        axis=1,
    )
    out["material"] = out.apply(
        lambda r: abs(r["var_bud"]) >= materiality.abs_threshold
        and (r["var_bud_pct"] is None
             or abs(r["var_bud_pct"]) >= materiality.pct_threshold),
        axis=1,
    )
    return out.sort_values("var_bud", key=lambda s: s.abs(), ascending=False)


def department_variances(
    detail: pd.DataFrame,
    materiality: MaterialityRule | None = None,
) -> pd.DataFrame:
    """Spend variance by department (the roll-up level).

    Cost accounts only (COGS, OpEx, Other); revenue is not a departmental spend
    budget. All spend is "lower is better", so favourable means under budget.
    """
    return _spend_variances(detail, "department", materiality)


def cost_centre_variances(
    detail: pd.DataFrame,
    materiality: MaterialityRule | None = None,
) -> pd.DataFrame:
    """Spend variance by cost centre, carrying its parent department."""
    out = _spend_variances(detail, "cost_centre", materiality)
    if "department" in detail.columns:
        parent = (detail[["cost_centre", "department"]]
                  .drop_duplicates("cost_centre").set_index("cost_centre")["department"])
        out.insert(0, "department", out["cost_centre"].map(parent))
        out = out.sort_values(["department", "budget"], ascending=[True, False]).reset_index(drop=True)
    return out


def _spend_variances(detail, level: str, materiality: MaterialityRule | None = None):
    materiality = materiality or MaterialityRule()
    spend = detail[detail["category"] != "Revenue"]
    grp = spend.groupby(level, as_index=False)[["actual", "budget", "prior_year"]].sum()
    grp["var_bud"] = grp["actual"] - grp["budget"]
    grp["var_bud_pct"] = grp.apply(lambda r: _pct(r["var_bud"], r["budget"]), axis=1)
    grp["fav_unfav"] = grp["var_bud"].apply(lambda v: "F" if v <= 0 else "U")
    # Same two-condition rule as build_report: the EUR floor must be cleared,
    # and the percentage must clear its floor or be not meaningful.
    grp["material"] = grp.apply(
        lambda r: abs(r["var_bud"]) >= materiality.abs_threshold
        and (r["var_bud_pct"] is None or pd.isna(r["var_bud_pct"])
             or abs(r["var_bud_pct"]) >= materiality.pct_threshold),
        axis=1,
    )
    return grp.sort_values("budget", ascending=False).reset_index(drop=True)


def entity_variances(
    detail: pd.DataFrame,
    materiality: MaterialityRule | None = None,
) -> pd.DataFrame:
    """Net income by legal entity (consolidation view).

    For each entity: revenue - all spend, actual vs budget. Net income is
    "higher is better".
    """
    materiality = materiality or MaterialityRule()
    rows = []
    for ent, sub in detail.groupby("entity"):
        rev_a = sub[sub.category == "Revenue"]["actual"].sum()
        rev_b = sub[sub.category == "Revenue"]["budget"].sum()
        spend_a = sub[sub.category != "Revenue"]["actual"].sum()
        spend_b = sub[sub.category != "Revenue"]["budget"].sum()
        net_a = rev_a - spend_a
        net_b = rev_b - spend_b
        var = net_a - net_b
        pct = _pct(var, net_b)
        rows.append({
            "entity": ent,
            "revenue": rev_a, "spend": spend_a,
            "net_actual": net_a, "net_budget": net_b,
            "var_bud": var, "var_bud_pct": pct,
            "fav_unfav": "F" if var >= 0 else "U",
        })
    return pd.DataFrame(rows).sort_values("revenue", ascending=False).reset_index(drop=True)


if __name__ == "__main__":
    from .synthetic_data import generate_month

    gl = generate_month()
    report = build_report(gl)
    pd.set_option("display.max_columns", None)
    pd.set_option("display.width", 160)
    print(report.to_string(index=False))

"""
End-to-end demo: generate synthetic data, build the report, print it.

Run:  python scripts/demo.py
"""

from __future__ import annotations
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from flux.synthetic_data import generate_month
from flux.engine import build_report, leaf_variances
from flux.commentary import generate_commentary


def _eur(x: float) -> str:
    return f"{x:,.0f}"


def _pct(x) -> str:
    if x is None or pd.isna(x):
        return "  n/a"
    return f"{x*100:+5.1f}%"


def print_pnl(report: pd.DataFrame, period: str) -> None:
    subtotal_lines = {"Gross profit", "Operating income (EBIT)", "Net income"}

    header = (
        f"{'':32}{'Actual':>13}{'Budget':>13}"
        f"{'Var':>13}{'Var %':>8}{'  ':>2}{'vs PY %':>9}"
    )
    print(f"\n  MANAGEMENT P&L  -  {period}  (EUR)")
    print("  " + "-" * (len(header)))
    print("  " + header)
    print("  " + "-" * (len(header)))

    for _, r in report.iterrows():
        name = r["line"]
        is_sub = name in subtotal_lines
        label = name.upper() if is_sub else name
        flag = " *" if r["material"] else "  "
        fu = r["fav_unfav"]
        line = (
            f"  {label:30}"
            f"{_eur(r['actual']):>13}"
            f"{_eur(r['budget']):>13}"
            f"{_eur(r['var_bud']):>13}"
            f"{_pct(r['var_bud_pct']):>8}"
            f"{flag} {fu}"
            f"{_pct(r['var_py_pct']):>9}"
        )
        if is_sub:
            print("  " + "-" * (len(header)))
        print(line)
    print("  " + "-" * (len(header)))
    print("  * = material variance (>= 25,000 EUR AND >= 10%)   F/U vs budget\n")


def print_material_leaves(gl: pd.DataFrame) -> None:
    leaves = leaf_variances(gl)
    flagged = leaves[leaves["material"]]
    print("  MATERIAL ACCOUNT-LEVEL VARIANCES (drivers for commentary)")
    print("  " + "-" * 74)
    for _, r in flagged.iterrows():
        print(
            f"  {r['account_code']}  {r['account_name']:26}"
            f"{_eur(r['var_bud']):>12}"
            f"{r['var_bud_pct']*100:>8.1f}%   "
            f"{'FAVOURABLE' if r['fav_unfav']=='F' else 'UNFAVOURABLE'}"
        )
    print()


if __name__ == "__main__":
    period = "2025-06"
    gl = generate_month(period=period)
    report = build_report(gl)

    print_pnl(report, period)
    print_material_leaves(gl)

    print("  MANAGEMENT COMMENTARY")
    print("  " + "-" * 74)
    for para in generate_commentary(report, gl).split("\n\n"):
        # simple wrap for console readability
        words, line = para.split(), ""
        for w in words:
            if len(line) + len(w) + 1 > 72:
                print("  " + line)
                line = w
            else:
                line = f"{line} {w}".strip()
        if line:
            print("  " + line)
        print()

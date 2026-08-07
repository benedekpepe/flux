"""
Excel formula builders.

Every reported figure in a Flux pack is a formula over the input sheet, so the
workbook recalculates when an input is edited. This module builds those formula
strings; it holds no styling and touches no worksheet.

The materiality thresholds live in two named lever cells on the P&L sheet. Every
other sheet points at those cells, so changing one number re-flags the whole
pack.
"""

from __future__ import annotations


# ---------------------------------------------------------------------------
# Materiality levers
# ---------------------------------------------------------------------------
LEVER_EUR = "C9"
LEVER_PCT = "G9"
LEV_E = "$C$9"
LEV_P = "$G$9"

# Absolute references for sheets other than the P&L, which own the levers.
PNL_SHEET = "P&L Report"
PL_E = f"'{PNL_SHEET}'!{LEV_E}"
PL_P = f"'{PNL_SHEET}'!{LEV_P}"

DEFAULT_ABS_THRESHOLD = 25_000
DEFAULT_PCT_THRESHOLD = 0.10


# ---------------------------------------------------------------------------
# "Not meaningful" thresholds - mirrored from engine.py so the workbook and the
# Python engine reach the same verdict on the same numbers.
# ---------------------------------------------------------------------------
NM_RATIO_CAP = 10          # variance above 1,000% of the base
NM_ABS_FLOOR = 100         # base below this counts as zero


def col_range(sheet: str, col: str, first: int, last: int) -> str:
    """An absolute single-column range on another sheet."""
    return f"'{sheet}'!${col}${first}:${col}${last}"


def pct_f(var_cell: str, base_cell: str) -> str:
    """Variance % against a base, marked n/m only when it stops being meaningful.

    A near-zero plan produces an arithmetically correct but useless percentage,
    so those lines read n/m and are judged on the amount instead. A small but
    real base still gets a percentage.
    """
    return (f'=IF(ABS({base_cell})<{NM_ABS_FLOOR},"n/m",'
            f'IF(ABS({var_cell}/{base_cell})>{NM_RATIO_CAP},"n/m",'
            f'{var_cell}/ABS({base_cell})))')


def flag_f(var_cell: str, pct_cell: str, lever_eur: str, lever_pct: str) -> str:
    """Material when the amount clears the EUR floor and either the percentage
    clears its floor or the percentage is not meaningful."""
    return (f'=IF(ABS({var_cell})<{lever_eur},"",'
            f'IF(NOT(ISNUMBER({pct_cell})),"MATERIAL",'
            f'IF(ABS({pct_cell})>={lever_pct},"MATERIAL","")))')


def fu_f(var_cell: str, higher_is_better: bool) -> str:
    """Favourable / unfavourable badge for a variance cell."""
    return (f'=IF({var_cell}>=0,"F","U")' if higher_is_better
            else f'=IF({var_cell}<=0,"F","U")')


class Variance:
    """Gate for the variance columns.

    A ledger extract without a plan has nothing to compare against. Writing
    `actual - 0` into the variance column would produce a full set of
    confident-looking numbers that all say "100% over budget", so when there is
    no budget the variance cells are left empty instead and the sheet says why.
    """

    def __init__(self, has_budget: bool):
        self.on = bool(has_budget)

    def __bool__(self) -> bool:
        return self.on

    def cell(self, formula: str):
        """The formula when a budget exists, otherwise a blank cell."""
        return formula if self.on else None

    def pct(self, var_cell: str, base_cell: str):
        return self.cell(pct_f(var_cell, base_cell))

    def flag(self, var_cell: str, pct_cell: str, lever_eur: str, lever_pct: str):
        return self.cell(flag_f(var_cell, pct_cell, lever_eur, lever_pct))

    def fu(self, var_cell: str, higher_is_better: bool):
        return self.cell(fu_f(var_cell, higher_is_better))


NO_BUDGET_NOTE = (
    "No budget was supplied with this file, so the variance, F/U and materiality "
    "columns are intentionally left empty rather than compared against zero. "
    "Upload a budget file to populate them."
)

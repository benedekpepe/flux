"""
Demo pack: the full multi-entity showcase, built from generated data.

Inputs (how a real finance function holds data):
  - GL Transactions : year-to-date transaction-level actuals, multi-entity,
                      multi-currency (with FX to EUR), tagged with cost centre,
                      expense type, posting date and source journal.
  - Budget          : full-year summarised plan (and prior year) by entity x
                      cost centre x account x period.

Every reporting sheet is a live formula view over those two inputs, and all
views reconcile to the source. They also share one set of columns - the month,
the year to date, the full-year plan and the run-rate projection against it,
then F/U and the materiality flag - so the sheets differ in what they cut the
ledger by, never in how they report it:
  - P&L Report     consolidated P&L, KPI cards, per-line commentary, and the
                   three lever cells the rest of the pack points at
  - Expense Report natural view: expense type
  - By Entity      consolidation: net income per legal entity
  - Departments & CCs  functional view: department roll-up with cost centres
  - Drivers        account-level movers with data bars

Output: output/flux_demo_pack.xlsx
"""

from __future__ import annotations
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from ..coa import (PNL_STRUCTURE, CATEGORY_FAVOURABLE, FAV_HIGHER,
                   SPEND_DEPARTMENTS, DEPARTMENT_COST_CENTRES, EXPENSE_GROUPS)
from ..commentary import line_comments
from ..engine import build_report
from .styling import (
    FONT, GREEN_INK, RED_INK,
    F_HEAD, F_BODY, F_SMALL, F_SUB, F_INPUT, F_KPI_LABEL, F_KPI_VALUE, F_NOTE,
    FILL_HEAD, FILL_IVORY, FILL_BAND, FILL_WHITE,
    CUR2, CUR_EUR, CUR2_EUR, PCT, RATE, KPI_DELTA, LCY_FORMATS,
    CENTER, LEFT, RIGHT, WRAP, indent, band_fill,
    GOLD_SIDE, HEADER_BOTTOM, SUBTOTAL_TOP, TOTAL_TOP,
    hide_grid, title_band, headers, widths, outline, lever,
    wrapped_height, fit_text_columns,
    named_style,
    quiet_indicators, collect_quiet_ranges, suppress_error_indicators,
)
# The columns every reporting sheet shares are written by `rows`, so a sheet
# here only supplies the five figures that depend on how it cuts the ledger.
from .rows import (report_cf as _report_cf, write_sum_tail as _sum_tail,
                   write_tail as _tail)
from .formulas import (
    Layout,
    LEVER_EUR, LEVER_PCT, LEVER_MONTHS, LEV_E, LEV_P, LEV_M,
    DEFAULT_ABS_THRESHOLD, DEFAULT_PCT_THRESHOLD,
)


# Width of the commentary column, shared by the column setup and the row-height
# estimate so the two cannot drift apart.
COMMENT_WIDTH = 58

# ---- source-sheet column letters ------------------------------------------
# ---- source-sheet layout -----------------------------------------------------
# (key, header, width, format) - column letters are derived, never hand-counted.
GL_SPEC = [
    ("doc_no", "Document", 12, "text"),
    ("doc_type", "Type", 6, "text"),
    ("doc_type_text", "Document type", 18, "text"),
    ("fiscal_year", "FY", 7, "int"),
    ("period", "Period", 9, "text"),
    ("period_no", "Per. key", 9, "int"),
    ("doc_date", "Doc date", 11, "text"),
    ("posting_date", "Posting date", 12, "text"),
    ("entry_date", "Entry date", 11, "text"),
    ("entry_time", "Time", 7, "text"),
    ("reference", "Reference", 14, "text"),
    ("header_text", "Header text", 24, "text"),
    ("created_by", "Created by", 12, "text"),
    ("reversed", "Rev.", 6, "text"),
    ("entity", "Entity", 12, "text"),
    ("company_code", "Co. code", 9, "text"),
    ("region", "Region", 10, "text"),
    ("department", "Department", 13, "text"),
    ("cost_centre", "Cost centre", 12, "text"),
    ("profit_centre", "Profit centre", 13, "text"),
    ("line_no", "Line", 6, "int"),
    ("account_code", "Account", 9, "text"),
    ("account_name", "Account name", 26, "text"),
    ("category", "Category", 11, "text"),
    ("expense_type", "Expense type", 22, "text"),
    ("group", "P&L group", 20, "text"),
    ("dc_indicator", "D/C", 6, "text"),
    ("currency", "Ccy", 6, "text"),
    ("amount_lcy", "Amount (local)", 15, "lcy"),
    ("fx_rate", "FX", 9, "rate"),
    ("amount_eur", "Amount (\u20ac)", 15, "cur2"),
    ("tax_code", "Tax", 6, "text"),
    ("partner", "Partner", 24, "text"),
    ("assignment", "Assignment", 18, "text"),
    ("line_text", "Item text", 44, "text"),
    ("source", "Source", 9, "text"),
]

BUD_SPEC = [
    ("period", "Period", 9, "text"),
    ("period_no", "Per. key", 9, "int"),
    ("version", "Version", 12, "text"),
    ("entity", "Entity", 13, "text"),
    ("region", "Region", 10, "text"),
    ("currency", "Ccy", 6, "text"),
    ("department", "Department", 13, "text"),
    ("cost_centre", "Cost centre", 12, "text"),
    ("account_code", "Account", 9, "text"),
    ("account_name", "Account name", 26, "text"),
    ("category", "Category", 11, "text"),
    ("expense_type", "Expense type", 22, "text"),
    ("group", "P&L group", 20, "text"),
    ("owner", "Budget owner", 14, "text"),
    ("budget_eur", "Budget", 14, "cur"),
    ("prior_eur", "Prior Yr Act", 15, "cur"),
]


def _letters(spec):
    return {key: get_column_letter(i) for i, (key, *_rest) in enumerate(spec, start=1)}


GL_COL = _letters(GL_SPEC)
BUD_COL = _letters(BUD_SPEC)

GL_PERIOD, GL_ENT, GL_DEPT = GL_COL["period"], GL_COL["entity"], GL_COL["department"]
GL_PERNO = GL_COL["period_no"]
GL_CC, GL_CODE, GL_CAT = GL_COL["cost_centre"], GL_COL["account_code"], GL_COL["category"]
GL_EXP, GL_EUR = GL_COL["expense_type"], GL_COL["amount_eur"]
BUD_PERIOD, BUD_ENT, BUD_DEPT = BUD_COL["period"], BUD_COL["entity"], BUD_COL["department"]
BUD_PERNO = BUD_COL["period_no"]
BUD_CC, BUD_CODE, BUD_CAT = BUD_COL["cost_centre"], BUD_COL["account_code"], BUD_COL["category"]
BUD_EXP, BUD_BUD, BUD_PRIOR = BUD_COL["expense_type"], BUD_COL["budget_eur"], BUD_COL["prior_eur"]

# ===========================================================================
# Input sheets
# ===========================================================================
def _write_source_sheet(ws, df, spec, title, meta, footnote) -> tuple[int, int]:
    """Write an input sheet from a column spec; returns (first_row, last_row)."""
    hide_grid(ws)
    last_col = get_column_letter(len(spec))
    title_band(ws, title, meta, last_col)

    numeric = {"cur", "cur2", "lcy", "rate", "int"}
    hrow = 5
    for i, (key, header, width, fmt) in enumerate(spec, start=1):
        cell = ws.cell(row=hrow, column=i, value=header)
        cell.font = F_HEAD; cell.fill = FILL_HEAD
        cell.alignment = CENTER if fmt in numeric else LEFT
        cell.border = HEADER_BOTTOM
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[hrow].height = 22

    fmt_map = {"cur": CUR_EUR, "cur2": CUR2_EUR, "rate": RATE, "int": "0"}
    wb = ws.parent

    # Resolved names are memoised: openpyxl's `name in wb.named_styles` rebuilds
    # the name list on every call, which at a quarter of a million cells costs
    # more than the styling it was meant to save.
    style_cache: dict[tuple, str] = {}

    def style_for(fmt, banded, ccy="EUR"):
        """One registered style per (format, band, currency) combination.

        Roughly a dozen styles cover a sheet of any length, so the per-cell work
        drops from four style writes to one dictionary hit.
        """
        cached = style_cache.get((fmt, banded, ccy))
        if cached is not None:
            return cached
        suffix = "b" if banded else "w"
        fill = FILL_BAND if banded else FILL_WHITE
        if fmt == "lcy":
            name = named_style(wb, f"flux_lcy_{ccy}_{suffix}", font=F_INPUT,
                               fill=fill, alignment=RIGHT,
                               number_format=LCY_FORMATS.get(ccy, CUR2))
        elif fmt in ("cur", "cur2"):
            name = named_style(wb, f"flux_{fmt}_{suffix}", font=F_INPUT, fill=fill,
                               alignment=RIGHT, number_format=fmt_map[fmt])
        elif fmt in ("rate", "int"):
            name = named_style(wb, f"flux_{fmt}_{suffix}", font=F_SMALL, fill=fill,
                               alignment=RIGHT, number_format=fmt_map[fmt])
        else:
            name = named_style(wb, f"flux_text_{suffix}", font=F_SMALL, fill=fill,
                               alignment=LEFT)
        style_cache[(fmt, banded, ccy)] = name
        return name

    first = hrow + 1
    # Plain dicts: `row.get(key)` on a Series goes through pandas indexing on
    # every cell, and there are a quarter of a million of them here.
    for i, row in enumerate(df.to_dict("records")):
        r = first + i
        banded = bool(i % 2)
        for c, (key, _h, _w, fmt) in enumerate(spec, start=1):
            v = row.get(key, "")
            if fmt in numeric and v != "":
                v = float(v) if fmt != "int" else int(v)
            cell = ws.cell(row=r, column=c, value=v)
            cell.style = style_for(fmt, banded, str(row.get("currency", "EUR")))
        ws.row_dimensions[r].height = 14
    last = first + len(df) - 1

    text_cols = [get_column_letter(i) for i, (_k, _h, _w, fmt) in enumerate(spec, start=1)
                 if fmt not in numeric]
    fit_text_columns(ws, text_cols, first, last)

    ws.cell(row=last + 2, column=1, value=footnote).font = F_NOTE
    quiet_indicators(ws, 5, last)
    ws.freeze_panes = f"A{first}"
    return first, last


def _write_gl_transactions(ws, txns, period) -> tuple[int, int]:
    return _write_source_sheet(
        ws, txns, GL_SPEC,
        "GL Transactions · year-to-date posting lines (multi-entity, multi-currency)",
        f"Year to date through {period}  ·  reporting currency \u20ac",
        "Blue cells are input amounts. Amount (EUR) = Amount (LCY) x FX. "
        "Revenue is credited (H), costs are debited (S). Only account, category, "
        "cost centre, period and the EUR amount drive the reports; the remaining "
        "columns are carried for traceability.",
    )


def _write_budget(ws, bud, period) -> tuple[int, int]:
    return _write_source_sheet(
        ws, bud, BUD_SPEC,
        "Budget · full-year plan by entity x department x cost centre x account",
        "Full year 2025  ·  \u20ac",
        "Blue cells are the plan (budget) and prior-year actual figures.",
    )


# ===========================================================================
# P&L Report
# ===========================================================================
def _write_pnl(ws, gl, glf, gll, bud, budf, budl, period, comments, report) -> None:
    """The management P&L, and the sheet that owns the pack's three levers.

    Month against budget answers what happened; year to date answers whether it
    is a pattern; the run rate answers where the year lands if it continues. A
    reader asks those three questions in that order, so they sit in that order,
    and every other sheet in the pack repeats the same columns for its own cut
    of the same ledger.
    """
    hide_grid(ws)
    L = Layout(1)
    com = get_column_letter(L.ncols + 1)
    title_band(ws, "Management P&L · Variance Report (consolidated)",
               f"Reporting month {period}  ·  with year to date and full-year run rate"
               f"  ·  \u20ac", com)
    gR = lambda col: f"'{gl}'!${col}${glf}:${col}${gll}"
    bR = lambda col: f"'{bud}'!${col}${budf}:${col}${budl}"
    perno = int(period[:4]) * 100 + int(period[5:7])
    mflt = f',{gR(GL_PERIOD)},"{period}"'
    mfltb = f',{bR(BUD_PERIOD)},"{period}"'
    yflt = f',{gR(GL_PERNO)},"<={perno}"'
    yfltb = f',{bR(BUD_PERNO)},"<={perno}"'

    # ---- the three levers every sheet points at ----
    lever(ws, "A9:B9", "Materiality floor (\u20ac)", LEVER_EUR,
          DEFAULT_ABS_THRESHOLD, CUR_EUR)
    lever(ws, "E9:F9", "Materiality floor (%)", LEVER_PCT,
          DEFAULT_PCT_THRESHOLD, PCT)
    lever(ws, "I9:J9", "Months elapsed", LEVER_MONTHS, int(period[5:7]), "0")
    ws.merge_cells(f"M9:{com}9")
    ws["M9"] = ("F/U judges the month variance; Flag names which timeframe "
                "clears both floors.")
    ws["M9"].font = F_NOTE; ws["M9"].alignment = LEFT

    headers(ws, 10, L.headers("") + ["Commentary"], center_from=2,
            center_to=L.ncols)

    subtotal_labels = {"Gross profit", "Operating income (EBIT)", "Net income"}
    label_row, first, r, cat_idx = {}, 11, 11, 0
    rows_higher, rows_lower = [], []
    for line in PNL_STRUCTURE:
        c = L.row(r)
        is_sub = line.label in subtotal_labels
        higher = line.favourable == FAV_HIGHER
        ws[f"A{r}"] = line.label
        if line.kind == "category":
            ws[c["act"]] = f'=SUMIFS({gR(GL_EUR)},{gR(GL_CAT)},"{line.category}"{mflt})'
            ws[c["bud"]] = f'=SUMIFS({bR(BUD_BUD)},{bR(BUD_CAT)},"{line.category}"{mfltb})'
            ws[c["yact"]] = f'=SUMIFS({gR(GL_EUR)},{gR(GL_CAT)},"{line.category}"{yflt})'
            ws[c["ybud"]] = f'=SUMIFS({bR(BUD_BUD)},{bR(BUD_CAT)},"{line.category}"{yfltb})'
            # No period criterion: the full-year plan is the whole budget sheet.
            ws[c["fybud"]] = f'=SUMIF({bR(BUD_CAT)},"{line.category}",{bR(BUD_BUD)})'
        else:
            parts = {key: [] for key in ("act", "bud", "yact", "ybud", "fybud")}
            for sign, ref in line.components:
                rr = label_row[ref]
                for key in parts:
                    parts[key].append(f"{sign}{getattr(L, key)}{rr}")
            for key, terms in parts.items():
                ws[c[key]] = "=" + "".join(terms)
        _tail(ws, L, r, higher=higher, bold=is_sub,
              lev_e=LEV_E, lev_p=LEV_P, months=LEV_M)

        base_fill = FILL_IVORY if is_sub else band_fill(cat_idx)
        if not is_sub:
            cat_idx += 1
        for col in L.span() + [com]:
            ws[f"{col}{r}"].fill = base_fill
        ws[f"A{r}"].font = F_SUB if is_sub else F_BODY
        ws[f"A{r}"].alignment = LEFT
        comment = comments.get(line.label, "")
        kc = ws[f"{com}{r}"]; kc.value = comment; kc.alignment = WRAP
        kc.font = F_SUB if is_sub else F_BODY
        if is_sub:
            for col in L.span() + [com]:
                ws[f"{col}{r}"].border = SUBTOTAL_TOP
        (rows_higher if higher else rows_lower).append(r)
        ws.row_dimensions[r].height = wrapped_height(comment, COMMENT_WIDTH)
        label_row[line.label] = r; r += 1
    last = r - 1
    _report_cf(ws, L, first, last, rows_higher, rows_lower)

    # ---- KPI cards ----
    cards = [("REVENUE", "Revenue"), ("OPERATING INCOME (EBIT)", "Operating income (EBIT)"),
             ("NET INCOME", "Net income")]
    for (title, key), (c1, c2) in zip(cards, [(1, 3), (5, 7), (9, 11)]):
        lr = label_row[key]; fav = report[report["line"] == key].iloc[0]["fav_unfav"] == "F"
        Lc = get_column_letter(c1); Rt = get_column_letter(c2)
        for rr in (5, 6, 7):
            ws.merge_cells(f"{Lc}{rr}:{Rt}{rr}")
            for cc in range(c1, c2 + 1):
                ws.cell(row=rr, column=cc).fill = FILL_IVORY
        ws[f"{Lc}5"] = title; ws[f"{Lc}5"].font = F_KPI_LABEL
        ws[f"{Lc}5"].alignment = LEFT
        ws[f"{Lc}6"] = f"={L.act}{lr}"; ws[f"{Lc}6"].font = F_KPI_VALUE
        ws[f"{Lc}6"].number_format = CUR_EUR
        ws[f"{Lc}6"].alignment = LEFT
        ws[f"{Lc}7"] = f"={L.pct}{lr}"
        ws[f"{Lc}7"].number_format = KPI_DELTA
        ws[f"{Lc}7"].font = Font(name=FONT, size=9, bold=True,
                                 color=GREEN_INK if fav else RED_INK)
        ws[f"{Lc}7"].alignment = LEFT
        outline(ws, 5, c1, 7, c2, GOLD_SIDE)
    ws.row_dimensions[5].height = 18; ws.row_dimensions[6].height = 26
    ws.row_dimensions[7].height = 16

    widths(ws, L.widths(26) + [COMMENT_WIDTH])
    quiet_indicators(ws, 5, last)
    ws.freeze_panes = f"A{first}"


# ===========================================================================
# Expense Report (natural view: expense type)
# ===========================================================================
def _write_expense_report(ws, gl, glf, gll, bud, budf, budl, period) -> None:
    hide_grid(ws)
    L = Layout(1)
    # The fiscal year comes from the period, not a constant: hardcoded, this
    # header still read "FY 2025" for a 2026 close.
    title_band(ws, "Expense Report · by expense type",
               f"Reporting month {period}  ·  YTD & FY {period[:4]}  ·  \u20ac",
               L.last_col)
    gR = lambda col: f"'{gl}'!${col}${glf}:${col}${gll}"
    bR = lambda col: f"'{bud}'!${col}${budf}:${col}${budl}"
    perno = int(period[:4]) * 100 + int(period[5:7])

    headers(ws, 5, L.headers("Expense Type"), center_from=2, center_to=L.ncols)

    first, r, i = 6, 6, 0
    group_rows, all_rows = [], []
    for group_name, types in EXPENSE_GROUPS:
        # A group with a single expense type needs no separate subtotal row.
        single = len(types) == 1
        type_rows = []
        for et in types:
            c = L.row(r)
            band = FILL_IVORY if single else band_fill(i)
            lab = ws.cell(row=r, column=1, value=et)
            lab.font = F_SUB if single else F_BODY
            lab.alignment = LEFT if single else indent(1)
            ws[c["act"]] = (f'=SUMIFS({gR(GL_EUR)},{gR(GL_EXP)},"{et}",'
                            f'{gR(GL_PERIOD)},"{period}")')
            ws[c["bud"]] = (f'=SUMIFS({bR(BUD_BUD)},{bR(BUD_EXP)},"{et}",'
                            f'{bR(BUD_PERIOD)},"{period}")')
            ws[c["yact"]] = (f'=SUMIFS({gR(GL_EUR)},{gR(GL_EXP)},"{et}",'
                             f'{gR(GL_PERNO)},"<={perno}")')
            ws[c["ybud"]] = (f'=SUMIFS({bR(BUD_BUD)},{bR(BUD_EXP)},"{et}",'
                             f'{bR(BUD_PERNO)},"<={perno}")')
            ws[c["fybud"]] = f'=SUMIF({bR(BUD_EXP)},"{et}",{bR(BUD_BUD)})'
            _tail(ws, L, r, higher=False, bold=single)
            for col in L.span():
                cell = ws[f"{col}{r}"]; cell.fill = band
                if single:
                    cell.border = SUBTOTAL_TOP
            ws.row_dimensions[r].height = 21 if single else 18
            type_rows.append(r); all_rows.append(r)
            i += 1; r += 1

        if single:
            group_rows.append(type_rows[0])
            i = 0
            continue

        ws.cell(row=r, column=1, value=group_name).font = F_SUB
        ws.cell(row=r, column=1).alignment = LEFT
        _sum_tail(ws, L, r, type_rows, higher=False)
        for col in L.span():
            cell = ws[f"{col}{r}"]; cell.fill = FILL_IVORY
            cell.border = SUBTOTAL_TOP
        ws.row_dimensions[r].height = 21
        group_rows.append(r); all_rows.append(r)
        i = 0
        r += 1
    last = r - 1

    ws.cell(row=r, column=1, value="Total expenses").font = F_SUB
    ws.cell(row=r, column=1).alignment = LEFT
    _sum_tail(ws, L, r, group_rows, higher=False)
    for col in L.span():
        ws[f"{col}{r}"].fill = FILL_IVORY
        ws[f"{col}{r}"].border = TOTAL_TOP
    ws.row_dimensions[r].height = 24
    all_rows.append(r)

    _report_cf(ws, L, first, r, [], all_rows)
    ws.cell(row=r + 2, column=1,
            value="Expense types are grouped as a management report reads them: cost of "
                  "sales, then operating costs with personnel first, then non-cash and "
                  "financing items. Bold rows are group subtotals. Every line is a cost, "
                  "so an overspend is unfavourable whichever timeframe it shows up in."
            ).font = F_NOTE

    widths(ws, L.widths(28))
    fit_text_columns(ws, ["A"], first, r)
    quiet_indicators(ws, 5, last + 2)
    ws.freeze_panes = f"B{first}"


# ===========================================================================
# By Entity
# ===========================================================================
def _write_by_entity(ws, ent_var, gl, glf, gll, bud, budf, budl, period) -> None:
    """Net income per legal entity, on the same columns as every other sheet.

    The measure is net income - revenue less spend - because that is the one
    figure that consolidates to the group P&L. The revenue and spend halves are
    on the P&L, where they are read against the rest of the structure.
    """
    hide_grid(ws)
    L = Layout(1)
    title_band(ws, "By Entity · net income by legal entity",
               f"Reporting month {period}  ·  with year to date  ·  \u20ac", L.last_col)
    headers(ws, 5, L.headers("Entity"), center_from=2, center_to=L.ncols)
    gR = lambda col: f"'{gl}'!${col}${glf}:${col}${gll}"
    bR = lambda col: f"'{bud}'!${col}${budf}:${col}${budl}"
    perno = int(period[:4]) * 100 + int(period[5:7])
    mflt = f',{gR(GL_PERIOD)},"{period}"'
    mfltb = f',{bR(BUD_PERIOD)},"{period}"'
    yflt = f',{gR(GL_PERNO)},"<={perno}"'
    yfltb = f',{bR(BUD_PERNO)},"<={perno}"'

    def net_gl(ent, flt):
        return (f'=SUMIFS({gR(GL_EUR)},{gR(GL_ENT)},"{ent}",{gR(GL_CAT)},"Revenue"{flt})'
                f'-SUMIFS({gR(GL_EUR)},{gR(GL_ENT)},"{ent}",{gR(GL_CAT)},"<>Revenue"{flt})')

    def net_bud(ent, flt):
        return (f'=SUMIFS({bR(BUD_BUD)},{bR(BUD_ENT)},"{ent}",{bR(BUD_CAT)},"Revenue"{flt})'
                f'-SUMIFS({bR(BUD_BUD)},{bR(BUD_ENT)},"{ent}",{bR(BUD_CAT)},"<>Revenue"{flt})')

    first, r, rows = 6, 6, []
    for i, (_, row) in enumerate(ent_var.iterrows()):
        ent = row["entity"]; band = band_fill(i)
        c = L.row(r)
        ws.cell(row=r, column=1, value=ent).font = F_BODY
        ws.cell(row=r, column=1).alignment = LEFT
        ws[c["act"]] = net_gl(ent, mflt)
        ws[c["bud"]] = net_bud(ent, mfltb)
        ws[c["yact"]] = net_gl(ent, yflt)
        ws[c["ybud"]] = net_bud(ent, yfltb)
        ws[c["fybud"]] = net_bud(ent, "")
        _tail(ws, L, r, higher=True)
        for col in L.span():
            ws[f"{col}{r}"].fill = band
        ws.row_dimensions[r].height = 22
        rows.append(r); r += 1
    last = r - 1

    ws.cell(row=r, column=1, value="Consolidated").font = F_SUB
    ws.cell(row=r, column=1).alignment = LEFT
    _sum_tail(ws, L, r, rows, higher=True)
    for col in L.span():
        ws[f"{col}{r}"].fill = FILL_IVORY
        ws[f"{col}{r}"].border = TOTAL_TOP
    ws.row_dimensions[r].height = 24
    rows.append(r)

    # Ranges run to `r`, the total row, not `last`: a total carries the same
    # verdict as the lines above it and was being left uncoloured.
    _report_cf(ws, L, first, r, rows, [])
    ws.cell(row=r + 2, column=1,
            value="Each entity's net income is its revenue less its spend, which is why "
                  "the entities consolidate to the group P&L. The revenue and spend "
                  "halves are on the P&L Report.").font = F_NOTE
    widths(ws, L.widths(18))
    fit_text_columns(ws, ["A"], first, r)
    quiet_indicators(ws, 5, last + 2)
    ws.freeze_panes = f"A{first}"


# ===========================================================================
# Cost Centres
# ===========================================================================
def _write_cost_centres(ws, dept_var, gl, glf, gll, bud, budf, budl, period) -> None:
    """Departmental spend variance with each department's cost centres nested."""
    hide_grid(ws)
    L = Layout(1)
    title_band(ws, "Departments & Cost Centres · spend variance",
               f"Reporting month {period}  ·  with year to date  ·  \u20ac", L.last_col)
    headers(ws, 5, L.headers("Department / Cost Centre"), center_from=2,
            center_to=L.ncols)

    gR = lambda col: f"'{gl}'!${col}${glf}:${col}${gll}"
    bR = lambda col: f"'{bud}'!${col}${budf}:${col}${budl}"
    perno = int(period[:4]) * 100 + int(period[5:7])
    mflt = f',{gR(GL_PERIOD)},"{period}"'
    mfltb = f',{bR(BUD_PERIOD)},"{period}"'
    yflt = f',{gR(GL_PERNO)},"<={perno}"'
    yfltb = f',{bR(BUD_PERNO)},"<={perno}"'

    order = dept_var.set_index("department")["budget"].to_dict()
    depts = sorted(SPEND_DEPARTMENTS, key=lambda d: order.get(d, 0), reverse=True)

    def spend_gl(col, key, flt):
        return f'=SUMIFS({gR(GL_EUR)},{gR(col)},"{key}",{gR(GL_CAT)},"<>Revenue"{flt})'

    def spend_bud(col, key, flt):
        return f'=SUMIFS({bR(BUD_BUD)},{bR(col)},"{key}",{bR(BUD_CAT)},"<>Revenue"{flt})'

    first, r, dept_rows, all_rows = 6, 6, [], []
    for dept in depts:
        c = L.row(r)
        ws.cell(row=r, column=1, value=dept).font = F_SUB
        ws.cell(row=r, column=1).alignment = LEFT
        ws[c["act"]] = spend_gl(GL_DEPT, dept, mflt)
        ws[c["bud"]] = spend_bud(BUD_DEPT, dept, mfltb)
        ws[c["yact"]] = spend_gl(GL_DEPT, dept, yflt)
        ws[c["ybud"]] = spend_bud(BUD_DEPT, dept, yfltb)
        ws[c["fybud"]] = spend_bud(BUD_DEPT, dept, "")
        _tail(ws, L, r, higher=False, bold=True)
        for col in L.span():
            ws[f"{col}{r}"].fill = FILL_IVORY
            ws[f"{col}{r}"].border = SUBTOTAL_TOP
        ws.row_dimensions[r].height = 22
        dept_rows.append(r); all_rows.append(r); r += 1

        for i, (cc_code, cc_name) in enumerate(DEPARTMENT_COST_CENTRES[dept]):
            c = L.row(r)
            label = ws.cell(row=r, column=1, value=f"{cc_code}   {cc_name}")
            label.font = F_BODY
            label.alignment = indent(2)
            ws[c["act"]] = spend_gl(GL_CC, cc_code, mflt)
            ws[c["bud"]] = spend_bud(BUD_CC, cc_code, mfltb)
            ws[c["yact"]] = spend_gl(GL_CC, cc_code, yflt)
            ws[c["ybud"]] = spend_bud(BUD_CC, cc_code, yfltb)
            ws[c["fybud"]] = spend_bud(BUD_CC, cc_code, "")
            _tail(ws, L, r, higher=False)
            band = band_fill(i)
            for col in L.span():
                ws[f"{col}{r}"].fill = band
            ws.row_dimensions[r].height = 18
            all_rows.append(r); r += 1
    last = r - 1

    # --- grand total: sum of the department rows only (no double counting) ---
    ws.cell(row=r, column=1, value="Total spend").font = F_SUB
    ws.cell(row=r, column=1).alignment = LEFT
    _sum_tail(ws, L, r, dept_rows, higher=False)
    for col in L.span():
        ws[f"{col}{r}"].fill = FILL_IVORY
        ws[f"{col}{r}"].border = TOTAL_TOP
    ws.row_dimensions[r].height = 24
    all_rows.append(r)

    _report_cf(ws, L, first, r, [], all_rows)
    ws.cell(row=r + 2, column=1,
            value="Departments are bold roll-ups; cost centres are indented beneath their "
                  "department, shown as code plus description. Totals sum the departments "
                  "only. Revenue is excluded: this is a spend view.").font = F_NOTE

    widths(ws, L.widths(32))
    fit_text_columns(ws, ["A"], first, r)
    quiet_indicators(ws, 5, last + 3)
    ws.freeze_panes = f"A{first}"


# ===========================================================================
# Drivers
# ===========================================================================
def _write_drivers(ws, agg, gl, glf, gll, bud, budf, budl, period) -> None:
    hide_grid(ws)
    L = Layout(3)
    title_band(ws, "Variance Drivers · account level",
               f"Reporting month {period}  ·  with year to date  ·  \u20ac", L.last_col)
    headers(ws, 5, L.headers("Account", "Name", "Category"),
            center_from=4, center_to=L.ncols)
    gR = lambda col: f"'{gl}'!${col}${glf}:${col}${gll}"
    bR = lambda col: f"'{bud}'!${col}${budf}:${col}${budl}"
    perno = int(period[:4]) * 100 + int(period[5:7])
    mflt = f',{gR(GL_PERIOD)},"{period}"'
    mfltb = f',{bR(BUD_PERIOD)},"{period}"'
    yflt = f',{gR(GL_PERNO)},"<={perno}"'
    yfltb = f',{bR(BUD_PERNO)},"<={perno}"'

    tmp = agg.copy(); tmp["_var"] = tmp["actual"] - tmp["budget"]
    tmp = tmp.sort_values("_var", key=lambda s: s.abs(), ascending=False).reset_index(drop=True)
    first, r = 6, 6
    rows_higher, rows_lower = [], []
    for i, (_, row) in enumerate(tmp.iterrows()):
        higher = CATEGORY_FAVOURABLE[row["category"]] == FAV_HIGHER
        code = row["account_code"]
        band = band_fill(i)
        c = L.row(r)
        ws.cell(row=r, column=1, value=code).font = F_BODY
        ws.cell(row=r, column=2, value=row["account_name"]).font = F_BODY
        ws.cell(row=r, column=3, value=row["category"]).font = F_BODY
        ws[c["act"]] = f'=SUMIFS({gR(GL_EUR)},{gR(GL_CODE)},"{code}"{mflt})'
        ws[c["bud"]] = f'=SUMIFS({bR(BUD_BUD)},{bR(BUD_CODE)},"{code}"{mfltb})'
        ws[c["yact"]] = f'=SUMIFS({gR(GL_EUR)},{gR(GL_CODE)},"{code}"{yflt})'
        ws[c["ybud"]] = f'=SUMIFS({bR(BUD_BUD)},{bR(BUD_CODE)},"{code}"{yfltb})'
        ws[c["fybud"]] = f'=SUMIF({bR(BUD_CODE)},"{code}",{bR(BUD_BUD)})'
        _tail(ws, L, r, higher=higher)
        for col in L.span():
            ws[f"{col}{r}"].fill = band
        for col in ("A", "B", "C"):
            ws[f"{col}{r}"].alignment = LEFT
        ws.row_dimensions[r].height = 18
        (rows_higher if higher else rows_lower).append(r)
        r += 1
    last = r - 1
    # No data bar on the variance column: the rows are already ordered by the
    # size of the movement, while a bar encodes the signed value, so the longest
    # bar was rarely the top row. The colour carries the direction, the F/U cell
    # carries the verdict and the flag carries the timeframe - a fourth encoding
    # on the same cell only competed with the digits behind it.
    _report_cf(ws, L, first, last, rows_higher, rows_lower)
    widths(ws, L.widths(12, 30, 12))
    fit_text_columns(ws, ["B", "C"], first, last)
    quiet_indicators(ws, 5, last)
    ws.freeze_panes = f"A{first}"

# ===========================================================================
def build_demo_pack(period: str, out_path: str | Path, seed: int = 42) -> Path:
    """Build the full multi-entity showcase pack from generated data."""
    from ..synthetic_data import (generate_ytd_transactions, generate_budget_year,
                                  monthly_detail, generate_month)
    from ..engine import department_variances, entity_variances

    txns = generate_ytd_transactions(period, seed)
    bud = generate_budget_year(seed).sort_values(
        ["period_no", "entity", "department", "cost_centre", "account_code"]
    ).reset_index(drop=True)
    detail = monthly_detail(period, seed)
    agg = generate_month(period, seed).drop(columns="period")

    report = build_report(agg)
    comments = line_comments(report, agg)
    dept_var = department_variances(detail)
    ent_var = entity_variances(detail)

    wb = Workbook()
    ws_gl = wb.active; ws_gl.title = "GL Transactions"
    glf, gll = _write_gl_transactions(ws_gl, txns, period)
    ws_bud = wb.create_sheet("Budget"); budf, budl = _write_budget(ws_bud, bud, period)

    GL, BUD = "GL Transactions", "Budget"
    _write_pnl(wb.create_sheet("P&L Report"), GL, glf, gll, BUD, budf, budl, period, comments, report)
    _write_expense_report(wb.create_sheet("Expense Report"), GL, glf, gll, BUD, budf, budl, period)
    _write_by_entity(wb.create_sheet("By Entity"), ent_var, GL, glf, gll, BUD, budf, budl, period)
    _write_cost_centres(wb.create_sheet("Departments & CCs"), dept_var, GL, glf, gll, BUD, budf, budl, period)
    _write_drivers(wb.create_sheet("Drivers"), agg, GL, glf, gll, BUD, budf, budl, period)

    desired = ["P&L Report", "Expense Report", "By Entity",
               "Departments & CCs", "Drivers", "Budget", "GL Transactions"]
    for i, name in enumerate(desired):
        wb.move_sheet(name, -wb.sheetnames.index(name) + i)
    wb.active = wb.sheetnames.index("P&L Report")

    out_path = Path(out_path); out_path.parent.mkdir(parents=True, exist_ok=True)
    quiet = collect_quiet_ranges(wb)
    wb.save(out_path)
    suppress_error_indicators(out_path, quiet)
    return out_path


if __name__ == "__main__":  # pragma: no cover - manual run
    root = Path(__file__).resolve().parents[3]
    out = build_demo_pack("2025-06", root / "output" / "flux_demo_pack.xlsx")
    print(f"Written: {out}")

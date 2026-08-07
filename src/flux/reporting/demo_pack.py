"""
Demo pack: the full multi-entity showcase, built from generated data.

Inputs (how a real finance function holds data):
  - GL Transactions : year-to-date transaction-level actuals, multi-entity,
                      multi-currency (with FX to EUR), tagged with cost centre,
                      expense type, posting date and source journal.
  - Budget          : full-year summarised plan (and prior year) by entity x
                      cost centre x account x period.

Every reporting sheet is a live formula view over those two inputs, filtered to
the reporting month (or year-to-date), and all views reconcile to the source:
  - P&L Report     consolidated P&L, KPI cards, per-line commentary
  - Expense Report natural view: expense type x (month / YTD / full-year budget)
  - By Entity      consolidation: net income per legal entity
  - Departments & CCs  functional view: department roll-up with cost centres
  - Drivers        account-level movers with data bars

Output: output/flux_demo_pack.xlsx
"""

from __future__ import annotations
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.formatting.rule import DataBarRule
from openpyxl.utils import get_column_letter

from ..coa import (PNL_STRUCTURE, CATEGORY_FAVOURABLE, FAV_HIGHER,
                   SPEND_DEPARTMENTS, DEPARTMENT_COST_CENTRES, EXPENSE_GROUPS)
from ..commentary import line_comments
from ..engine import build_report
from .styling import (
    FONT, BAR, GREEN_INK, RED_INK,
    F_HEAD, F_BODY, F_SMALL, F_SUB, F_INPUT, F_LABEL, F_KPI_LABEL, F_KPI_VALUE,
    F_NOTE, F_FU, F_FLAG,
    FILL_HEAD, FILL_IVORY, FILL_BAND, FILL_WHITE, FILL_LEVER,
    CUR, CUR2, CUR_EUR, CUR2_EUR, PCT, RATE, KPI_DELTA, LCY_FORMATS,
    CENTER, LEFT, RIGHT, WRAP, indent, band_fill,
    GOLD_SIDE, HEADER_BOTTOM, SUBTOTAL_TOP, TOTAL_TOP,
    hide_grid, title_band, headers, widths, outline, badge_cf, wrapped_height, fit_text_columns,
    variance_cf, spend_variance_cf,
    named_style,
    quiet_indicators, collect_quiet_ranges, suppress_error_indicators,
)
from .formulas import (
    LEVER_EUR, LEVER_PCT, LEV_E, LEV_P, PL_E, PL_P,
    DEFAULT_ABS_THRESHOLD, DEFAULT_PCT_THRESHOLD,
    pct_f, flag_f,
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
    hide_grid(ws)
    title_band(ws, "Management P&L · Variance Report (consolidated)", f"Reporting month {period}  ·  \u20ac", "K")
    gR = lambda col: f"'{gl}'!${col}${glf}:${col}${gll}"
    bR = lambda col: f"'{bud}'!${col}${budf}:${col}${budl}"

    ws.merge_cells("A9:B9"); ws["A9"] = "Materiality floor (\u20ac)"; ws["A9"].font = F_LABEL; ws["A9"].alignment = LEFT
    ws[LEVER_EUR] = DEFAULT_ABS_THRESHOLD; ws[LEVER_EUR].font = F_INPUT; ws[LEVER_EUR].number_format = CUR_EUR
    ws[LEVER_EUR].fill = FILL_LEVER; ws[LEVER_EUR].alignment = CENTER
    ws.merge_cells("E9:F9"); ws["E9"] = "Materiality floor (%)"; ws["E9"].font = F_LABEL; ws["E9"].alignment = LEFT
    ws[LEVER_PCT] = DEFAULT_PCT_THRESHOLD; ws[LEVER_PCT].font = F_INPUT; ws[LEVER_PCT].number_format = PCT
    ws[LEVER_PCT].fill = FILL_LEVER; ws[LEVER_PCT].alignment = CENTER
    outline(ws, 9, 3, 9, 3, GOLD_SIDE); outline(ws, 9, 7, 9, 7, GOLD_SIDE)

    headers(ws, 10, ["", "Actual", "Budget", "Var (Bud)", "Var %", "Prior Yr Act",
                      "Var (PY)", "Var %", "F/U", "Flag", "Commentary"], center_from=2, center_to=10)

    subtotal_labels = {"Gross profit", "Operating income (EBIT)", "Net income"}
    label_row, first, r, cat_idx = {}, 11, 11, 0
    for line in PNL_STRUCTURE:
        A, B, C, D, E = f"A{r}", f"B{r}", f"C{r}", f"D{r}", f"E{r}"
        Fc, G, H, I, J, K = f"F{r}", f"G{r}", f"H{r}", f"I{r}", f"J{r}", f"K{r}"
        is_sub = line.label in subtotal_labels
        ws[A] = line.label
        if line.kind == "category":
            ws[B] = f'=SUMIFS({gR(GL_EUR)},{gR(GL_CAT)},"{line.category}",{gR(GL_PERIOD)},"{period}")'
            ws[C] = f'=SUMIFS({bR(BUD_BUD)},{bR(BUD_CAT)},"{line.category}",{bR(BUD_PERIOD)},"{period}")'
            ws[Fc] = f'=SUMIFS({bR(BUD_PRIOR)},{bR(BUD_CAT)},"{line.category}",{bR(BUD_PERIOD)},"{period}")'
        else:
            pb, pc, pf = [], [], []
            for sign, ref in line.components:
                rr = label_row[ref]; pb.append(f"{sign}B{rr}"); pc.append(f"{sign}C{rr}"); pf.append(f"{sign}F{rr}")
            ws[B] = "=" + "".join(pb); ws[C] = "=" + "".join(pc); ws[Fc] = "=" + "".join(pf)
        ws[D] = f"={B}-{C}"; ws[E] = pct_f(D, C)
        ws[G] = f"={B}-{Fc}"; ws[H] = pct_f(G, Fc)
        ws[I] = f'=IF({D}>=0,"F","U")' if line.favourable == FAV_HIGHER else f'=IF({D}<=0,"F","U")'
        ws[J] = flag_f(D, E, LEV_E, LEV_P)
        base_fill = FILL_IVORY if is_sub else (FILL_BAND if cat_idx % 2 else FILL_WHITE)
        if not is_sub:
            cat_idx += 1
        for col in "ABCDEFGHIJK":
            ws[f"{col}{r}"].fill = base_fill
        for col in ("B", "C", "D", "F", "G"):
            cc = ws[f"{col}{r}"]; cc.number_format = CUR; cc.alignment = RIGHT; cc.font = F_SUB if is_sub else F_BODY
        for col in ("E", "H"):
            cc = ws[f"{col}{r}"]; cc.number_format = PCT; cc.alignment = RIGHT; cc.font = F_SUB if is_sub else F_BODY
        ws[A].font = F_SUB if is_sub else F_BODY; ws[A].alignment = LEFT
        ws[I].alignment = CENTER; ws[I].font = F_FU
        ws[J].alignment = CENTER; ws[J].font = F_FLAG
        comment = comments.get(line.label, "")
        kc = ws[K]; kc.value = comment; kc.alignment = WRAP
        kc.font = F_SUB if is_sub else F_BODY
        if is_sub:
            for col in "ABCDEFGHIJK":
                ws[f"{col}{r}"].border = SUBTOTAL_TOP
        ws.row_dimensions[r].height = wrapped_height(comment, COMMENT_WIDTH)
        label_row[line.label] = r; r += 1
    last = r - 1
    badge_cf(ws, f"I{first}:I{last}", f"J{first}:J{last}")
    variance_cf(ws, f"D{first}:E{last}", f"$I{first}")

    cards = [("REVENUE", "Revenue"), ("OPERATING INCOME (EBIT)", "Operating income (EBIT)"), ("NET INCOME", "Net income")]
    for (title, key), (c1, c2) in zip(cards, [(1, 3), (5, 7), (9, 11)]):
        lr = label_row[key]; fav = report[report["line"] == key].iloc[0]["fav_unfav"] == "F"
        L = get_column_letter(c1); Rt = get_column_letter(c2)
        for rr in (5, 6, 7):
            ws.merge_cells(f"{L}{rr}:{Rt}{rr}")
            for cc in range(c1, c2 + 1):
                ws.cell(row=rr, column=cc).fill = FILL_IVORY
        ws[f"{L}5"] = title; ws[f"{L}5"].font = F_KPI_LABEL
        ws[f"{L}5"].alignment = LEFT
        ws[f"{L}6"] = f"=B{lr}"; ws[f"{L}6"].font = F_KPI_VALUE; ws[f"{L}6"].number_format = CUR_EUR
        ws[f"{L}6"].alignment = LEFT
        ws[f"{L}7"] = f"=E{lr}"
        ws[f"{L}7"].number_format = KPI_DELTA
        ws[f"{L}7"].font = Font(name=FONT, size=9, bold=True, color=GREEN_INK if fav else RED_INK)
        ws[f"{L}7"].alignment = LEFT
        outline(ws, 5, c1, 7, c2, GOLD_SIDE)
    ws.row_dimensions[5].height = 18; ws.row_dimensions[6].height = 26; ws.row_dimensions[7].height = 16

    widths(ws, [26, 13, 13, 13, 9, 13, 13, 9, 6, 11, COMMENT_WIDTH])
    quiet_indicators(ws, 5, last)
    ws.freeze_panes = f"A{first}"


# ===========================================================================
# Expense Report (natural view: expense type x month/YTD/FY)
# ===========================================================================
def _write_expense_report(ws, gl, glf, gll, bud, budf, budl, period) -> None:
    hide_grid(ws)
    # The fiscal year comes from the period, not a constant: hardcoded, this
    # header still read "FY 2025" for a 2026 close.
    title_band(ws, "Expense Report · by expense type",
                f"Reporting month {period}  ·  YTD & FY {period[:4]}  ·  \u20ac", "L")
    gR = lambda col: f"'{gl}'!${col}${glf}:${col}${gll}"
    bR = lambda col: f"'{bud}'!${col}${budf}:${col}${budl}"
    perno = int(period[:4]) * 100 + int(period[5:7])

    headers(ws, 5, ["Expense Type", "Month Act", "Month Bud", "Month Var", "Var %",
                     "YTD Act", "YTD Bud", "YTD Var", "Var %", "FY Budget", "FY used", "Flag"],
             center_from=2, center_to=12)

    first, r, i = 6, 6, 0
    group_rows = []
    for group_name, types in EXPENSE_GROUPS:
        # A group with a single expense type needs no separate subtotal row.
        single = len(types) == 1
        g_first = r
        for et in types:
            B, C, D, E = f"B{r}", f"C{r}", f"D{r}", f"E{r}"
            F, G, H, I, J, K, L = (f"F{r}", f"G{r}", f"H{r}", f"I{r}",
                                   f"J{r}", f"K{r}", f"L{r}")
            band = FILL_IVORY if single else (band_fill(i))
            lab = ws.cell(row=r, column=1, value=et)
            lab.font = F_SUB if single else F_BODY
            lab.alignment = (LEFT if single else
                             indent(1))
            ws[B] = f'=SUMIFS({gR(GL_EUR)},{gR(GL_EXP)},"{et}",{gR(GL_PERIOD)},"{period}")'
            ws[C] = f'=SUMIFS({bR(BUD_BUD)},{bR(BUD_EXP)},"{et}",{bR(BUD_PERIOD)},"{period}")'
            ws[D] = f"={B}-{C}"; ws[E] = pct_f(D, C)
            ws[F] = (f'=SUMIFS({gR(GL_EUR)},{gR(GL_EXP)},"{et}",'
                     f'{gR(GL_PERNO)},"<={perno}")')
            ws[G] = (f'=SUMIFS({bR(BUD_BUD)},{bR(BUD_EXP)},"{et}",'
                     f'{bR(BUD_PERNO)},"<={perno}")')
            ws[H] = f"={F}-{G}"; ws[I] = pct_f(H, G)
            ws[J] = f'=SUMIF({bR(BUD_EXP)},"{et}",{bR(BUD_BUD)})'
            ws[K] = f'=IF({J}=0,"",{F}/{J})'
            ws[L] = flag_f(D, E, PL_E, PL_P)
            body = F_SUB if single else F_BODY
            for col in (B, C, D, F, G, H, J):
                ws[col].number_format = CUR; ws[col].font = body; ws[col].alignment = RIGHT
            for col in (E, I, K):
                ws[col].number_format = PCT; ws[col].font = body; ws[col].alignment = RIGHT
            ws[L].alignment = CENTER; ws[L].font = F_FLAG
            for c in range(1, 13):
                cell = ws.cell(row=r, column=c); cell.fill = band
                if single:
                    cell.border = SUBTOTAL_TOP
            ws.row_dimensions[r].height = 21 if single else 18
            i += 1; r += 1
        g_last = r - 1

        if single:
            group_rows.append(g_last)
            i = 0
            continue

        # group subtotal
        ws.cell(row=r, column=1, value=group_name).font = F_SUB
        ws.cell(row=r, column=1).alignment = LEFT
        for col in ("B", "C", "D", "F", "G", "H", "J"):
            ws[f"{col}{r}"] = f"=SUM({col}{g_first}:{col}{g_last})"
            ws[f"{col}{r}"].number_format = CUR; ws[f"{col}{r}"].font = F_SUB
            ws[f"{col}{r}"].alignment = RIGHT
        ws[f"E{r}"] = pct_f(f"D{r}", f"C{r}")
        ws[f"I{r}"] = pct_f(f"H{r}", f"G{r}")
        ws[f"K{r}"] = f'=IF(J{r}=0,"",F{r}/J{r})'
        for col in ("E", "I", "K"):
            ws[f"{col}{r}"].number_format = PCT; ws[f"{col}{r}"].font = F_SUB
            ws[f"{col}{r}"].alignment = RIGHT
        ws[f"L{r}"] = flag_f(f"D{r}", f"E{r}", PL_E, PL_P)
        ws[f"L{r}"].alignment = CENTER; ws[f"L{r}"].font = F_FLAG
        for c in range(1, 13):
            cell = ws.cell(row=r, column=c); cell.fill = FILL_IVORY; cell.border = SUBTOTAL_TOP
        ws.row_dimensions[r].height = 21
        group_rows.append(r)
        i = 0
        r += 1
    last = r - 1

    ws.cell(row=r, column=1, value="Total expenses").font = F_SUB
    for col in ("B", "C", "D", "F", "G", "H", "J"):
        ws[f"{col}{r}"] = "=" + "+".join(f"{col}{g}" for g in group_rows)
        ws[f"{col}{r}"].number_format = CUR; ws[f"{col}{r}"].font = F_SUB; ws[f"{col}{r}"].alignment = RIGHT
    ws[f"E{r}"] = pct_f(f"D{r}", f"C{r}")
    ws[f"I{r}"] = pct_f(f"H{r}", f"G{r}")
    ws[f"K{r}"] = f'=IF(J{r}=0,"",F{r}/J{r})'
    for col in ("E", "I", "K"):
        ws[f"{col}{r}"].number_format = PCT; ws[f"{col}{r}"].font = F_SUB; ws[f"{col}{r}"].alignment = RIGHT
    for col in "ABCDEFGHIJKL":
        ws[f"{col}{r}"].fill = FILL_IVORY; ws[f"{col}{r}"].border = SUBTOTAL_TOP
    ws.row_dimensions[r].height = 24

    badge_cf(ws, None, f"L{first}:L{r}")
    spend_variance_cf(ws, f"D{first}:E{r}")
    spend_variance_cf(ws, f"H{first}:I{r}")
    ws.cell(row=r + 2, column=1,
            value="Expense types are grouped as a management report reads them: cost of "
                  "sales, then operating costs with personnel first, then non-cash and "
                  "financing items. Bold rows are group subtotals.").font = F_NOTE

    widths(ws, [28, 13, 13, 12, 8, 13, 13, 12, 8, 14, 8, 11])
    fit_text_columns(ws, ["A"], first, r)
    quiet_indicators(ws, 5, last + 2)
    ws.freeze_panes = f"B{first}"


# ===========================================================================
# By Entity
# ===========================================================================
def _write_by_entity(ws, ent_var, gl, glf, gll, bud, budf, budl, period) -> None:
    hide_grid(ws)
    title_band(ws, "By Entity · net income by legal entity", f"Reporting month {period}  ·  \u20ac", "H")
    headers(ws, 5, ["Entity", "Revenue", "Spend", "Net income", "Budget net",
                     "Var (Bud)", "Var %", "F/U"], center_from=2, center_to=8)
    gR = lambda col: f"'{gl}'!${col}${glf}:${col}${gll}"
    bR = lambda col: f"'{bud}'!${col}${budf}:${col}${budl}"
    pf = f'{gR(GL_PERIOD)},"{period}"'; pfb = f'{bR(BUD_PERIOD)},"{period}"'

    first, r = 6, 6
    for i, (_, row) in enumerate(ent_var.iterrows()):
        ent = row["entity"]; band = band_fill(i)
        B, C, D, E, Fc, G, H = f"B{r}", f"C{r}", f"D{r}", f"E{r}", f"F{r}", f"G{r}", f"H{r}"
        ws.cell(row=r, column=1, value=ent).font = F_BODY
        ws[B] = f'=SUMIFS({gR(GL_EUR)},{gR(GL_ENT)},"{ent}",{gR(GL_CAT)},"Revenue",{pf})'
        ws[C] = f'=SUMIFS({gR(GL_EUR)},{gR(GL_ENT)},"{ent}",{gR(GL_CAT)},"<>Revenue",{pf})'
        ws[D] = f"={B}-{C}"
        ws[E] = (f'=SUMIFS({bR(BUD_BUD)},{bR(BUD_ENT)},"{ent}",{bR(BUD_CAT)},"Revenue",{pfb})'
                 f'-SUMIFS({bR(BUD_BUD)},{bR(BUD_ENT)},"{ent}",{bR(BUD_CAT)},"<>Revenue",{pfb})')
        ws[Fc] = f"={D}-{E}"; ws[G] = pct_f(Fc, E); ws[H] = f'=IF({Fc}>=0,"F","U")'
        for col in ("B", "C", "D", "E", "F"):
            ws[f"{col}{r}"].number_format = CUR; ws[f"{col}{r}"].font = F_BODY; ws[f"{col}{r}"].alignment = RIGHT
        ws[G].number_format = PCT; ws[G].font = F_BODY; ws[G].alignment = RIGHT
        ws[H].alignment = CENTER; ws[H].font = F_FU
        for c in range(1, 9):
            ws.cell(row=r, column=c).fill = band
        ws.cell(row=r, column=1).alignment = LEFT; ws.row_dimensions[r].height = 22
        r += 1
    last = r - 1
    ws.cell(row=r, column=1, value="Consolidated").font = F_SUB
    for col, lo in (("B", "B"), ("C", "C"), ("E", "E")):
        ws[f"{col}{r}"] = f"=SUM({col}{first}:{col}{last})"
    ws[f"D{r}"] = f"=B{r}-C{r}"; ws[f"F{r}"] = f"=D{r}-E{r}"
    ws[f"G{r}"] = f'=IF(E{r}=0,"",F{r}/ABS(E{r}))'; ws[f"H{r}"] = f'=IF(F{r}>=0,"F","U")'
    for col in ("B", "C", "D", "E", "F"):
        ws[f"{col}{r}"].number_format = CUR; ws[f"{col}{r}"].font = F_SUB; ws[f"{col}{r}"].alignment = RIGHT
    ws[f"G{r}"].number_format = PCT; ws[f"G{r}"].font = F_SUB; ws[f"G{r}"].alignment = RIGHT
    ws[f"H{r}"].alignment = CENTER; ws[f"H{r}"].font = F_SUB
    for col in "ABCDEFGH":
        ws[f"{col}{r}"].fill = FILL_IVORY; ws[f"{col}{r}"].border = SUBTOTAL_TOP
    ws.row_dimensions[r].height = 24
    badge_cf(ws, f"H{first}:H{last}", None)
    variance_cf(ws, f"F{first}:G{last}", f"$H{first}")
    widths(ws, [16, 15, 15, 15, 15, 14, 9, 6])
    quiet_indicators(ws, 5, last + 2)
    ws.freeze_panes = f"A{first}"


# ===========================================================================
# Cost Centres
# ===========================================================================
def _write_cost_centres(ws, dept_var, gl, glf, gll, bud, budf, budl, period) -> None:
    """Departmental spend variance with each department's cost centres nested."""
    hide_grid(ws)
    title_band(ws, "Departments & Cost Centres · spend variance", f"Reporting month {period}  ·  \u20ac", "G")
    headers(ws, 5, ["Department / Cost Centre", "Actual", "Budget", "Var (Bud)", "Var %",
                     "F/U", "Flag"], center_from=2, center_to=7)

    gR = lambda col: f"'{gl}'!${col}${glf}:${col}${gll}"
    bR = lambda col: f"'{bud}'!${col}${budf}:${col}${budl}"
    pf = f'{gR(GL_PERIOD)},"{period}"'
    pfb = f'{bR(BUD_PERIOD)},"{period}"'

    order = dept_var.set_index("department")["budget"].to_dict()
    depts = sorted(SPEND_DEPARTMENTS, key=lambda d: order.get(d, 0), reverse=True)

    first = 6
    r = first
    dept_rows = []
    for dept in depts:
        # --- department roll-up row ---
        B, C, D, E, Fc, G = f"B{r}", f"C{r}", f"D{r}", f"E{r}", f"F{r}", f"G{r}"
        ws.cell(row=r, column=1, value=dept).font = F_SUB
        ws.cell(row=r, column=1).alignment = LEFT
        ws[B] = f'=SUMIFS({gR(GL_EUR)},{gR(GL_DEPT)},"{dept}",{gR(GL_CAT)},"<>Revenue",{pf})'
        ws[C] = f'=SUMIFS({bR(BUD_BUD)},{bR(BUD_DEPT)},"{dept}",{bR(BUD_CAT)},"<>Revenue",{pfb})'
        ws[D] = f"={B}-{C}"
        ws[E] = pct_f(D, C)
        ws[Fc] = f'=IF({D}<=0,"F","U")'
        ws[G] = flag_f(D, E, PL_E, PL_P)
        for col in (B, C, D):
            ws[col].number_format = CUR; ws[col].font = F_SUB; ws[col].alignment = RIGHT
        ws[E].number_format = PCT; ws[E].font = F_SUB; ws[E].alignment = RIGHT
        ws[Fc].alignment = CENTER; ws[Fc].font = F_FU
        ws[G].alignment = CENTER; ws[G].font = F_FLAG
        for c in range(1, 8):
            ws.cell(row=r, column=c).fill = FILL_IVORY
            ws.cell(row=r, column=c).border = SUBTOTAL_TOP
        ws.row_dimensions[r].height = 22
        dept_rows.append(r)
        r += 1

        # --- cost centres under it ---
        for i, (cc_code, cc_name) in enumerate(DEPARTMENT_COST_CENTRES[dept]):
            B, C, D, E, Fc, G = f"B{r}", f"C{r}", f"D{r}", f"E{r}", f"F{r}", f"G{r}"
            label = ws.cell(row=r, column=1, value=f"{cc_code}   {cc_name}")
            label.font = F_BODY
            label.alignment = indent(2)
            ws[B] = f'=SUMIFS({gR(GL_EUR)},{gR(GL_CC)},"{cc_code}",{gR(GL_CAT)},"<>Revenue",{pf})'
            ws[C] = f'=SUMIFS({bR(BUD_BUD)},{bR(BUD_CC)},"{cc_code}",{bR(BUD_CAT)},"<>Revenue",{pfb})'
            ws[D] = f"={B}-{C}"
            ws[E] = pct_f(D, C)
            ws[Fc] = f'=IF({D}<=0,"F","U")'
            ws[G] = flag_f(D, E, PL_E, PL_P)
            band = band_fill(i)
            for col in (B, C, D):
                ws[col].number_format = CUR; ws[col].font = F_BODY; ws[col].alignment = RIGHT
            ws[E].number_format = PCT; ws[E].font = F_BODY; ws[E].alignment = RIGHT
            ws[Fc].alignment = CENTER; ws[Fc].font = F_FU
            ws[G].alignment = CENTER; ws[G].font = F_FLAG
            for c in range(1, 8):
                ws.cell(row=r, column=c).fill = band
            ws.row_dimensions[r].height = 18
            r += 1
    last = r - 1

    # --- grand total: sum of the department rows only (no double counting) ---
    ws.cell(row=r, column=1, value="Total spend").font = F_SUB
    for col in ("B", "C"):
        ws[f"{col}{r}"] = "=" + "+".join(f"{col}{dr}" for dr in dept_rows)
    ws[f"D{r}"] = f"=B{r}-C{r}"
    ws[f"E{r}"] = f'=IF(C{r}=0,"",D{r}/ABS(C{r}))'
    ws[f"F{r}"] = f'=IF(D{r}<=0,"F","U")'
    for col in ("B", "C", "D"):
        ws[f"{col}{r}"].number_format = CUR; ws[f"{col}{r}"].font = F_SUB; ws[f"{col}{r}"].alignment = RIGHT
    ws[f"E{r}"].number_format = PCT; ws[f"E{r}"].font = F_SUB; ws[f"E{r}"].alignment = RIGHT
    ws[f"F{r}"].alignment = CENTER; ws[f"F{r}"].font = F_SUB
    for col in "ABCDEFG":
        ws[f"{col}{r}"].fill = FILL_IVORY
        ws[f"{col}{r}"].border = TOTAL_TOP
    ws.row_dimensions[r].height = 24

    badge_cf(ws, f"F{first}:F{last}", f"G{first}:G{last}")
    variance_cf(ws, f"D{first}:E{r}", f"$F{first}")
    ws.cell(row=r + 2, column=1,
            value="Departments are bold roll-ups; cost centres are indented beneath their "
                  "department, shown as code plus description. Totals sum the departments "
                  "only. Revenue is excluded: this is a spend view.").font = F_NOTE

    widths(ws, [32, 15, 15, 15, 10, 6, 11])
    fit_text_columns(ws, ["A"], first, r)
    quiet_indicators(ws, 5, last + 3)
    ws.freeze_panes = f"A{first}"


# ===========================================================================
# Drivers
# ===========================================================================
def _write_drivers(ws, agg, gl, glf, gll, bud, budf, budl, period) -> None:
    hide_grid(ws)
    title_band(ws, "Variance Drivers · account level", f"Reporting month {period}  ·  \u20ac", "I")
    headers(ws, 5, ["Account", "Name", "Category", "Actual", "Budget", "Var (Bud)", "Var %", "F/U", "Flag"],
             center_from=4, center_to=9)
    gR = lambda col: f"'{gl}'!${col}${glf}:${col}${gll}"
    bR = lambda col: f"'{bud}'!${col}${budf}:${col}${budl}"
    pf = f'{gR(GL_PERIOD)},"{period}"'; pfb = f'{bR(BUD_PERIOD)},"{period}"'

    tmp = agg.copy(); tmp["_var"] = tmp["actual"] - tmp["budget"]
    tmp = tmp.sort_values("_var", key=lambda s: s.abs(), ascending=False).reset_index(drop=True)
    first, r = 6, 6
    for i, (_, row) in enumerate(tmp.iterrows()):
        favdir = CATEGORY_FAVOURABLE[row["category"]]; code = row["account_code"]
        band = band_fill(i)
        D, E, Fc, G, H, I = f"D{r}", f"E{r}", f"F{r}", f"G{r}", f"H{r}", f"I{r}"
        ws.cell(row=r, column=1, value=code).font = F_BODY
        ws.cell(row=r, column=2, value=row["account_name"]).font = F_BODY
        ws.cell(row=r, column=3, value=row["category"]).font = F_BODY
        ws[D] = f'=SUMIFS({gR(GL_EUR)},{gR(GL_CODE)},"{code}",{pf})'
        ws[E] = f'=SUMIFS({bR(BUD_BUD)},{bR(BUD_CODE)},"{code}",{pfb})'
        for col in (D, E):
            ws[col].number_format = CUR; ws[col].font = F_BODY; ws[col].alignment = RIGHT
        ws[Fc] = f"={D}-{E}"; ws[Fc].number_format = CUR; ws[Fc].font = F_BODY; ws[Fc].alignment = RIGHT
        ws[G] = pct_f(Fc, E); ws[G].number_format = PCT; ws[G].font = F_BODY; ws[G].alignment = RIGHT
        ws[H] = f'=IF({Fc}>=0,"F","U")' if favdir == FAV_HIGHER else f'=IF({Fc}<=0,"F","U")'
        ws[H].alignment = CENTER; ws[H].font = F_FU
        ws[I] = flag_f(Fc, G, PL_E, PL_P)
        ws[I].alignment = CENTER; ws[I].font = F_FLAG
        for c in range(1, 10):
            ws.cell(row=r, column=c).fill = band
            if c <= 3:
                ws.cell(row=r, column=c).alignment = LEFT
        ws.row_dimensions[r].height = 18
        r += 1
    last = r - 1
    ws.conditional_formatting.add(f"F{first}:F{last}",
        DataBarRule(start_type="min", end_type="max", color=BAR, showValue=True))
    badge_cf(ws, f"H{first}:H{last}", f"I{first}:I{last}")
    variance_cf(ws, f"F{first}:G{last}", f"$H{first}")
    widths(ws, [12, 30, 12, 15, 15, 15, 10, 6, 11])
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

    desired = ["P&L Report", "Expense Report", "By Entity", "Departments & CCs", "Drivers", "Budget", "GL Transactions"]
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

"""
Excel pack for ingested client data.

The sheets adapt to whichever dimensions the uploaded file actually carries:

    always            GL Input, P&L Report, Drivers
    + expense_type    Expense Report      (natural view)
    + department      Departments & CCs   (functional view, hierarchy taken
                                           from the data, not from a fixed
                                           chart of accounts)
    + entity          By Entity           (consolidation)

Every figure is a live formula over the GL Input sheet, so the pack recalculates
when an input is edited. Every reporting sheet carries the same columns - the
month, the year to date, the full-year plan and the run rate against it, then
F/U and the materiality flag - so the sheets differ only in what they cut the
ledger by. Styling, formula and row helpers are shared with the demo pack.
"""

from __future__ import annotations
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from ..coa import PNL_STRUCTURE, CATEGORY_FAVOURABLE, FAV_HIGHER, EXPENSE_GROUPS
from ..commentary import line_comments
from ..engine import build_report, has_budget as _detect_budget
from .styling import (
    FONT, GREEN_INK, RED_INK,
    F_BODY, F_SMALL, F_SUB, F_INPUT, F_KPI_LABEL, F_KPI_VALUE, F_NOTE, F_META,
    FILL_IVORY, FILL_BAND, FILL_WHITE,
    CUR_EUR, PCT, KPI_DELTA,
    LEFT, RIGHT, WRAP, indent, band_fill,
    GOLD_SIDE, SUBTOTAL_TOP, TOTAL_TOP,
    hide_grid, title_band, headers, widths, outline, lever, note,
    wrapped_height, fit_text_columns,
    quiet_indicators, collect_quiet_ranges, suppress_error_indicators,
)
# The columns every reporting sheet shares are written by `rows`, so a sheet
# here only supplies the five figures that depend on how it cuts the ledger.
from .rows import (BASE_KEYS, report_cf as _report_cf,
                   write_sum_tail as _sum_tail, write_tail as _tail)
from .formulas import (
    Layout,
    LEVER_EUR, LEVER_PCT, LEVER_MONTHS, LEV_E, LEV_P, LEV_M,
    DEFAULT_ABS_THRESHOLD, DEFAULT_PCT_THRESHOLD, meta_line,
    Variance, NO_BUDGET_NOTE, SINGLE_PERIOD_NOTE,
)


OPTIONAL_DIMS = [("expense_type", "Expense type", 22),
                 ("department", "Department", 16),
                 ("cost_centre", "Cost centre", 14),
                 ("entity", "Entity", 16)]


def _present(df: pd.DataFrame, col: str) -> bool:
    """True when a dimension exists and actually carries values."""
    if col not in df.columns:
        return False
    vals = df[col].astype(str).str.strip()
    return bool((vals != "").any() and (vals != "(multiple)").any())


# ---------------------------------------------------------------------------
# GL Input
# ---------------------------------------------------------------------------
def _gl_input(ws, agg, period, dims, budgeted=True):
    hide_grid(ws)
    spec = []
    if "period" in dims:
        spec += [("period", "Period", 11), ("period_no", "Per. key", 10)]
    spec += [("account_code", "Account", 12), ("account_name", "Account name", 30),
             ("category", "Category", 12)]
    spec += [(k, label, w) for k, label, w in OPTIONAL_DIMS if k in dims]
    spec += [("actual", "Actual", 15), ("budget", "Budget", 15),
             ("prior_year", "Prior Yr Act", 15)]
    letters = {k: get_column_letter(i) for i, (k, _l, _w) in enumerate(spec, start=1)}
    money = {"actual", "budget", "prior_year"}
    ints = {"period_no"}

    title_band(ws, "GL Input · account level", meta_line(period),
                   get_column_letter(len(spec)))
    headers(ws, 5, [l for _k, l, _w in spec],
                center_from=len(spec) - 2, center_to=len(spec))
    for i, (_k, _l, w) in enumerate(spec, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    first = 6
    for i, (_, r) in enumerate(agg.iterrows()):
        row = first + i
        band = band_fill(i)
        for c, (key, _l, _w) in enumerate(spec, start=1):
            v = r.get(key, "")
            if key in money:
                cell = ws.cell(row=row, column=c, value=float(v or 0))
                cell.number_format = CUR_EUR; cell.font = F_INPUT; cell.alignment = RIGHT
            elif key in ints:
                cell = ws.cell(row=row, column=c, value=int(v) if pd.notna(v) else None)
                cell.number_format = "0"; cell.font = F_SMALL; cell.alignment = RIGHT
            else:
                cell = ws.cell(row=row, column=c, value="" if pd.isna(v) else str(v))
                cell.font = F_BODY; cell.alignment = LEFT
            cell.fill = band
        ws.row_dimensions[row].height = 17
    last = first + len(agg) - 1
    if not budgeted:
        note(ws, last + 2, NO_BUDGET_NOTE)
    quiet_indicators(ws, 5, last)
    ws.freeze_panes = f"A{first}"
    return first, last, letters


# ---------------------------------------------------------------------------
# Filters shared by every reporting sheet
# ---------------------------------------------------------------------------
def _filters(R, GLC, perno):
    """The month and year-to-date criteria for the SUMIFS on every sheet.

    A file with no period column cannot be cut by time, so both criteria are
    empty and the month, the year to date and the full year describe the same
    postings. That is stated on the sheet rather than hidden: three identical
    columns with no explanation read like a bug.
    """
    if perno and "period_no" in GLC:
        return (f',{R(GLC["period_no"])},{perno}',
                f',{R(GLC["period_no"])},"<={perno}"')
    return "", ""


def _months_elapsed(period: str | None) -> int:
    """The default for the months lever, read off the reporting period."""
    try:
        return int(str(period)[5:7])
    except (TypeError, ValueError):
        return 1


# ---------------------------------------------------------------------------
# P&L
# ---------------------------------------------------------------------------
def _pnl(ws, gl, glf, gll, GLC, period, comments, report, var, perno=None,
         single_period=False):
    """The management P&L, and the sheet that owns the pack's lever cells.

    Prior-year actuals are not on this sheet: a front page has room for one
    secondary comparison, and against budget beats against last year. The
    prior-year column is still on the GL Input sheet, where it can be read
    beside the figures it belongs to.
    """
    hide_grid(ws)
    L = Layout(1)
    com = get_column_letter(L.ncols + 1)
    title_band(ws, "Management P&L · Variance Report",
               meta_line(period, single_period), com)
    R = lambda c: f"'{gl}'!${c}${glf}:${c}${gll}"
    cat, act, bud = R(GLC["category"]), R(GLC["actual"]), R(GLC["budget"])
    mflt, yflt = _filters(R, GLC, perno)

    # The two materiality floors are only levers when there is something to
    # flag, so with no budget they are left off rather than shown as controls
    # that do nothing. Months elapsed stays either way: the run rate is built
    # from actuals alone.
    if var:
        lever(ws, "A9:B9", "Materiality floor (\u20ac)", LEVER_EUR,
              DEFAULT_ABS_THRESHOLD, CUR_EUR)
        lever(ws, "E9:F9", "Materiality floor (%)", LEVER_PCT,
              DEFAULT_PCT_THRESHOLD, PCT)
    lever(ws, "I9:J9", "Months elapsed", LEVER_MONTHS,
          _months_elapsed(period), "0")
    ws.merge_cells(f"M9:{com}9")
    ws["M9"] = ("F/U judges the month variance; Flag names which timeframe "
                "clears both floors.")
    ws["M9"].font = F_NOTE; ws["M9"].alignment = LEFT
    # The KPI cards now run to row 8, so the lever row sits directly beneath
    # them. A little more height is what keeps it from reading as a fourth line
    # of the card.
    ws.row_dimensions[9].height = 20

    headers(ws, 10, L.headers("") + ["Commentary"], center_from=2,
            center_to=L.ncols)
    width_list = L.widths(26) + [52]

    subs = {"Gross profit", "Operating income (EBIT)", "Net income"}
    label_row, first, r, ci = {}, 11, 11, 0
    rows_higher, rows_lower = [], []
    for line in PNL_STRUCTURE:
        c = L.row(r)
        is_sub = line.label in subs
        higher = line.favourable == FAV_HIGHER
        ws[f"A{r}"] = line.label
        if line.kind == "category":
            crit = f'{cat},"{line.category}"'
            ws[c["act"]] = f'=SUMIFS({act},{crit}{mflt})'
            ws[c["bud"]] = var.cell(f'=SUMIFS({bud},{crit}{mflt})')
            ws[c["yact"]] = f'=SUMIFS({act},{crit}{yflt})'
            ws[c["ybud"]] = var.cell(f'=SUMIFS({bud},{crit}{yflt})')
            # No period criterion: the full-year plan is the whole input sheet.
            ws[c["fybud"]] = var.cell(f'=SUMIFS({bud},{crit})')
        else:
            parts = {key: [] for key in BASE_KEYS}
            for sign, ref in line.components:
                rr = label_row[ref]
                for key in parts:
                    parts[key].append(f"{sign}{getattr(L, key)}{rr}")
            for key, terms in parts.items():
                formula = "=" + "".join(terms)
                ws[c[key]] = formula if key in ("act", "yact") else var.cell(formula)
        _tail(ws, L, r, higher=higher, bold=is_sub, var=var,
              lev_e=LEV_E, lev_p=LEV_P, months=LEV_M)

        fill = FILL_IVORY if is_sub else (FILL_BAND if ci % 2 else FILL_WHITE)
        if not is_sub:
            ci += 1
        for col in L.span() + [com]:
            ws[f"{col}{r}"].fill = fill
        ws[f"A{r}"].font = F_SUB if is_sub else F_BODY
        ws[f"A{r}"].alignment = LEFT
        comment = comments.get(line.label, "")
        kc = ws[f"{com}{r}"]; kc.value = comment; kc.alignment = WRAP
        kc.font = F_SUB if is_sub else F_BODY
        if is_sub:
            for col in L.span() + [com]:
                ws[f"{col}{r}"].border = SUBTOTAL_TOP
        (rows_higher if higher else rows_lower).append(r)
        ws.row_dimensions[r].height = wrapped_height(comment, width_list[-1])
        label_row[line.label] = r; r += 1
    last = r - 1
    _report_cf(ws, L, first, last, rows_higher, rows_lower, var=var)

    # The big number is the month - the sheet is titled by it, F/U judges it and
    # the commentary describes it - with the year to date beneath as context, so
    # the card answers "and how is the year going" without the reader scanning
    # across to the YTD block.
    for (title, key), (c1, c2) in zip(
        [("REVENUE", "Revenue"), ("OPERATING INCOME (EBIT)", "Operating income (EBIT)"),
         ("NET INCOME", "Net income")],
        [(1, 3), (5, 7), (9, 11)]):
        lr = label_row[key]; fav = report[report["line"] == key].iloc[0]["fav_unfav"] == "F"
        Lc = get_column_letter(c1); Rt = get_column_letter(c2)
        for rr in (5, 6, 7, 8):
            ws.merge_cells(f"{Lc}{rr}:{Rt}{rr}")
            for cc in range(c1, c2 + 1):
                ws.cell(row=rr, column=cc).fill = FILL_IVORY
        ws[f"{Lc}5"] = title; ws[f"{Lc}5"].font = F_KPI_LABEL
        ws[f"{Lc}5"].alignment = LEFT
        ws[f"{Lc}6"] = f"={L.act}{lr}"; ws[f"{Lc}6"].font = F_KPI_VALUE
        ws[f"{Lc}6"].number_format = CUR_EUR
        ws[f"{Lc}6"].alignment = LEFT
        ws[f"{Lc}7"] = var.cell(f"={L.pct}{lr}")
        ws[f"{Lc}7"].number_format = KPI_DELTA
        ws[f"{Lc}7"].font = Font(name=FONT, size=9, bold=True,
                                 color=GREEN_INK if fav else RED_INK)
        ws[f"{Lc}7"].alignment = LEFT
        # Muted, and built as text: the YTD line is context for the month above
        # it, so it must not compete with the green or red delta.
        ws[f"{Lc}8"] = (f'="YTD  "&TEXT({L.yact}{lr},"#,##0")&" \u20ac"'
                        f'&IF(ISNUMBER({L.ypct}{lr}),"  \u00b7  "'
                        f'&TEXT({L.ypct}{lr},"+0.0%;-0.0%")&" vs budget","")')
        ws[f"{Lc}8"].font = F_META; ws[f"{Lc}8"].alignment = LEFT
        outline(ws, 5, c1, 8, c2, GOLD_SIDE)
    ws.row_dimensions[5].height = 18; ws.row_dimensions[6].height = 26
    ws.row_dimensions[7].height = 14; ws.row_dimensions[8].height = 14
    widths(ws, width_list)

    footnote = last + 2
    if not var:
        note(ws, footnote, NO_BUDGET_NOTE); footnote += 1
    if single_period:
        note(ws, footnote, SINGLE_PERIOD_NOTE)
    quiet_indicators(ws, 5, last + 3)
    ws.freeze_panes = f"A{first}"


# ---------------------------------------------------------------------------
# By Entity
# ---------------------------------------------------------------------------
def _by_entity(ws, values, gl, glf, gll, GLC, period, var, perno=None,
               single_period=False):
    """Net income per legal entity, on the columns every other sheet uses.

    Net income - revenue less spend - is the one measure that consolidates to
    the group P&L, so it is the measure this sheet reports. The revenue and
    spend halves are on the P&L, read against the rest of the structure.
    """
    hide_grid(ws)
    L = Layout(1)
    title_band(ws, "By Entity · consolidation", meta_line(period, single_period), L.last_col)
    headers(ws, 5, L.headers("Entity"), center_from=2, center_to=L.ncols)
    R = lambda c: f"'{gl}'!${c}${glf}:${c}${gll}"
    dim, cat = R(GLC["entity"]), R(GLC["category"])
    act, bud = R(GLC["actual"]), R(GLC["budget"])
    mflt, yflt = _filters(R, GLC, perno)

    def net(source, ent, flt):
        return (f'=SUMIFS({source},{dim},"{ent}",{cat},"Revenue"{flt})'
                f'-SUMIFS({source},{dim},"{ent}",{cat},"<>Revenue"{flt})')

    first, r, rows = 6, 6, []
    for i, ent in enumerate(values):
        c = L.row(r)
        ws.cell(row=r, column=1, value=ent).font = F_BODY
        ws.cell(row=r, column=1).alignment = LEFT
        ws[c["act"]] = net(act, ent, mflt)
        ws[c["bud"]] = var.cell(net(bud, ent, mflt))
        ws[c["yact"]] = net(act, ent, yflt)
        ws[c["ybud"]] = var.cell(net(bud, ent, yflt))
        ws[c["fybud"]] = var.cell(net(bud, ent, ""))
        _tail(ws, L, r, higher=True, var=var)
        band = band_fill(i)
        for col in L.span():
            ws[f"{col}{r}"].fill = band
        ws.row_dimensions[r].height = 19
        rows.append(r); r += 1
    last = r - 1

    ws.cell(row=r, column=1, value="Total").font = F_SUB
    ws.cell(row=r, column=1).alignment = LEFT
    _sum_tail(ws, L, r, rows, higher=True, var=var)
    for col in L.span():
        ws[f"{col}{r}"].fill = FILL_IVORY
        ws[f"{col}{r}"].border = TOTAL_TOP
    ws.row_dimensions[r].height = 24
    rows.append(r)

    _report_cf(ws, L, first, r, rows, [], var=var)
    ws.cell(row=r + 2, column=1,
            value="Each entity's net income is its revenue less its spend, which is why "
                  "the entities consolidate to the group P&L.").font = F_NOTE
    widths(ws, L.widths(26))
    fit_text_columns(ws, ["A"], first, r)
    quiet_indicators(ws, 5, last + 2)
    ws.freeze_panes = f"A{first}"


def _group_expense_types(values: list[str]) -> list[tuple[str, list[str]]]:
    """Arrange the expense types present into readable groups.

    Known types follow the standard order (personnel first, then other operating
    costs, cost of sales, non-cash and financing). Anything a client calls by its
    own name still gets reported, collected under "Other expense types" so no
    line is silently dropped.
    """
    remaining = list(values)
    groups: list[tuple[str, list[str]]] = []
    for name, types in EXPENSE_GROUPS:
        present = [t for t in types if t in remaining]
        if present:
            groups.append((name, present))
            remaining = [t for t in remaining if t not in present]
    if remaining:
        groups.append(("Other expense types", sorted(remaining)))
    return groups


def _expense_sheet(ws, groups, gl, glf, gll, GLC, period, var, perno=None,
                   single_period=False):
    """Expense report by type, grouped with subtotals."""
    hide_grid(ws)
    L = Layout(1)
    title_band(ws, "Expense Report · by expense type", meta_line(period, single_period), L.last_col)
    R = lambda c: f"'{gl}'!${c}${glf}:${c}${gll}"
    dim, cat = R(GLC["expense_type"]), R(GLC["category"])
    act, bud = R(GLC["actual"]), R(GLC["budget"])
    mflt, yflt = _filters(R, GLC, perno)

    headers(ws, 5, L.headers("Expense type"), center_from=2, center_to=L.ncols)

    def write_line(r, et):
        c = L.row(r)
        crit = f'{cat},"<>Revenue",{dim},"{et}"'
        ws[c["act"]] = f'=SUMIFS({act},{crit}{mflt})'
        ws[c["bud"]] = var.cell(f'=SUMIFS({bud},{crit}{mflt})')
        ws[c["yact"]] = f'=SUMIFS({act},{crit}{yflt})'
        ws[c["ybud"]] = var.cell(f'=SUMIFS({bud},{crit}{yflt})')
        ws[c["fybud"]] = var.cell(f'=SUMIFS({bud},{crit})')

    def dress(r, label, *, bold, fill, border=False, level=0):
        cell = ws.cell(row=r, column=1, value=label)
        cell.font = F_SUB if bold else F_BODY
        cell.alignment = indent(level) if level else LEFT
        for col in L.span():
            cc = ws[f"{col}{r}"]; cc.fill = fill
            if border:
                cc.border = SUBTOTAL_TOP
        ws.row_dimensions[r].height = 21 if bold else 18

    first, r, group_rows, all_rows = 6, 6, [], []
    for gname, types in groups:
        # A group holding one type needs no separate subtotal line.
        single = len(types) == 1
        type_rows = []
        for i, et in enumerate(types):
            write_line(r, et)
            _tail(ws, L, r, higher=False, bold=single, var=var)
            dress(r, et, bold=single, fill=FILL_IVORY if single else band_fill(i),
                  border=single, level=0 if single else 1)
            type_rows.append(r); all_rows.append(r); r += 1
        if single:
            group_rows.append(type_rows[0])
            continue
        _sum_tail(ws, L, r, type_rows, higher=False, var=var)
        dress(r, gname, bold=True, fill=FILL_IVORY, border=True)
        group_rows.append(r); all_rows.append(r); r += 1
    last = r - 1

    _sum_tail(ws, L, r, group_rows, higher=False, var=var)
    dress(r, "Total expenses", bold=True, fill=FILL_IVORY)
    for col in L.span():
        ws[f"{col}{r}"].border = TOTAL_TOP
    ws.row_dimensions[r].height = 24
    all_rows.append(r)

    _report_cf(ws, L, first, r, [], all_rows, var=var)
    ws.cell(row=r + 2, column=1,
            value="Grouped as a cost owner reads them: personnel first, then other "
                  "operating costs, cost of sales, and non-cash and financing items. "
                  "Bold rows are group subtotals. Every line is a cost, so an overspend "
                  "is unfavourable whichever timeframe it shows up in.").font = F_NOTE
    widths(ws, L.widths(30))
    fit_text_columns(ws, ["A"], first, r)
    quiet_indicators(ws, 5, last + 3)
    ws.freeze_panes = f"B{first}"


# ---------------------------------------------------------------------------
# Departments & cost centres (hierarchy read from the data)
# ---------------------------------------------------------------------------
def _departments(ws, hierarchy, gl, glf, gll, GLC, period, has_cc, var,
                 perno=None, single_period=False):
    hide_grid(ws)
    L = Layout(1)
    title_band(ws, "Departments & Cost Centres · spend variance", meta_line(period, single_period), L.last_col)
    headers(ws, 5, L.headers("Department / Cost centre"), center_from=2,
            center_to=L.ncols)
    R = lambda c: f"'{gl}'!${c}${glf}:${c}${gll}"
    dep, cat = R(GLC["department"]), R(GLC["category"])
    act, bud = R(GLC["actual"]), R(GLC["budget"])
    cc = R(GLC["cost_centre"]) if has_cc else None
    mflt, yflt = _filters(R, GLC, perno)

    def write_line(r, dim, key):
        c = L.row(r)
        crit = f'{dim},"{key}",{cat},"<>Revenue"'
        ws[c["act"]] = f'=SUMIFS({act},{crit}{mflt})'
        ws[c["bud"]] = var.cell(f'=SUMIFS({bud},{crit}{mflt})')
        ws[c["yact"]] = f'=SUMIFS({act},{crit}{yflt})'
        ws[c["ybud"]] = var.cell(f'=SUMIFS({bud},{crit}{yflt})')
        ws[c["fybud"]] = var.cell(f'=SUMIFS({bud},{crit})')

    first, r, dept_rows, all_rows = 6, 6, [], []
    for dept, ccs in hierarchy.items():
        ws.cell(row=r, column=1, value=dept).font = F_SUB
        ws.cell(row=r, column=1).alignment = LEFT
        write_line(r, dep, dept)
        _tail(ws, L, r, higher=False, bold=True, var=var)
        for col in L.span():
            ws[f"{col}{r}"].fill = FILL_IVORY
            ws[f"{col}{r}"].border = SUBTOTAL_TOP
        ws.row_dimensions[r].height = 22
        dept_rows.append(r); all_rows.append(r); r += 1

        if has_cc:
            for i, c_name in enumerate(ccs):
                lab = ws.cell(row=r, column=1, value=str(c_name))
                lab.font = F_BODY
                lab.alignment = indent(2)
                write_line(r, cc, c_name)
                _tail(ws, L, r, higher=False, var=var)
                band = band_fill(i)
                for col in L.span():
                    ws[f"{col}{r}"].fill = band
                ws.row_dimensions[r].height = 18
                all_rows.append(r); r += 1
    last = r - 1

    ws.cell(row=r, column=1, value="Total spend").font = F_SUB
    ws.cell(row=r, column=1).alignment = LEFT
    _sum_tail(ws, L, r, dept_rows, higher=False, var=var)
    for col in L.span():
        ws[f"{col}{r}"].fill = FILL_IVORY
        ws[f"{col}{r}"].border = TOTAL_TOP
    ws.row_dimensions[r].height = 24
    all_rows.append(r)

    _report_cf(ws, L, first, r, [], all_rows, var=var)
    ws.cell(row=r + 2, column=1,
            value="Departments are roll-ups; cost centres are indented beneath them. "
                  "Totals sum the departments only. Revenue is excluded: this is a "
                  "spend view.").font = F_NOTE
    widths(ws, L.widths(32))
    fit_text_columns(ws, ["A"], first, r)
    quiet_indicators(ws, 5, last + 3)
    ws.freeze_panes = f"A{first}"


# ---------------------------------------------------------------------------
# Drivers
# ---------------------------------------------------------------------------
def _drivers(ws, agg, gl, glf, gll, GLC, period, var, perno=None,
             single_period=False):
    hide_grid(ws)
    L = Layout(3)
    title_band(ws, "Variance Drivers · account level", meta_line(period, single_period), L.last_col)
    headers(ws, 5, L.headers("Account", "Account name", "Category"),
            center_from=4, center_to=L.ncols)
    R = lambda c: f"'{gl}'!${c}${glf}:${c}${gll}"
    code_r, act, bud = R(GLC["account_code"]), R(GLC["actual"]), R(GLC["budget"])
    mflt, yflt = _filters(R, GLC, perno)
    keep = agg
    if perno and "period_no" in agg.columns:
        keep = agg[agg["period_no"] == perno]
    tmp = (keep.groupby(["account_code", "account_name", "category"], as_index=False)
               [["actual", "budget"]].sum())
    tmp["_var"] = tmp["actual"] - tmp["budget"]
    tmp = tmp.sort_values("_var", key=lambda s: s.abs(), ascending=False)

    first, r = 6, 6
    rows_higher, rows_lower = [], []
    for i, (_, row) in enumerate(tmp.iterrows()):
        higher = CATEGORY_FAVOURABLE.get(row["category"], FAV_HIGHER) == FAV_HIGHER
        code = str(row["account_code"])
        c = L.row(r)
        ws.cell(row=r, column=1, value=code).font = F_BODY
        ws.cell(row=r, column=2, value=str(row["account_name"])).font = F_BODY
        ws.cell(row=r, column=3, value=str(row["category"])).font = F_BODY
        ws[c["act"]] = f'=SUMIFS({act},{code_r},"{code}"{mflt})'
        ws[c["bud"]] = var.cell(f'=SUMIFS({bud},{code_r},"{code}"{mflt})')
        ws[c["yact"]] = f'=SUMIFS({act},{code_r},"{code}"{yflt})'
        ws[c["ybud"]] = var.cell(f'=SUMIFS({bud},{code_r},"{code}"{yflt})')
        ws[c["fybud"]] = var.cell(f'=SUMIFS({bud},{code_r},"{code}")')
        _tail(ws, L, r, higher=higher, var=var)
        band = band_fill(i)
        for col in L.span():
            ws[f"{col}{r}"].fill = band
        for col in ("A", "B", "C"):
            ws[f"{col}{r}"].alignment = LEFT
        ws.row_dimensions[r].height = 18
        (rows_higher if higher else rows_lower).append(r)
        r += 1
    last = r - 1
    # No data bar on the variance column: see the note in demo_pack. The rows
    # are ordered by the size of the movement and a bar encodes the signed
    # value, so the two disagreed on every sheet that carried one.
    _report_cf(ws, L, first, last, rows_higher, rows_lower, var=var)
    widths(ws, L.widths(12, 30, 12))
    fit_text_columns(ws, ["B", "C"], first, last)
    quiet_indicators(ws, 5, last)
    ws.freeze_panes = f"A{first}"


# ---------------------------------------------------------------------------
def build_client_pack(agg: pd.DataFrame, period: str | None = None,
                      out_path: str | Path = "flux_pack.xlsx",
                      budgeted: bool | None = None) -> Path:
    """Build the pack, adapting to the dimensions and periods in the data.

    The reporting month is the latest period present; when the file spans more
    than one period the pack also shows year-to-date beside it, which is how a
    management report is read.

    With no budget in the data the variance, F/U and materiality columns are
    left empty rather than measured against zero, and the pack says so.
    """
    from .. import ingest

    agg = agg.copy()
    for col in ("actual", "budget", "prior_year"):
        if col not in agg.columns:
            agg[col] = 0.0
        agg[col] = agg[col].fillna(0.0)
    for col in ("expense_type", "department", "cost_centre", "entity"):
        if col in agg.columns:
            agg[col] = agg[col].fillna("").astype(str).str.strip()

    # Reporting period follows from the data.
    has_period = "period" in agg.columns and agg["period"].astype(str).str.strip().ne("").any()
    all_periods = []
    perno = None
    if has_period:
        if "period_no" not in agg.columns:
            agg["period_no"] = agg["period"].map(ingest.period_key)
        latest, all_periods = ingest.reporting_period(agg)
        period = period or latest
        perno = ingest.period_key(period)
    period = period or "current period"
    # Every sheet carries the month, the year to date and the full year whatever
    # the file holds. When the file holds one period the three describe the same
    # postings, which the pack states rather than leaves the reader to notice.
    single_period = not (has_period and len(all_periods) > 1 and perno is not None)

    dims = [k for k, _l, _w in OPTIONAL_DIMS if _present(agg, k)]
    gl_dims = (["period"] if has_period else []) + dims

    # The ledger reads chronologically, then by account.
    sort_cols = (["period_no"] if "period_no" in agg.columns else []) + ["account_code"]
    agg = agg.sort_values(sort_cols, kind="stable").reset_index(drop=True)

    # Figures behind the commentary and KPI colouring use the reporting month.
    month = agg
    if perno is not None and "period_no" in agg.columns:
        month = agg[agg["period_no"] == perno]
        if month.empty:
            month = agg
    budgeted = _detect_budget(agg) if budgeted is None else bool(budgeted)
    var = Variance(budgeted)

    report = build_report(month, budgeted=budgeted)
    comments = line_comments(report, month)

    wb = Workbook()
    ws_gl = wb.active; ws_gl.title = "GL Input"
    glf, gll, GLC = _gl_input(ws_gl, agg, period, gl_dims, budgeted)
    _pnl(wb.create_sheet("P&L Report"), "GL Input", glf, gll, GLC, period,
         comments, report, var, perno, single_period)

    order = ["P&L Report"]
    if "expense_type" in dims:
        vals = [v for v in month.loc[month["category"] != "Revenue", "expense_type"].unique()
                if v and v != "(multiple)"]
        if vals:
            _expense_sheet(wb.create_sheet("Expense Report"), _group_expense_types(vals),
                           "GL Input", glf, gll, GLC, period, var, perno,
                           single_period)
            order.append("Expense Report")
    if "entity" in dims:
        vals = [v for v in month["entity"].unique() if v and v != "(multiple)"]
        if len(vals) > 1:
            _by_entity(wb.create_sheet("By Entity"), sorted(vals), "GL Input",
                       glf, gll, GLC, period, var, perno, single_period)
            order.append("By Entity")
    if "department" in dims:
        spend = month[month["category"] != "Revenue"]
        has_cc = "cost_centre" in dims
        hierarchy = {}
        for dept in sorted(v for v in spend["department"].unique() if v and v != "(multiple)"):
            ccs = (sorted(v for v in spend.loc[spend["department"] == dept, "cost_centre"].unique()
                          if v and v != "(multiple)") if has_cc else [])
            hierarchy[dept] = ccs
        if hierarchy:
            _departments(wb.create_sheet("Departments & CCs"), hierarchy, "GL Input",
                         glf, gll, GLC, period, has_cc and any(hierarchy.values()),
                         var, perno, single_period)
            order.append("Departments & CCs")

    _drivers(wb.create_sheet("Drivers"), agg, "GL Input", glf, gll, GLC, period,
             var, perno, single_period)
    order += ["Drivers", "GL Input"]
    for i, name in enumerate(order):
        wb.move_sheet(name, -wb.sheetnames.index(name) + i)
    wb.active = wb.sheetnames.index("P&L Report")

    out_path = Path(out_path); out_path.parent.mkdir(parents=True, exist_ok=True)
    quiet = collect_quiet_ranges(wb)
    wb.save(out_path)
    suppress_error_indicators(out_path, quiet)
    return out_path


if __name__ == "__main__":  # pragma: no cover - manual run
    from .. import ingest
    root = Path(__file__).resolve().parents[3]
    std, _report, _issues = ingest.ingest(root / "data" / "input_template.xlsx")
    out = build_client_pack(ingest.aggregate_to_accounts(std),
                            out_path=root / "output" / "flux_client_pack.xlsx")
    print(f"Written: {out}")

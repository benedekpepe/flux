"""
PowerPoint pack: the deck a controller sends upward after close.

The Excel pack is the working file - every figure a live formula over the input
sheet. This is the other half of the same job: five slides that state the
result, name what moved it, and stop. Same engine, same numbers, same palette.

Built with python-pptx rather than a Node deck library so the capability ships
inside the Python package with no extra runtime.

Run:  PYTHONPATH=src python -m flux.reporting.pptx_pack
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import (XL_CHART_TYPE, XL_LABEL_POSITION,
                             XL_LEGEND_POSITION, XL_TICK_LABEL_POSITION)
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

from ..commentary import generate_commentary, line_comments
from ..engine import (MaterialityRule, build_report, department_variances,
                      entity_variances, has_budget, leaf_variances)


# ---------------------------------------------------------------------------
# Palette - the same colours the Excel packs are rendered in.
# ---------------------------------------------------------------------------
NAVY = RGBColor(0x1B, 0x24, 0x38)
DEEP = RGBColor(0x13, 0x1A, 0x29)
BRASS = RGBColor(0xC6, 0xA1, 0x5B)
IVORY = RGBColor(0xF5, 0xF1, 0xE8)
BAND = RGBColor(0xFA, 0xF8, 0xF2)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
INK = RGBColor(0x23, 0x27, 0x2E)
MUTE = RGBColor(0x8A, 0x8F, 0x98)
PALE = RGBColor(0xD7, 0xDC, 0xE6)
SLATE = RGBColor(0x8A, 0xA0, 0xCC)
GREEN = RGBColor(0x2E, 0x6B, 0x3E)
RED = RGBColor(0xA6, 0x3A, 0x3A)
AMBER = RGBColor(0x8A, 0x5A, 0x00)
AMBER_FILL = RGBColor(0xFB, 0xEE, 0xDA)

# Calibri renders true to width in preview and ships with Office, so the
# slides measure the same here as they will on the reader's machine.
FONT = "Calibri"

W, H = Inches(13.333), Inches(7.5)
MARGIN = Inches(0.62)


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
def _mag(v: float) -> str:
    """Money at the scale a slide is read from across a room."""
    a = abs(v)
    if a >= 1_000_000:
        return f"{v / 1_000_000:,.1f}m"
    if a >= 1_000:
        return f"{v / 1_000:,.0f}k"
    return f"{v:,.0f}"


def _signed(v: float) -> str:
    return ("+" if v >= 0 else "-") + _mag(abs(v)).lstrip("-")


def _pct(p) -> str:
    if p is None or (isinstance(p, float) and pd.isna(p)):
        return "n/m"
    return f"{p * 100:+.1f}%"


def _blank(prs: Presentation):
    """A slide with nothing on it: every element here is placed explicitly."""
    return prs.slides.add_slide(prs.slide_layouts[6])


def _fill(shape, colour, line=None, width=Pt(1)):
    shape.fill.solid()
    shape.fill.fore_color.rgb = colour
    if line is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = line
        shape.line.width = width
    shape.shadow.inherit = False


def _text(slide, x, y, w, h, runs, *, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
          spacing=None):
    """Place a text box. `runs` is a list of (text, size, bold, colour) tuples;
    a None entry starts a new paragraph."""
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    para = tf.paragraphs[0]
    para.alignment = align
    for item in runs:
        if item is None:
            para = tf.add_paragraph()
            para.alignment = align
            if spacing is not None:
                para.space_before = spacing
            continue
        txt, size, bold, colour = item
        run = para.add_run()
        run.text = txt
        run.font.name = FONT
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = colour
    return box


def _mark(slide, x, y, scale=1.0, plan=SLATE, actual=BRASS):
    """The Flux mark, drawn as shapes.

    Three rectangles: the plan bar, the shorter actual bar, and the rule showing
    where plan sat. Native shapes rather than an image, so it stays sharp at any
    size and the package needs no renderer to ship a logo.
    """
    u = Inches(0.028) * scale                       # the mark's grid unit
    def r(gx, gy, gw, gh, colour):
        s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                   x + Emu(int(u * gx)), y + Emu(int(u * gy)),
                                   Emu(int(u * gw)), Emu(int(u * gh)))
        s.adjustments[0] = 0.18
        _fill(s, colour)
        return s
    r(0, 0, 15, 30, plan)
    r(19, 12, 15, 18, actual)
    r(19, 0, 15, 4, actual)


def _slide_header(slide, eyebrow, title, meta=""):
    """Light content slide: small brass eyebrow, then the headline."""
    _text(slide, MARGIN, Inches(0.46), Inches(9.0), Inches(0.3),
          [(eyebrow.upper(), 11, True, BRASS)])
    _text(slide, MARGIN, Inches(0.74), Inches(9.6), Inches(0.6),
          [(title, 30, True, NAVY)])
    if meta:
        # Below the mark, not beside it: at this size the mark is only 0.4" wide
        # and the two collided on every content slide.
        _text(slide, W - MARGIN - Inches(4.2), Inches(1.02), Inches(4.2), Inches(0.3),
              [(meta, 12, False, MUTE)], align=PP_ALIGN.RIGHT)
    _mark(slide, W - MARGIN - Inches(0.42), Inches(0.46), scale=0.36,
          plan=RGBColor(0xC8, 0xD0, 0xE0), actual=BRASS)


def _dark_bg(slide):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, W, H)
    _fill(bg, DEEP)
    return bg


# ---------------------------------------------------------------------------
# Slide 1 - cover
# ---------------------------------------------------------------------------
def _cover(prs, entity, period, ni, budgeted):
    s = _blank(prs)
    _dark_bg(s)
    _mark(s, MARGIN, Inches(1.62), scale=1.45)

    word_x = MARGIN + Inches(1.85)
    _text(s, word_x, Inches(1.60), Inches(7.5), Inches(0.9),
          [("FLUX", 52, True, BRASS)])
    _text(s, word_x, Inches(2.44), Inches(9.0), Inches(0.4),
          [("Management reporting pack", 19, False, PALE)])

    _text(s, MARGIN, Inches(3.75), Inches(8.0), Inches(0.5),
          [(entity, 30, True, WHITE)])
    _text(s, MARGIN, Inches(4.35), Inches(8.0), Inches(0.4),
          [(f"Reporting month {period}  ·  currency EUR", 15, False, MUTE)])

    # The headline number belongs on the cover: it is what the reader opened for.
    card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                              W - MARGIN - Inches(4.0), Inches(3.55),
                              Inches(4.0), Inches(1.55))
    card.adjustments[0] = 0.09
    _fill(card, NAVY, line=BRASS, width=Pt(1.25))
    _text(s, W - MARGIN - Inches(3.7), Inches(3.78), Inches(3.4), Inches(0.28),
          [("NET INCOME", 11, True, BRASS)])
    _text(s, W - MARGIN - Inches(3.7), Inches(4.06), Inches(3.4), Inches(0.55),
          [(f"{ni['actual']:,.0f} EUR", 30, True, WHITE)])
    sub = (f"{_signed(ni['var_bud'])} ({_pct(ni['var_bud_pct'])}) vs budget"
           if budgeted else "no budget supplied")
    _text(s, W - MARGIN - Inches(3.7), Inches(4.66), Inches(3.4), Inches(0.3),
          [(sub, 12, False, RED if budgeted and ni["var_bud"] < 0 else PALE)])

    _text(s, MARGIN, H - Inches(0.95), Inches(11.0), Inches(0.3),
          [("Generated by Flux from the general ledger. Every figure ties to the "
            "Excel pack issued with this deck.", 11, False, MUTE)])
    return s


# ---------------------------------------------------------------------------
# Slide 2 - the result
# ---------------------------------------------------------------------------
def _kpi_card(s, x, y, w, label, value, delta, favourable, budgeted):
    card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, Inches(1.5))
    card.adjustments[0] = 0.09
    _fill(card, IVORY, line=BRASS, width=Pt(1))
    _text(s, x + Inches(0.26), y + Inches(0.2), w - Inches(0.5), Inches(0.25),
          [(label.upper(), 11, True, MUTE)])
    _text(s, x + Inches(0.26), y + Inches(0.48), w - Inches(0.5), Inches(0.5),
          [(f"{value:,.0f}", 28, True, NAVY)])
    if budgeted:
        _text(s, x + Inches(0.26), y + Inches(1.06), w - Inches(0.5), Inches(0.28),
              [(delta, 12, True, GREEN if favourable else RED)])


def _result(prs, report, gl, period, budgeted):
    s = _blank(prs)
    _slide_header(s, "The result", "Where the month landed", f"Reporting month {period}")

    def line(label):
        return report[report.line == label].iloc[0]

    rev, ebit, ni = line("Revenue"), line("Operating income (EBIT)"), line("Net income")
    gap = Inches(0.34)
    cw = (W - 2 * MARGIN - 2 * gap) / 3
    for i, r in enumerate((rev, ebit, ni)):
        _kpi_card(s, MARGIN + i * (cw + gap), Inches(1.62), cw, r.line, r.actual,
                  f"{_signed(r.var_bud)} ({_pct(r.var_bud_pct)}) vs budget"
                  if budgeted else "", r.fav_unfav == "F", budgeted)

    # The narrative's first two paragraphs: the bottom line, then revenue and margin.
    text = generate_commentary(report, gl).split("\n\n")
    body = s.shapes.add_textbox(MARGIN, Inches(3.45), W - 2 * MARGIN, Inches(3.4))
    tf = body.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = 0
    for i, para_text in enumerate(text):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(12)
        run = p.add_run()
        run.text = para_text
        run.font.name = FONT
        run.font.size = Pt(15)
        run.font.color.rgb = INK
    s.notes_slide.notes_text_frame.text = (
        "Figures are the reporting month against budget. The commentary is "
        "generated from the account-level variances, not written by hand."
    )
    return s


# ---------------------------------------------------------------------------
# Slide 3 - the P&L
# ---------------------------------------------------------------------------
def _pnl_table(prs, report, comments, period, budgeted):
    s = _blank(prs)
    _slide_header(s, "Profit and loss", "The management P&L", f"Reporting month {period}  ·  EUR")

    heads = (["Line", "Actual", "Budget", "Variance", "Var %", "F/U", ""]
             if budgeted else ["Line", "Actual", "Prior year", "Var (PY)", "", "", ""])
    rows = len(report) + 1
    # The money columns hold at most nine characters, the commentary a sentence,
    # so width goes where the text actually is.
    widths = [Inches(2.45), Inches(1.25), Inches(1.25), Inches(1.3),
              Inches(0.95), Inches(0.55), Inches(4.34)]

    tbl_shape = s.shapes.add_table(rows, len(heads), MARGIN, Inches(1.66),
                                   sum(widths, Inches(0)), Inches(0.48) * rows)
    tbl = tbl_shape.table
    tbl.first_row = False
    for i, wd in enumerate(widths):
        tbl.columns[i].width = wd
    tbl.rows[0].height = Inches(0.4)

    def cell(r, c, txt, *, size=12, bold=False, colour=INK, align=PP_ALIGN.RIGHT,
             fill=None):
        cl = tbl.cell(r, c)
        cl.text = ""
        cl.margin_left = cl.margin_right = Inches(0.09)
        cl.margin_top = cl.margin_bottom = Inches(0.02)
        cl.vertical_anchor = MSO_ANCHOR.MIDDLE
        if fill is not None:
            cl.fill.solid(); cl.fill.fore_color.rgb = fill
        else:
            cl.fill.background()
        p = cl.text_frame.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = txt
        run.font.name = FONT
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = colour

    for c, h in enumerate(heads):
        cell(0, c, h, size=11, bold=True, colour=WHITE, fill=NAVY,
             align=PP_ALIGN.LEFT if c == 0 else PP_ALIGN.RIGHT)

    subs = {"Gross profit", "Operating income (EBIT)", "Net income"}
    for i, r in enumerate(report.itertuples(), start=1):
        sub = r.line in subs
        bg = IVORY if sub else (BAND if i % 2 else WHITE)
        tbl.rows[i].height = Inches(0.48)
        cell(i, 0, r.line, bold=sub, align=PP_ALIGN.LEFT, fill=bg)
        cell(i, 1, f"{r.actual:,.0f}", bold=sub, fill=bg)
        if budgeted:
            cell(i, 2, f"{r.budget:,.0f}", bold=sub, fill=bg)
            cell(i, 3, f"{r.var_bud:,.0f}", bold=sub, fill=bg,
                 colour=GREEN if r.fav_unfav == "F" else RED)
            cell(i, 4, _pct(r.var_bud_pct), bold=sub, fill=bg)
            cell(i, 5, r.fav_unfav, bold=True, align=PP_ALIGN.CENTER, fill=bg,
                 colour=GREEN if r.fav_unfav == "F" else RED)
        else:
            cell(i, 2, f"{r.prior_year:,.0f}", bold=sub, fill=bg)
            cell(i, 3, f"{r.var_py:,.0f}", bold=sub, fill=bg)
            cell(i, 4, "", fill=bg); cell(i, 5, "", fill=bg)
        note = comments.get(r.line, "")
        cell(i, 6, note if len(note) <= 132 else note[:129].rstrip(" ,;") + "...",
             size=10, colour=MUTE, align=PP_ALIGN.LEFT, fill=bg)

    if budgeted:
        _text(s, MARGIN, H - Inches(0.82), Inches(11.5), Inches(0.3),
              [("Favourable / unfavourable follows the account type: revenue and "
                "profit lines are higher-is-better, cost lines lower-is-better.",
                10, False, MUTE)])
    return s


# ---------------------------------------------------------------------------
# Slide 4 - what moved it
# ---------------------------------------------------------------------------
def _drivers(prs, gl, materiality, period, budgeted):
    s = _blank(prs)
    _slide_header(s, "Drivers", "What moved the result", f"Reporting month {period}  ·  EUR")

    leaves = leaf_variances(gl, materiality)
    material = leaves[leaves["material"]].copy()
    if material.empty:
        _text(s, MARGIN, Inches(2.4), Inches(9.0), Inches(0.6),
              [("No account cleared both materiality floors this month.",
                18, False, INK)])
        return s

    material["mag"] = material["var_bud"].abs()
    material = material.sort_values("mag", ascending=False).head(8)
    # Charted bottom-up so the largest mover sits at the top of the bar chart.
    plot = material.iloc[::-1]

    data = CategoryChartData()
    data.categories = [n if len(n) < 34 else n[:31] + "..." for n in plot["account_name"]]
    data.add_series("Variance vs budget", tuple(plot["var_bud"]))

    gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, MARGIN, Inches(1.62),
                            Inches(8.1), Inches(4.6), data)
    chart = gf.chart
    chart.has_legend = False
    chart.has_title = False
    plot_area = chart.plots[0]
    plot_area.gap_width = 55
    plot_area.has_data_labels = True
    labels = plot_area.data_labels
    labels.number_format = '#,##0;(#,##0)'
    labels.number_format_is_linked = False
    labels.font.size = Pt(10)
    labels.font.name = FONT
    # Inside the bar, in white: outside the end, the longest bar's figure ran
    # straight into its own category name on the left.
    labels.font.color.rgb = WHITE
    labels.font.bold = True
    labels.position = XL_LABEL_POSITION.INSIDE_END

    # One colour per bar: a cost overrun and a revenue shortfall are both bad
    # news but have opposite signs, so sign alone cannot carry the meaning.
    for idx, fu in enumerate(plot["fav_unfav"]):
        point = chart.series[0].points[idx]
        point.format.fill.solid()
        point.format.fill.fore_color.rgb = GREEN if fu == "F" else RED

    for axis in (chart.category_axis, chart.value_axis):
        axis.tick_labels.font.size = Pt(11)
        axis.tick_labels.font.name = FONT
        axis.tick_labels.font.color.rgb = INK
        axis.has_major_gridlines = False
    # Account names belong at the left edge. Left at the axis they sit on top of
    # the bars, and a shortfall runs left of zero, straight through its own label.
    chart.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
    chart.value_axis.tick_labels.number_format = '#,##0;(#,##0)'
    chart.value_axis.tick_labels.number_format_is_linked = False
    chart.value_axis.has_major_gridlines = True
    chart.value_axis.major_gridlines.format.line.color.rgb = RGBColor(0xE4, 0xE7, 0xEC)
    chart.value_axis.major_gridlines.format.line.width = Pt(0.75)

    # The side panel says what the chart cannot: why these lines and no others.
    panel = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                               MARGIN + Inches(8.45), Inches(1.62),
                               Inches(3.2), Inches(4.6))
    panel.adjustments[0] = 0.05
    _fill(panel, IVORY, line=BRASS, width=Pt(1))

    unfav = material[material.fav_unfav == "U"]["mag"].sum()
    fav = material[material.fav_unfav == "F"]["mag"].sum()
    runs = [("MATERIALITY", 11, True, BRASS), None,
            (f"{materiality.abs_threshold:,.0f} EUR and "
             f"{materiality.pct_threshold:.0%}", 17, True, NAVY), None,
            ("An account is flagged only when it clears both floors, so a large "
             "percentage on a small base does not crowd out the real movers.",
             11, False, INK), None,
            (f"{len(material)} account(s) qualified", 12, True, NAVY), None,
            (f"Unfavourable  {_mag(unfav)}", 12, False, RED), None,
            (f"Favourable  {_mag(fav)}" if fav else "Favourable  none",
             12, False, GREEN if fav else MUTE), None,
            ("LARGEST SINGLE MOVER", 11, True, BRASS), None,
            (str(material.iloc[0]["account_name"]), 13, True, NAVY), None,
            (f"{_signed(material.iloc[0]['var_bud'])}  "
             f"({_pct(material.iloc[0]['var_bud_pct'])} vs budget)", 12, False,
             GREEN if material.iloc[0]["fav_unfav"] == "F" else RED)]
    _text(s, MARGIN + Inches(8.7), Inches(1.88), Inches(2.7), Inches(4.1), runs,
          spacing=Pt(9))
    return s


# ---------------------------------------------------------------------------
# Slide 5 - where the money went
# ---------------------------------------------------------------------------
def _spend(prs, detail, period, materiality):
    s = _blank(prs)
    _slide_header(s, "Spend", "Where the money went", f"Reporting month {period}  ·  EUR")

    dept = department_variances(detail, materiality).sort_values("actual", ascending=False)
    plot = dept.iloc[::-1]

    data = CategoryChartData()
    data.categories = list(plot["department"])
    # Budget first so the clustered bars put actual on top, where the eye lands.
    data.add_series("Budget", tuple(plot["budget"]))
    data.add_series("Actual", tuple(plot["actual"]))

    gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, MARGIN, Inches(1.62),
                            Inches(7.4), Inches(4.5), data)
    chart = gf.chart
    chart.has_title = False
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.legend.font.size = Pt(11)
    chart.legend.font.name = FONT
    chart.plots[0].gap_width = 60
    for series, colour in zip(chart.series, (SLATE, BRASS)):
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = colour
    for axis in (chart.category_axis, chart.value_axis):
        axis.tick_labels.font.size = Pt(11)
        axis.tick_labels.font.name = FONT
        axis.tick_labels.font.color.rgb = INK
    chart.category_axis.has_major_gridlines = False
    chart.value_axis.has_major_gridlines = True
    chart.value_axis.major_gridlines.format.line.color.rgb = RGBColor(0xE4, 0xE7, 0xEC)
    chart.value_axis.major_gridlines.format.line.width = Pt(0.75)
    chart.value_axis.tick_labels.number_format = '#,##0'
    chart.value_axis.tick_labels.number_format_is_linked = False

    # Entity consolidation beside it: the same spend, cut the other way.
    ent = entity_variances(detail)
    x0 = MARGIN + Inches(7.75)
    _text(s, x0, Inches(1.66), Inches(3.9), Inches(0.3),
          [("BY LEGAL ENTITY", 11, True, BRASS)])
    y = Inches(2.06)
    for r in ent.itertuples():
        card = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x0, y,
                                  Inches(3.9), Inches(1.02))
        card.adjustments[0] = 0.12
        _fill(card, BAND, line=RGBColor(0xE4, 0xE7, 0xEC), width=Pt(1))
        _text(s, x0 + Inches(0.22), y + Inches(0.14), Inches(2.2), Inches(0.28),
              [(r.entity, 14, True, NAVY)])
        _text(s, x0 + Inches(0.22), y + Inches(0.48), Inches(2.4), Inches(0.4),
              [(f"Net {r.net_actual:,.0f}", 13, False, INK)])
        fu_colour = GREEN if r.fav_unfav == "F" else RED
        _text(s, x0 + Inches(2.5), y + Inches(0.36), Inches(1.2), Inches(0.4),
              [(_signed(r.var_bud), 16, True, fu_colour)], align=PP_ALIGN.RIGHT)
        y = y + Inches(1.18)

    _text(s, MARGIN, H - Inches(0.82), Inches(11.5), Inches(0.3),
          [("Departmental figures are spend only: revenue is excluded, so the "
            "totals tie to the cost lines of the P&L, not to net income.",
            10, False, MUTE)])
    return s


# ---------------------------------------------------------------------------
def build_pptx_pack(gl: pd.DataFrame,
                    period: str,
                    out_path: str | Path,
                    detail: pd.DataFrame | None = None,
                    entity: str = "Demo Company Ltd",
                    materiality: MaterialityRule | None = None,
                    budgeted: bool | None = None) -> Path:
    """Build the management deck.

    Parameters
    ----------
    gl : account-level frame, the same shape the Excel pack is built from.
    period : reporting period label, e.g. "2025-06".
    detail : optional transaction-level frame carrying department and entity.
        Without it the spend slide is omitted rather than faked.
    budgeted : whether a plan exists; detected from the data when omitted.
    """
    materiality = materiality or MaterialityRule()
    budgeted = has_budget(gl) if budgeted is None else bool(budgeted)
    report = build_report(gl, materiality, budgeted=budgeted)
    comments = line_comments(report, gl, materiality)
    ni = report[report.line == "Net income"].iloc[0]

    prs = Presentation()
    prs.slide_width, prs.slide_height = W, H

    _cover(prs, entity, period, ni, budgeted)
    _result(prs, report, gl, period, budgeted)
    _pnl_table(prs, report, comments, period, budgeted)
    if budgeted:
        _drivers(prs, gl, materiality, period, budgeted)
    if detail is not None and {"department", "entity"} <= set(detail.columns):
        _spend(prs, detail, period, materiality)

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(out_path)
    return out_path


if __name__ == "__main__":  # pragma: no cover - manual run
    from ..synthetic_data import generate_month, monthly_detail
    root = Path(__file__).resolve().parents[3]
    period = "2025-06"
    out = build_pptx_pack(generate_month(period).drop(columns="period"), period,
                          root / "output" / "flux_management_pack.pptx",
                          detail=monthly_detail(period))
    print(f"Written: {out}")

"""
Flux - FP&A reporting app.

Upload a GL / trial-balance file, confirm the column mapping, and download a
management pack. Or use built-in sample data to see the full multi-entity
showcase without uploading anything.

Run:  streamlit run app.py
"""

import sys
import tempfile
import time
from datetime import date
from pathlib import Path

import pandas as pd
import streamlit as st

# The package lives under src/ so it can be installed, but Streamlit Cloud runs
# this file directly from the repo root, so make src/ importable either way.
sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from flux import ingest
from flux.ingest import (match_columns, _content_detect, apply_mapping,
                         CANONICAL, REQUIRED)
from flux.coa import EXPENSE_TYPES
from flux.engine import build_report, has_budget
from flux.reporting import build_client_pack, build_demo_pack, build_pptx_pack
from flux.synthetic_data import (generate_budget_year, generate_month,
                                 generate_ytd_transactions, monthly_detail)


ASSETS = Path(__file__).resolve().parent / "assets"

st.set_page_config(
    page_title="Flux · FP&A Reporting",
    page_icon=str(ASSETS / "flux-mark.svg"),
    layout="wide",
)

# The lockup sits in the app chrome; the mark alone is used where space is tight
# (collapsed sidebar, browser tab). Both are drawn in tones that hold up on a
# white and on a navy ground, because Streamlit reports the active theme but
# documents it as unreliable on first load.
st.logo(str(ASSETS / "flux-logo.svg"), icon_image=str(ASSETS / "flux-mark.svg"),
        size="large")

# No page heading: the lockup in the chrome already carries the name, and
# repeating it in the body set it twice in two different styles.
st.subheader("Automated management reporting")
st.caption("Turn a general ledger into a management P&L with variance analysis, "
           "favourable / unfavourable classification and materiality flagging.")

SAMPLE_PERIOD = "2025-06"
def _demo_blocks():
    """The sample company's findings, shared by its workbook and its deck.

    Imported here rather than at module scope: the analysis slide is one extra
    on one tab, and a missing helper used to stop the whole app from starting.
    A deck without the slide is a far better failure than a blank page.
    """
    try:
        from flux.reporting import demo_analysis_blocks
    except ImportError:
        return None
    return demo_analysis_blocks(SAMPLE_PERIOD)


def _demo_ytd_detail():
    """The cumulative detail behind the deck's spend slide, same source again."""
    try:
        from flux.reporting import demo_ytd_detail
    except ImportError:
        return None
    return demo_ytd_detail(SAMPLE_PERIOD)


XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PPTX_MIME = ("application/vnd.openxmlformats-officedocument"
             ".presentationml.presentation")

st.sidebar.markdown("### How it works")
st.sidebar.caption(
    "Flux reads the reporting month from your data: the latest period that "
    "carries postings. A file spanning several months also gets a year-to-date "
    "view."
)
st.sidebar.markdown("---")
st.sidebar.caption(
    "Upload your own GL/trial balance, or use the built-in sample company "
    "(multi-entity, multi-currency) to see the full pack."
)


@st.cache_data(show_spinner=False)
def _sample_ytd() -> pd.DataFrame:
    """The demo ledger and plan rolled up to the reporting month.

    Cached: it is the same frame on every click, and rolling five thousand
    postings up again would add a second to a button the user is watching.
    """
    cut = ingest.period_key(SAMPLE_PERIOD)
    txn = generate_ytd_transactions(SAMPLE_PERIOD)
    bud = generate_budget_year()
    actual = (txn[txn.period_no <= cut]
              .groupby(["account_code", "account_name", "category"], as_index=False)
              ["amount_eur"].sum().rename(columns={"amount_eur": "actual"}))
    plan = (bud[bud.period_no <= cut]
            .groupby(["account_code", "account_name", "category"], as_index=False)
            [["budget_eur", "prior_eur"]].sum()
            .rename(columns={"budget_eur": "budget", "prior_eur": "prior_year"}))
    return actual.merge(plan, on=["account_code", "account_name", "category"],
                        how="outer").fillna(0.0)


@st.cache_data(show_spinner=False)
def _sample_fy_budget() -> pd.DataFrame:
    """The demo company's full-year plan, for the outlook slide."""
    fy = (generate_budget_year()
          .groupby(["account_code", "account_name", "category"], as_index=False)
          [["budget_eur", "prior_eur"]].sum()
          .rename(columns={"budget_eur": "budget", "prior_eur": "prior_year"}))
    fy["actual"] = 0.0
    return fy


def _pnl_preview(std: pd.DataFrame, budgeted: bool = True):
    rep = build_report(std, budgeted=budgeted)
    ni = rep[rep.line == "Net income"].iloc[0]
    ebit = rep[rep.line == "Operating income (EBIT)"].iloc[0]
    rev = rep[rep.line == "Revenue"].iloc[0]

    def _delta(x):
        if not budgeted:
            return None
        p = x.var_bud_pct
        return "n/m vs budget" if p is None or pd.isna(p) else f"{p*100:+.1f}% vs budget"

    c1, c2, c3 = st.columns(3)
    c1.metric("Revenue", f"{rev.actual:,.0f} €", _delta(rev))
    c2.metric("EBIT", f"{ebit.actual:,.0f} €", _delta(ebit))
    c3.metric("Net income", f"{ni.actual:,.0f} €", _delta(ni))

    cols = (["line", "actual", "budget", "var_bud", "var_bud_pct", "fav_unfav", "material"]
            if budgeted else ["line", "actual", "prior_year"])
    labels = (["Line", "Actual", "Budget", "Var (Bud)", "Var %", "F/U", "Material"]
              if budgeted else ["Line", "Actual", "Prior Yr Act"])
    show = rep[cols].copy()
    show.columns = labels
    st.dataframe(show, width="stretch", hide_index=True)


def _build(label: str, done: str, key: str, fn, out: Path):
    """Run a pack build with visible progress and an honest summary.

    The sample workbook takes a few seconds to write. Without this the button
    click looks like nothing happened, and the natural response is to click it
    again. The finished label reports the real size and elapsed time rather than
    a generic tick, so the wait is accounted for.
    """
    with st.status(label, expanded=False) as status:
        started = time.perf_counter()
        try:
            fn(out)
        except Exception as exc:  # surfaced, not swallowed into a dead button
            status.update(label=f"Could not build the {done.lower()}", state="error")
            st.error(f"{type(exc).__name__}: {exc}")
            return
        size_kb = out.stat().st_size / 1024
        elapsed = time.perf_counter() - started
        size = f"{size_kb / 1024:,.1f} MB" if size_kb >= 1024 else f"{size_kb:,.0f} KB"
        status.update(label=f"{done} ready — {size} in {elapsed:.1f}s",
                      state="complete")
    _offer_download(out, "", key)


def _offer_download(path: Path, label: str, key: str):
    """Hand the finished file to the browser.

    The bytes are held in session state rather than rebuilt inside the button
    block: clicking a download button triggers a Streamlit rerun, which would
    otherwise collapse the block that created it and make the button vanish.
    """
    st.session_state[key] = {"name": path.name, "data": path.read_bytes()}


def _render_download(key: str, label: str, mime: str = XLSX_MIME):
    payload = st.session_state.get(key)
    if not payload:
        return
    st.download_button(label, payload["data"], file_name=payload["name"],
                       mime=mime, width="stretch", key=f"dl_{key}")


# Sample data leads. A visitor arriving on the upload tab sees an empty form and
# nothing the tool can do; arriving on the sample tab they see the KPIs, the P&L
# and a pack they can generate in one click, with the upload tab right beside it.
tab_sample, tab_upload = st.tabs(["Use sample data", "Upload your data"])

# ---------------------------------------------------------------------------
with tab_upload:
    st.markdown("#### 1 · Upload your data")
    col_a, col_b = st.columns([3, 1])
    with col_a:
        # Both uploaders share this column. Left to the full page width, the
        # budget one rendered wider than the actuals one and the mismatch read
        # as a mistake rather than a hierarchy.
        up = st.file_uploader(
            "Actuals — GL extract or trial balance (Flux will read it and show "
            "you every assumption it made)", type=["csv", "xlsx", "xls"],
            accept_multiple_files=False, key="up_actuals",
        )
        up_bud = st.file_uploader(
            "Budget — optional second file, if the plan lives outside the ledger",
            type=["csv", "xlsx", "xls"], accept_multiple_files=False, key="up_budget",
        )
    with col_b:
        # A label of its own, rather than an invisible spacer: the left column
        # starts with the uploader's label, so without one the button sat half a
        # line high. It also gives the button a name it was missing.
        st.markdown("Input template · recommended")
        tpath = Path(__file__).parent / "data" / "input_template.xlsx"
        if tpath.exists():
            st.download_button(
                "Download template", tpath.read_bytes(),
                file_name="flux_input_template.xlsx", mime=XLSX_MIME,
                width="stretch",
            )
            # Not "more accurate": the same engine runs either way, and where
            # the inference is right the numbers are identical. What the
            # template removes is the guessing, which is a claim that can be
            # checked - unlike "more accurate", which invites the question of
            # when the other path is wrong.
            st.caption(
                "The columns are already named, so nothing has to be inferred: "
                "no column mapping, no chart of accounts, no expense types read "
                "off account names, and no subtotal rows to filter out."
            )
        st.markdown("")
        st.caption(
            "**The pack adapts to your columns.** An expense-type column adds the "
            "Expense Report, a department column adds Departments & Cost Centres, "
            "and more than one entity adds a consolidation sheet."
        )

    if up is not None:
        raw = ingest.load_file(up)
        st.markdown("#### 2 · Confirm the column mapping")
        st.caption("Flux guessed these from your headers and data. Override any that look wrong.")

        matches = match_columns(list(raw.columns))
        _content_detect(raw, matches)
        auto = {m.field: m.source for m in matches}
        ignored = ingest.unmapped_columns(raw, matches)
        if ignored:
            st.caption(
                f"{len(raw.columns)} columns read · {len(ignored)} not needed for the "
                f"report and ignored: " + ", ".join(map(str, ignored[:8]))
                + (" …" if len(ignored) > 8 else "")
            )

        options = ["(none)"] + list(raw.columns)
        overrides = {}
        mcols = st.columns(4)
        for i, fieldname in enumerate(CANONICAL):
            with mcols[i % 4]:
                default = auto.get(fieldname) or "(none)"
                idx = options.index(default) if default in options else 0
                label = fieldname + (" *" if fieldname in REQUIRED else "")
                pick = st.selectbox(label, options, index=idx, key=f"map_{fieldname}")
                if pick != "(none)":
                    overrides[fieldname] = pick

        final = match_columns(list(raw.columns), overrides)
        _content_detect(raw, final)
        missing = [f for f in REQUIRED if not any(m.field == f and m.source for m in final)]
        if missing:
            st.warning(f"Still need a column for: {', '.join(missing)}")
        else:
            std = apply_mapping(raw, final)
            # A row silently dropped changes the totals as surely as a subtotal
            # silently counted, so both are said out loud.
            for note in std.attrs.get("dropped_rows", []):
                st.info(note)
            # The classification step only exists when there was something to
            # classify, so the steps after it are numbered from what came before.
            step = 3
            cat_info = std.attrs.get("category_inference")
            exp_info = std.attrs.get("expense_type_inference")
            std, notes = ingest.normalise_signs(std)

            # The column mapping is confirmed above; the classification is the
            # other guess in the pipeline and it was never shown. A misread
            # chart of accounts produces a finished-looking pack with the bottom
            # line inverted, so it gets the same treatment: say what was assumed,
            # then let it be changed.
            if cat_info is not None:
                st.markdown("#### 3 · Confirm the account classification")
                warnings = ingest.category_issues(
                    list(std["account_code"]), list(std["account_name"]),
                    list(std["category"]), cat_info)
                if exp_info is not None:
                    warnings += ingest.expense_type_issues(exp_info)
                if cat_info["style"]:
                    st.caption(
                        f"Read as a {cat_info['style']} chart of accounts, "
                        f"{int(cat_info['confidence'] * 100)}% corroborated by the "
                        "account names. Change anything that looks wrong."
                    )
                else:
                    st.caption("No known chart of accounts matched, so accounts "
                               "were classified from their names. Change anything "
                               "that looks wrong.")
                for w in warnings:
                    st.warning(w)

                editable = std[["account_code", "account_name", "category"]].copy()
                if "expense_type" in std.columns:
                    editable["expense_type"] = std["expense_type"]
                edited = st.data_editor(
                    editable, key="category_editor", hide_index=True,
                    width="stretch",
                    column_config={
                        "account_code": st.column_config.TextColumn("Account", disabled=True),
                        "account_name": st.column_config.TextColumn("Name", disabled=True),
                        "category": st.column_config.SelectboxColumn(
                            "Category", options=["Revenue", "COGS", "OpEx", "Other"],
                            required=True),
                        "expense_type": st.column_config.SelectboxColumn(
                            "Expense type", options=[""] + EXPENSE_TYPES
                            + [ingest.UNCLASSIFIED_EXPENSE]),
                    },
                )
                std["category"] = edited["category"].values
                if "expense_type" in edited.columns:
                    std["expense_type"] = edited["expense_type"].fillna("").values
                step = 4

            # The reporting month follows from the data: the latest period with
            # postings. Several periods also give a year-to-date view.
            active, found = ingest.reporting_period(std)
            if not found:
                active = st.text_input(
                    "This file has no period column — label the report as",
                    value=date.today().strftime("%Y-%m"), key="period_label",
                )
            if len(found) > 1:
                notes.append(
                    f"File covers {found[0]} to {found[-1]} — reporting {active} "
                    "with year to date."
                )
            elif found:
                notes.append(f"Reporting period {active}.")

            # Posting lines collapse while keeping the analytical dimensions.
            n_raw = len(std)
            std = ingest.aggregate_to_accounts(std)
            if n_raw > len(std):
                notes.append(f"{n_raw} posting lines aggregated to {len(std)} rows.")

            budgeted = has_budget(std)

            if up_bud is not None:
                braw = ingest.load_file(up_bud)
                bmatch = match_columns(list(braw.columns))
                _content_detect(braw, bmatch)
                bmiss = [f for f in REQUIRED if not any(m.field == f and m.source for m in bmatch)]
                if bmiss:
                    st.warning(f"Budget file: could not identify {', '.join(bmiss)}.")
                else:
                    bstd, _ = ingest.normalise_signs(apply_mapping(braw, bmatch))
                    bstd = ingest.aggregate_to_accounts(bstd)
                    std, bnotes = ingest.merge_budget(std, bstd)
                    notes += bnotes
                    budgeted = has_budget(std)
            elif not budgeted:
                st.info(
                    "No budget column found in this file — it looks like an actuals-only "
                    "extract. The pack will report actuals and leave the variance columns "
                    "empty rather than compare against zero. Upload the budget as a second "
                    "file above to get the full variance pack."
                )

            # Every dimension present adds a sheet, so say which were found
            # before the notes are rendered.
            dims = [d for d in ("expense_type", "department", "cost_centre", "entity")
                    if d in std.columns and (std[d].astype(str).str.strip() != "").any()]
            if dims:
                notes.append("Dimensions found: " + ", ".join(dims)
                             + " — the pack adds a sheet for each.")

            for n in notes:
                st.caption(n)

            preview = std
            if "period" in std.columns and len(found) > 1:
                preview = ingest.filter_period(std, active)

            st.markdown(f"#### {step} · Preview")
            with st.expander("Standardised data (fed to the engine)"):
                st.dataframe(std, width="stretch", hide_index=True)
            _pnl_preview(preview, budgeted)

            st.markdown(f"#### {step + 1} · Download the pack")
            st.caption(
                "The workbook is the working file — every figure a live formula "
                "over the input sheet. The deck is the management slides you "
                "send upward, drawn from the same numbers."
            )
            gen_x, gen_p = st.columns(2)
            if gen_x.button("Generate Excel pack", type="primary", width="stretch"):
                with gen_x:
                    _build("Writing the workbook — live formulas over your ledger…",
                           "Excel pack", "pack_upload",
                           lambda o: build_client_pack(std, active, o,
                                                       budgeted=budgeted),
                           Path(tempfile.gettempdir()) / "flux_pack.xlsx")
            if gen_p.button("Generate PowerPoint deck", width="stretch"):
                with gen_p:
                    ytd_frame = ingest.year_to_date(std, active)
                    _build("Building the slides…", "Deck", "deck_upload",
                           lambda o: build_pptx_pack(
                               preview, active, o, budgeted=budgeted,
                               ytd=ytd_frame if not ytd_frame.empty else None,
                               months=len(found) or None),
                           Path(tempfile.gettempdir()) / "flux_management_pack.pptx")
            with gen_x:
                _render_download("pack_upload", "Download flux_pack.xlsx")
            with gen_p:
                _render_download("deck_upload", "Download flux_management_pack.pptx",
                                 PPTX_MIME)

# ---------------------------------------------------------------------------
with tab_sample:
    st.markdown("#### Sample company · full multi-entity pack")
    st.caption(
        "Three legal entities in EUR, USD and HUF, transaction-level actuals and "
        "a full-year budget. Generates the complete pack: P&L, Expense Report, "
        "By Entity, Departments & CCs, Drivers."
    )
    sample = generate_month(SAMPLE_PERIOD).drop(columns="period")
    _pnl_preview(sample, True)
    col_x, col_p = st.columns(2)
    if col_x.button("Generate Excel pack", type="primary", width="stretch"):
        with col_x:
            _build("Generating the ledger and writing seven sheets…",
                   "Excel pack", "pack_sample",
                   lambda o: build_demo_pack(SAMPLE_PERIOD, o),
                   Path(tempfile.gettempdir()) / "flux_sample_pack.xlsx")
    if col_p.button("Generate PowerPoint deck", width="stretch"):
        with col_p:
            _build("Building the slides…", "Deck", "deck_sample",
                   lambda o: build_pptx_pack(
                       sample, SAMPLE_PERIOD, o,
                       detail=monthly_detail(SAMPLE_PERIOD),
                       ytd=_sample_ytd(), fy_budget=_sample_fy_budget(),
                       months=int(SAMPLE_PERIOD[5:7]),
                       analysis_blocks=_demo_blocks(),
                       ytd_detail=_demo_ytd_detail()),
                   Path(tempfile.gettempdir()) / "flux_sample_deck.pptx")
    with col_x:
        _render_download("pack_sample", "Download flux_sample_pack.xlsx")
    with col_p:
        _render_download("deck_sample", "Download flux_sample_deck.pptx", PPTX_MIME)

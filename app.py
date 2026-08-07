"""
Flux - FP&A reporting app.

Upload a GL / trial-balance file, confirm the column mapping, and download a
management pack. Or use built-in sample data to see the full multi-entity
showcase without uploading anything.

Run:  streamlit run app.py
"""

import sys
import tempfile
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
from flux.engine import build_report, has_budget
from flux.reporting import build_client_pack, build_demo_pack
from flux.synthetic_data import generate_month


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

# The page heading uses the theme's primary colour rather than a hardcoded navy,
# so it stays legible in dark mode. See .streamlit/config.toml.
st.title(":primary[Flux]")
st.caption("Automated management reporting — turn a general ledger into a variance pack.")

SAMPLE_PERIOD = "2025-06"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

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


def _offer_download(path: Path, label: str, key: str):
    """Hand the finished file to the browser.

    The bytes are held in session state rather than rebuilt inside the button
    block: clicking a download button triggers a Streamlit rerun, which would
    otherwise collapse the block that created it and make the button vanish.
    """
    st.session_state[key] = {"name": path.name, "data": path.read_bytes()}


def _render_download(key: str, label: str):
    payload = st.session_state.get(key)
    if not payload:
        return
    st.download_button(label, payload["data"], file_name=payload["name"],
                       mime=XLSX_MIME, width="stretch", key=f"dl_{key}")


tab_upload, tab_sample = st.tabs(["Upload your data", "Use sample data"])

# ---------------------------------------------------------------------------
with tab_upload:
    st.markdown("#### 1 · Upload your data")
    col_a, col_b = st.columns([3, 1])
    up = col_a.file_uploader(
        "Actuals — GL extract or trial balance", type=["csv", "xlsx", "xls"],
        accept_multiple_files=False, key="up_actuals",
    )
    with col_b:
        tpath = Path(__file__).parent / "data" / "input_template.xlsx"
        if tpath.exists():
            st.download_button(
                "Download template", tpath.read_bytes(),
                file_name="flux_input_template.xlsx", mime=XLSX_MIME,
                width="stretch",
            )
            st.caption("Optional — Flux reads most exports as they are.")

    up_bud = st.file_uploader(
        "Budget — optional second file, if the plan lives outside the ledger",
        type=["csv", "xlsx", "xls"], accept_multiple_files=False, key="up_budget",
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
            std, notes = ingest.normalise_signs(std)

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

            st.markdown("#### 3 · Preview")
            with st.expander("Standardised data (fed to the engine)"):
                st.dataframe(std, width="stretch", hide_index=True)
            _pnl_preview(preview, budgeted)

            st.markdown("#### 4 · Download the pack")
            if st.button("Generate pack", type="primary", width="stretch"):
                out = Path(tempfile.gettempdir()) / "flux_pack.xlsx"
                build_client_pack(std, active, out, budgeted=budgeted)
                _offer_download(out, "Download flux_pack.xlsx", "pack_upload")
            _render_download("pack_upload", "Download flux_pack.xlsx")

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
    if st.button("Generate sample pack", type="primary", width="stretch"):
        out = Path(tempfile.gettempdir()) / "flux_sample_pack.xlsx"
        build_demo_pack(SAMPLE_PERIOD, out)
        _offer_download(out, "Download flux_sample_pack.xlsx", "pack_sample")
    _render_download("pack_sample", "Download flux_sample_pack.xlsx")

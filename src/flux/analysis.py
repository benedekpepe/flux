"""
Variance analysis: what the numbers point at.

The commentary column says what moved and by how much. This says what that
pattern implies - how concentrated the movement is, how long it has been
running, where it puts the full year, and which question it makes worth asking.

What it will not do is say *why*. The ledger does not carry causes. A line text
in a real extract reads "Invoice 88213" or "Reclass to 5211", not "campaign
overspend", so any sentence beginning "because" would be invented - and an
invented cause in a management pack is worse than no analysis at all, because it
is the one thing the reader cannot check against the numbers beside it.

The distinction the module is built on: a one-month spike and a six-month drift
are the same variance and a different problem. One is a timing question for the
accountant, the other a run-rate question for the budget owner. Which one it is
follows from the data, so the pack can say it.
"""

from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from .commentary import _join_names, _mag, _pct
from .engine import MaterialityRule


@dataclass(frozen=True)
class Finding:
    """One line of analysis: a short heading and the sentence under it."""
    heading: str
    text: str


#: Below this share, "a few lines explain most of it" is not true and saying it
#: would be worse than saying nothing.
CONCENTRATION_FLOOR = 0.55
#: How many drivers to name before the list stops being a list.
MAX_DRIVERS = 3


def concentration(drivers: pd.DataFrame, total_var: float, *,
                  name_col: str = "account_name") -> Finding | None:
    """How much of the movement sits in how few lines.

    A variance spread evenly over forty accounts and one sitting in three are
    different problems, and the totals row cannot tell them apart. Only the
    lines moving the *same way* as the total count: an offsetting favourable
    line is not part of what has to be explained.
    """
    if drivers.empty or not total_var:
        return None
    same_way = drivers[drivers["var_bud"] * total_var > 0].copy()
    if same_way.empty:
        return None
    same_way = same_way.reindex(
        same_way["var_bud"].abs().sort_values(ascending=False).index)
    top = same_way.head(MAX_DRIVERS)
    share = top["var_bud"].sum() / total_var
    if share < CONCENTRATION_FLOOR:
        return Finding(
            "Concentration",
            f"The movement is spread across {len(same_way)} lines rather than "
            f"sitting in a few: the largest {len(top)} account for "
            f"{_pct(share)} of it. There is no single driver to chase.")
    named = _join_names([f"{r[name_col].lower()} ({_mag(r['var_bud'])})"
                         for _, r in top.iterrows()])
    plural = "line accounts" if len(top) == 1 else "lines account"
    if share >= 0.99:
        # Over 100% is arithmetically right and reads as an error: the named
        # lines exceed the net movement because something else offsets them,
        # which is worth saying rather than printing "101%".
        offset = same_way["var_bud"].sum() - total_var
        tail = (f" - more than the net movement, because {_mag(offset)} of it "
                "is offset by lines going the other way"
                if abs(offset) > abs(total_var) * 0.02 else "")
        return Finding("Concentration",
                       f"Effectively all of the movement sits in {len(top)} "
                       f"{'line' if len(top) == 1 else 'lines'}: {named}{tail}.")
    return Finding(
        "Concentration",
        f"{len(top)} {plural} for {_pct(share)} of it: {named}.")


def persistence(history: pd.DataFrame, *, higher_is_better: bool) -> Finding | None:
    """How many of the months so far went the wrong way.

    `history` is one row per period with `actual` and `budget`. This is the
    difference between a bad month and a bad year, and it is the first thing a
    reader asks after seeing a variance.

    It counts, and stops there. An earlier version also said whether the gap was
    widening or narrowing, and that was a claim the pack could not support: the
    only other figure describing the series is the run rate, which is
    `year to date / months x 12` - an average, and an average cannot tell a
    widening gap from a narrowing one from a single spike. Three ledgers with
    identical run rates and opposite trends produced the same projection and
    three different verdicts here, with nothing on the page to settle which was
    right. Counting is verifiable against the months themselves; the shape of
    the series is not, until the pack shows it.
    """
    if history is None or len(history) < 2:
        return None
    hist = history.sort_values("period_no").copy()
    hist["var"] = hist["actual"] - hist["budget"]
    adverse = hist["var"] < 0 if higher_is_better else hist["var"] > 0
    n_bad, n = int(adverse.sum()), len(hist)

    if n_bad == 0:
        return Finding("Persistence",
                       f"Not once in {n} months has this line been on the wrong "
                       "side of plan.")
    # The noun belongs to the total, not to the count: "1 of 6 months".
    word = "month" if n == 1 else "months"
    return Finding("Persistence", f"Adverse in {n_bad} of {n} {word}.")


def full_year(run_rate: float, fy_budget: float, ytd_actual: float,
              months_elapsed: int, *, higher_is_better: bool) -> Finding | None:
    """Where the year lands, and what the rest of it would have to do.

    The second half is the part a reader acts on. "€511k over plan" is a fact;
    "the remaining six months would have to come in €511k under plan" is the
    same fact in the units of the decision they are about to make.
    """
    if not months_elapsed or months_elapsed >= 12 or not fy_budget:
        return None
    gap = run_rate - fy_budget
    adverse = gap < 0 if higher_is_better else gap > 0
    if abs(gap) < 1:
        return Finding("Full year",
                       "At the current run rate the year lands on plan.")
    left = 12 - months_elapsed
    remaining_plan = fy_budget - ytd_actual
    direction = "over" if gap > 0 else "under"
    verdict = "" if adverse else " - favourable, but worth confirming it is not timing"
    return Finding(
        "Full year",
        f"At the current run rate the year lands {_mag(gap)} {direction} a "
        f"{_mag(fy_budget)} plan{verdict}. Holding the plan leaves "
        f"{_mag(remaining_plan)} for the remaining {left} months, against "
        f"{_mag(ytd_actual / months_elapsed)} a month so far.")


def question(history: pd.DataFrame, *, higher_is_better: bool) -> Finding | None:
    """The question the shape of the variance makes worth asking.

    Not a cause - the ledger does not hold one - but the right question to put
    to the person who does. A single month out of line is an accounting
    question: accrual, cut-off, a double posting, an invoice that landed early.
    A steady drift is a planning question: the plan and the activity level
    disagree, and one of them is going to have to move.
    """
    if history is None or len(history) < 3:
        return None
    hist = history.sort_values("period_no").copy()
    hist["var"] = hist["actual"] - hist["budget"]
    adverse = hist["var"] < 0 if higher_is_better else hist["var"] > 0
    n_bad, n = int(adverse.sum()), len(hist)
    if n_bad == 0:
        return None

    if n_bad == 1:
        when = hist.loc[adverse, "period"].iloc[0]
        return Finding(
            "Ask",
            f"One month out of line ({when}) against {n - 1} on plan. That "
            "shape is usually timing rather than level - worth checking the "
            "cut-off, any accrual released into the month, and whether a "
            "single invoice landed in the wrong period before treating it as "
            "a change in run rate.")
    if n_bad >= n - 1:
        return Finding(
            "Ask",
            f"Adverse in {n_bad} of {n} months, which is a level rather than "
            "an event: the plan and the actual activity disagree. Worth asking "
            "the budget owner whether the plan was set before the current "
            "activity level, and whether the run rate is expected to hold - "
            "one of the two numbers has to move.")
    return Finding(
        "Ask",
        f"Adverse in {n_bad} of {n} months, so neither a one-off nor a settled "
        "level. Worth looking at what the months differ by before deciding "
        "whether to re-plan or to treat it as noise.")


def findings(*, drivers: pd.DataFrame, total_var: float, history: pd.DataFrame,
             run_rate: float, fy_budget: float, ytd_actual: float,
             months_elapsed: int, higher_is_better: bool,
             name_col: str = "account_name",
             materiality: MaterialityRule | None = None) -> list[Finding]:
    """The whole analysis for one line, in the order a reader wants it.

    Concentration first (what is it), then persistence (how long), then the
    full year (what it costs), then the question (what to do about it). Any
    part that the data cannot support is left out rather than padded.
    """
    out = [
        concentration(drivers, total_var, name_col=name_col),
        persistence(history, higher_is_better=higher_is_better),
        full_year(run_rate, fy_budget, ytd_actual, months_elapsed,
                  higher_is_better=higher_is_better),
        question(history, higher_is_better=higher_is_better),
    ]
    return [f for f in out if f is not None]


NO_CAUSE_NOTE = (
    "This analysis is derived from the figures on this sheet: how concentrated "
    "the movement is, how many months it has run, and where it puts the full "
    "year. It does not say why a line moved, because the ledger does not record "
    "why - it points at the question to ask and who is likely to hold the answer."
)

"""
Ingestion layer: map an arbitrary client file onto the Flux internal schema.

Real GL / trial-balance exports never use the same column names. This module
resolves them in layers, cheapest and most certain first:

  1. Synonym dictionary   - known aliases (English + Hungarian), normalised.
  2. Fuzzy matching       - closest header by string similarity, above a cutoff.
  3. Content detection    - infer from the data when the header is unhelpful
                            (e.g. 4-digit account codes, numeric amounts).
  4. (LLM fallback)       - hook left for a model call on the hard leftovers.

The result is an approvable/overridable mapping (canonical field -> source
column, with the method and score), so a person can confirm or correct it once
and the mapping can be saved per client. The engine is unchanged: apply_mapping
returns the exact account-level shape build_report expects.

Target schema (account level):
  required : account_code, account_name, actual, budget
  optional : prior_year, category, cost_centre, entity
Category, if absent, is inferred from the account-code range.
"""

# Deliberately no `from __future__ import annotations` here.
#
# With string annotations, dataclasses has to guess whether each one is the
# KW_ONLY marker, and does so by looking the class's module up in sys.modules.
# In an environment where the module is not registered there - a hot-reloading
# host, a loader that execs the module itself - that lookup returns None and
# the import dies with an unrelated-looking AttributeError before any of this
# code runs. The annotations in this module are all evaluatable at definition
# time, so there is nothing to gain from deferring them and a deployment to
# lose.
from dataclasses import dataclass
from difflib import SequenceMatcher
import os
from io import StringIO
from pathlib import Path
import re

import pandas as pd

HAS_LLM = bool(os.environ.get("ANTHROPIC_API_KEY"))


CANONICAL = ["account_code", "account_name", "actual", "budget", "prior_year",
             "debit", "credit", "dc_indicator",
             "period", "category", "expense_type", "department", "cost_centre", "entity"]
# A file must identify its accounts and carry at least one amount column.
# Actuals extracts have "actual"; budget extracts have only the plan figures;
# a trial balance often has separate debit and credit columns instead.
REQUIRED = ["account_code", "account_name"]
AMOUNT_FIELDS = ["actual", "budget", "debit", "credit"]

# Known aliases per canonical field (English + Hungarian), stored normalised.
_RAW_SYNONYMS = {
    "account_code": ["account", "account code", "acct", "acct no", "account no",
                     "account number", "gl account", "gl code", "gl", "nominal",
                     "nominal code", "code", "szamla", "szamlaszam", "fokonyvi szam",
                     "fokonyvi szamla", "number"],
    "account_name": ["account name", "name", "description", "account description",
                     "account desc", "gl name", "gl description", "account text",
                     "megnevezes", "szamla megnevezes", "szamla nev", "leiras"],
    "actual": ["actual", "actuals", "amount", "amt", "value", "actual amount",
               "ytd actual", "balance", "net", "net amount", "osszeg", "teny",
               "actual eur", "amount eur", "current", "current year", "tenyleges"],
    "budget": ["budget", "budgeted", "plan", "planned", "forecast", "budget amount",
               "terv", "budget eur", "bud", "tervezett", "keret"],
    "prior_year": ["prior year actual", "prior year actuals", "py actual", "prior yr act",
                   "prior year", "prior", "py", "last year", "ly", "prior yr",
                   "previous year", "last year actual", "elozo ev", "elozo evi teny",
                   "tavaly", "bazis"],
    # A trial balance often splits the amount into two columns rather than
    # signing one. Both are read and netted to a single actual.
    "debit": ["debit", "debit amount", "dr", "dr amount", "debits", "tartozik",
              "tartozik osszeg", "terheles", "t oldal"],
    "credit": ["credit", "credit amount", "cr", "cr amount", "credits", "kovetel",
               "kovetel osszeg", "jovairas", "k oldal"],
    # A single amount column plus a debit/credit marker per line, the way SAP
    # and most ERP line-item reports are exported.
    "dc_indicator": ["dc", "d c", "dc indicator", "debit credit", "debit credit indicator",
                     "debit/credit", "posting indicator", "sign", "d/c",
                     "tartozik kovetel", "t k", "egyenleg iranya"],
    "category": ["category", "type", "account type", "pnl category", "kategoria",
                 "csoport", "class", "tipus"],
    "department": ["department", "dept", "division", "function", "org unit",
                   "organisational unit", "reszleg", "osztaly", "divizio", "terulet"],
    "cost_centre": ["cost centre", "cost center", "cost centre code", "cost center code",
                    "cc code", "cc no", "cost centre no", "cc id", "cctr", "cc",
                    "koltseghely", "koltseghely kod", "kh kod", "kh",
                    "koltseghely azonosito", "szervezeti egyseg"],
    "period": ["period", "accounting period", "posting period", "fiscal period",
               "reporting period", "year month", "yearmonth", "month", "per",
               "idoszak", "konyvelesi idoszak", "elszamolasi idoszak", "honap",
               "periodus", "targyhonap"],
    "expense_type": ["expense type", "cost type", "expense category", "cost category",
                     "nature", "expense nature", "koltsegnem", "koltseg tipus",
                     "koltseg kategoria", "kiadas tipusa"],
    "entity": ["entity", "company", "legal entity", "company code", "subsidiary",
               "entitas", "ceg", "tarsasag", "tarsasagi kod"],
}

_FUZZY_CUTOFF = 0.82


def _normalize(s: str) -> str:
    s = str(s).lower().strip()
    s = re.sub(r"[^0-9a-zA-Zaaeeiioouuoou\s]", " ", s)  # keep basic letters/digits
    s = re.sub(r"\s+", " ", s).strip()
    return s


_SYNONYMS = {k: {_normalize(v) for v in vs} for k, vs in _RAW_SYNONYMS.items()}


@dataclass
class ColMatch:
    field: str
    source: str | None
    method: str            # synonym | fuzzy | content | override | unmatched
    score: float


def _fuzzy(a: str, b: str) -> float:
    return SequenceMatcher(None, a, b).ratio()


def match_columns(headers: list[str], overrides: dict | None = None) -> list[ColMatch]:
    """Resolve each canonical field to a source column."""
    overrides = overrides or {}
    norm = {h: _normalize(h) for h in headers}
    used: set[str] = set()
    matches: list[ColMatch] = []

    for field_name in CANONICAL:
        # 0. explicit override wins.
        if field_name in overrides and overrides[field_name] in headers:
            src = overrides[field_name]
            matches.append(ColMatch(field_name, src, "override", 1.0)); used.add(src); continue

        # 1. exact synonym match.
        hit = None
        for h in headers:
            if h in used:
                continue
            if norm[h] in _SYNONYMS[field_name]:
                hit = h; break
        if hit:
            matches.append(ColMatch(field_name, hit, "synonym", 1.0)); used.add(hit); continue

        # 2. fuzzy against synonym set.
        best, best_score = None, 0.0
        for h in headers:
            if h in used:
                continue
            score = max((_fuzzy(norm[h], syn) for syn in _SYNONYMS[field_name]), default=0.0)
            if score > best_score:
                best, best_score = h, score
        if best is not None and best_score >= _FUZZY_CUTOFF:
            matches.append(ColMatch(field_name, best, "fuzzy", round(best_score, 2))); used.add(best); continue

        matches.append(ColMatch(field_name, None, "unmatched", 0.0))
    return matches


def _looks_like_money(col: pd.Series) -> bool:
    """Conservative test for a monetary column.

    Guards against the classic failure of treating document numbers, years or
    line counters as amounts: those are integer-like, sequential or clustered in
    a narrow range, whereas money varies widely and usually carries decimals.
    """
    vals = col.map(_coerce_amount).dropna()
    if len(vals) < 5:
        return False
    if vals.abs().max() == 0:
        return False
    # Year-like or ID-like: all integers within a tight band.
    integers = (vals % 1 == 0).mean() > 0.98
    spread = vals.abs().max() / max(vals.abs().min(), 1e-9)
    if integers and spread < 3:
        return False
    # Monotonic sequences are counters or document numbers, not amounts.
    if vals.is_monotonic_increasing and integers:
        return False
    return True


def _content_detect(df: pd.DataFrame, matches: list[ColMatch]) -> None:
    """Fill still-unmatched fields from the data itself, in place.

    Only structural fields are inferred. Monetary fields are never guessed from
    an unnamed column: silently summing the wrong column is far worse than
    asking the person which column to use.
    """
    used = {m.source for m in matches if m.source}
    free = [c for c in df.columns if c not in used]
    by_field = {m.field: m for m in matches}

    # account_code: a column that is mostly 3-6 digit codes.
    if by_field["account_code"].source is None:
        for c in free:
            vals = df[c].dropna().astype(str).str.strip()
            if len(vals) and (vals.str.fullmatch(r"\d{3,6}").mean() > 0.7):
                by_field["account_code"].source = c
                by_field["account_code"].method = "content"
                by_field["account_code"].score = 0.75
                used.add(c); free.remove(c); break

    # actual: only when exactly one free column convincingly looks monetary.
    if by_field["actual"].source is None:
        candidates = [c for c in free if _looks_like_money(df[c])]
        if len(candidates) == 1:
            c = candidates[0]
            by_field["actual"].source = c
            by_field["actual"].method = "content"
            by_field["actual"].score = 0.6


_TRAILING_MARKER = re.compile(r"(?i)[\s.]*\b(cr|dr)\b[\s.]*$")
_BLANKS = {"", "na", "n/a", "n.a.", "-", "\u2013", "\u2014", "nil", "none", "null"}


def _coerce_amount(x):
    """Parse a number that may use EU or US formatting; return float or NaN.

    Real ledger exports write a negative in at least four ways, and getting any
    of them wrong is worse than failing loudly: the sign flips or the row is
    silently dropped, and the report still looks finished. Handled here:

        -1.234,50     leading minus
        (1.234,50)    accounting parentheses
        1.234,50-     trailing minus (the SAP default)
        1,234.50 CR   credit marker

    Thousands and decimal separators are resolved from the rightmost separator,
    so both 1.234,50 and 1,234.50 read correctly.
    """
    if x is None:
        return float("nan")
    if isinstance(x, bool):
        return float("nan")
    if isinstance(x, (int, float)):
        return float(x)

    s = str(x).strip()
    # Normalise the punctuation real exports use: unicode minus, non-breaking
    # and thin spaces as thousands separators.
    s = (s.replace("\u2212", "-").replace("\u00a0", " ")
          .replace("\u202f", " ").replace("\u2009", " ").strip())
    if s.lower() in _BLANKS:
        return float("nan")

    negative = False

    # Accounting parentheses, optionally around a currency symbol.
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1].strip()

    # Credit / debit marker, before the symbol strip removes the letters.
    marker = _TRAILING_MARKER.search(s)
    if marker:
        if marker.group(1).lower() == "cr":
            negative = not negative
        s = s[:marker.start()].strip()

    s = re.sub(r"[^\d,.\-]", "", s)          # drop currency symbols, codes, spaces

    # Trailing minus, as written by SAP and several Hungarian systems.
    if s.endswith("-"):
        negative = not negative
        s = s[:-1]
    if s.startswith("-"):
        negative = not negative
        s = s[1:]
    s = s.replace("-", "")                   # any minus left is stray

    has_dot, has_comma = "." in s, "," in s
    if has_dot and has_comma:
        # rightmost separator is the decimal point.
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif has_comma:
        if re.fullmatch(r"-?\d{1,3}(,\d{3})+", s):   # 1,234,567 -> thousands
            s = s.replace(",", "")
        elif re.search(r",\d{1,2}$", s):             # 12,5 -> decimal
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    elif has_dot:
        if re.fullmatch(r"-?\d{1,3}(\.\d{3})+", s):   # 445.000 / 1.234.567 -> thousands
            s = s.replace(".", "")
        # otherwise keep as a decimal point
    try:
        value = float(s)
    except ValueError:
        return float("nan")
    return -value if negative else value


def year_to_date(std: pd.DataFrame, period: str) -> pd.DataFrame:
    """Aggregate every period up to and including `period` to account level.

    A file that spans several months already holds the year to date; this rolls
    it up so the cumulative view can be reported beside the month. Returns an
    empty frame when the data carries no periods, so callers can drop the
    cumulative view rather than present the month twice under two labels.
    """
    if "period" not in std.columns:
        return std.iloc[0:0]
    cutoff = period_key(period)
    if cutoff is None:
        return std.iloc[0:0]
    keys = std["period"].map(period_key)
    upto = std[keys.notna() & (keys <= cutoff)]
    if upto.empty or upto["period"].nunique() < 2:
        return std.iloc[0:0]
    return aggregate_to_accounts(upto.drop(columns=["period"]))



# ---------------------------------------------------------------------------
# Category inference
# ---------------------------------------------------------------------------
# The account code alone cannot say what a line is. "5" opens material cost in a
# US chart, every cost by nature in a Hungarian one, and an expense account in
# SAP; "8" is revenue in SAP and a cost in Hungary. Reading the first digit and
# hoping was silently reporting Hungarian revenue as an operating expense, and
# an SAP chart with its revenue and materials swapped - a finished-looking pack
# with the bottom line inverted, which is the worst way to be wrong.
#
# So two independent readings are taken - the account name and the account code
# - and they are made to agree before either is trusted.

CATEGORY_KEYWORDS: list[tuple[str, tuple[str, ...]]] = [
    # Order matters: the first match wins, so the narrow classes are tested
    # before the broad ones. "Interest income" is financing, not revenue.
    ("Other", (
        "ertekcsokken", "ecs leiras", "terven feluli", "amortiz", "depreciat",
        "kamat", "interest", "arfolyam", "exchange difference", "fx ",
        "penzugyi muveletek", "financial expense", "financial income",
        "tarsasagi ado", "income tax", "corporate tax", "deferred tax",
        "rendkivuli", "extraordinary", "impairment", "ertekveszt",
        "abschreibung", "zinsen", "steuer vom einkommen", "ausserordentlich",
    )),
    ("Revenue", (
        "arbevetel", "ertekesites", "bevetel", "revenue", "sales", "turnover",
        "income from", "fee income", "subscription income", "net sales",
        "erlose", "umsatzerlose", "umsatz", "ertrage aus",
    )),
    ("COGS", (
        "anyagkoltseg", "anyagjellegu", "kozvetlen", "elabe", "eladott aru",
        "eladott (koz)vetitett", "alapanyag", "alkatresz",
        "cost of sales", "cost of goods", "cogs", "direct labour",
        "direct labor", "material", "component", "merchandise", "purchases",
        "wareneingang", "wareneinsatz", "rohstoffe", "fremdleistung",
        "freight", "fuvar", "szallitasi koltseg", "hosting", "infrastructure",
    )),
    ("OpEx", (
        "berkoltseg", "ber ", "berek", "szemelyi jellegu", "bergarulek",
        "jarulek", "szocialis hozzajarulas", "premium", "jutalom", "bonusz",
        "igenybe vett szolgaltatas", "egyeb szolgaltatas", "berleti",
        "iroda", "rezsi", "marketing", "hirdetes", "reklam", "biztositas",
        "utazas", "kikuldetes", "reprezentacio", "szoftver", "licenc",
        "konyveles", "ugyved", "tanacsadas", "alvallalkozo", "megbizasi",
        "salary", "salaries", "wage", "payroll", "personnel", "staff",
        "advertis", " ads", "ads ", "promotion", "rent", "office", "utilit",
        "facilit",
        "insurance", "travel", "entertainment", "software", "licence",
        "license", "consult", "legal", "audit", "professional",
        "contractor", "recruit", "training", "telephone", "telecom",
        "operating expense", "egyeb rafordit", "egyeb koltseg", "overhead",
        "occupancy", "repairs", "maintenance", "karbantartas",
        "lohne", "gehalter", "personalkosten", "miete", "raumkosten",
        "werbung", "versicherung", "reisekosten", "buromaterial",
    )),
]


def _fold(text) -> str:
    """Lowercase and strip accents, so Hungarian names match either spelling."""
    import unicodedata
    folded = unicodedata.normalize("NFKD", str(text).lower())
    return "".join(ch for ch in folded if not unicodedata.combining(ch))


def category_from_name(name) -> str | None:
    """The category a human would read off the account name, or None."""
    text = _fold(name)
    if not text.strip():
        return None
    for category, words in CATEGORY_KEYWORDS:
        if any(w in text for w in words):
            return category
    return None


def _hungarian(code: str) -> str | None:
    """The Hungarian statutory chart: 5 and 8 are costs, 9 is revenue."""
    two, one = code[:2], code[:1]
    if one == "9":
        return "Other" if two in {"97", "98", "99"} else "Revenue"
    if two in {"57", "86", "87", "88", "89"}:
        return "Other"
    if two in {"51", "52", "53"} or one == "8":
        return "COGS" if two in {"51", "81", "82"} else "OpEx"
    if one == "5":
        return "OpEx"
    return None


CHART_STYLES: dict[str, "callable"] = {
    # The chart the engine's own demo uses, and the common English/IFRS shape.
    "4-5-6-7": lambda c: {"4": "Revenue", "5": "COGS", "6": "OpEx",
                          "7": "Other"}.get(c[:1]),
    # SAP's standard chart: revenue in the 800s, primary costs in the 400s -
    # but the 400s are not one thing. 40-41 is material consumption, 42-47 is
    # personnel and other operating cost, 48 is depreciation.
    "SAP": lambda c: ({"40": "COGS", "41": "COGS", "48": "Other"}.get(c[:2])
                      or {"4": "OpEx", "8": "Revenue", "6": "OpEx",
                          "7": "OpEx", "2": "Other"}.get(c[:1])),
    # SKR03, the German small-business chart: revenue in the 8000s like SAP but
    # a different cost structure - 3 is goods inward, 4 is operating cost. It
    # collides with SAP on the 4s and 8s, which is exactly what the agreement
    # score against the account names is there to settle.
    "SKR03": lambda c: ({"48": "Other"}.get(c[:2])
                        or {"8": "Revenue", "3": "COGS", "4": "OpEx",
                            "2": "Other"}.get(c[:1])),
    "Hungarian": _hungarian,
}


def infer_categories(codes, names) -> tuple[list[str], dict]:
    """Classify every account, and say how confident the classification is.

    Each candidate chart style is scored on how often its reading of the account
    code agrees with the reading of the account name. The winner has to beat a
    floor, so an unrecognised chart falls back to the names rather than being
    forced into the shape of a chart it is not.

    Returns the categories and a dict describing what happened, so the caller
    can put it in front of the user instead of quietly proceeding.
    """
    codes = [str(c).strip() for c in codes]
    names = list(names)
    by_name = [category_from_name(n) for n in names]
    named = [i for i, c in enumerate(by_name) if c]

    scores, coverage = {}, {}
    for style, fn in CHART_STYLES.items():
        read = [fn(codes[i]) for i in named]
        agree = sum(1 for i, r in zip(named, read) if r == by_name[i])
        scores[style] = (agree / len(named)) if named else 0.0
        coverage[style] = sum(1 for c in codes if fn(c)) / max(1, len(codes))
        if not any(read):
            scores[style] = 0.0

    best = max(scores, key=scores.get) if scores else None
    confidence = scores.get(best, 0.0)
    # Below the floor the codes are telling a different story from the names,
    # which means the chart is one this list does not know.
    trusted = best if (confidence >= 0.6 and len(named) >= 3) else None
    blind = False
    if trusted is None and len(named) < 3:
        # An extract whose account names say nothing - "Konto 4001" - leaves the
        # codes as the only evidence there is. Refusing to read them would
        # classify the whole ledger as operating expense, which is worse than
        # reading them and saying so. The widest-covering chart wins.
        trusted = max(coverage, key=coverage.get)
        blind = coverage[trusted] > 0
        trusted = trusted if blind else None

    # A confirmed chart style wins over a single keyword hit. The style was
    # corroborated by most of the accounts on the file, which makes it
    # systematic evidence; one word in one name can be ambiguous - "subscription
    # revenue" and "software subscription" share a keyword and nothing else. The
    # names classify what the codes cannot, and every disagreement is reported.
    out, sources, disagreed = [], [], []
    for i, code in enumerate(codes):
        by_code = CHART_STYLES[trusted](code) if trusted else None
        if by_code:
            out.append(by_code); sources.append("code")
            if by_name[i] and by_name[i] != by_code:
                disagreed.append(f"{code} {names[i]} "
                                 f"(code: {by_code}, name: {by_name[i]})")
        elif by_name[i]:
            out.append(by_name[i]); sources.append("name")
        else:
            out.append("OpEx"); sources.append("default")

    info = {
        "style": trusted,
        "blind": blind,
        "confidence": round(confidence, 2),
        "coverage": {k: round(v, 2) for k, v in coverage.items()},
        "scores": {k: round(v, 2) for k, v in scores.items()},
        "named": len(named),
        "defaulted": [f"{codes[i]} {names[i]}" for i, s in enumerate(sources)
                      if s == "default"],
        "disagreed": disagreed,
        "categories": out,
        "sources": sources,
    }
    return out, info


def category_issues(codes, names, categories, info) -> list[str]:
    """What the user has to be told before they trust the pack.

    A misread chart produces a report that looks finished and has the bottom
    line inverted. Every one of these says the same thing in a different way:
    check the classification before you send this to anyone.
    """
    notes = []
    # Silence is the right answer when every account corroborates the chart:
    # a warning that always fires is a warning nobody reads. Anything less than
    # complete agreement is worth a sentence, because a misread chart produces a
    # finished-looking pack with the bottom line inverted.
    if info["style"] and not info.get("blind") and info["confidence"] >= 0.999 \
            and not info.get("disagreed") and "Revenue" in set(categories) \
            and not info["defaulted"]:
        return []
    if info.get("blind"):
        notes.append(
            f"No account name carried a usable signal, so the {info['style']} "
            "chart of accounts was chosen from the account codes alone and "
            "nothing corroborates it. Check the classification before relying "
            "on the P&L."
        )
    elif info["style"]:
        notes.append(
            f"Accounts classified using the {info['style']} chart of accounts, "
            f"which {int(info['confidence'] * 100)}% of the account names agree "
            "with. Check the classification if that is not your chart."
        )
    else:
        notes.append(
            "The account codes did not match any chart of accounts Flux knows "
            f"(best fit {int(info['confidence'] * 100)}%), so accounts were "
            "classified from their names alone. Review the classification "
            "before relying on the P&L."
        )
    if "Revenue" not in set(categories):
        notes.append(
            "No account was classified as revenue. Either this extract holds "
            "costs only, or the revenue accounts were misread - a P&L with no "
            "revenue reports the whole result as a loss."
        )
    # The loudest signal of a misread chart: a name that says revenue and a
    # classification that says otherwise.
    contradicted = [f"{codes[i]} {names[i]}" for i in range(len(codes))
                    if category_from_name(names[i]) == "Revenue"
                    and categories[i] != "Revenue"]
    if contradicted:
        notes.append(
            "These accounts read as revenue by name but were classified "
            f"otherwise: {'; '.join(contradicted[:6])}"
            f"{' and others' if len(contradicted) > 6 else ''}."
        )
    if info.get("disagreed"):
        notes.append(
            "The account name and the account code disagreed on these, and the "
            "code was followed because the chart of accounts is corroborated "
            f"across the file. Change any that are wrong: "
            f"{'; '.join(info['disagreed'][:5])}"
            f"{' and others' if len(info['disagreed']) > 5 else ''}."
        )
    if len(info["defaulted"]) > max(2, len(codes) // 5):
        notes.append(
            f"{len(info['defaulted'])} of {len(codes)} accounts could not be "
            "classified from either their code or their name and defaulted to "
            "operating expenses."
        )
    return notes


# ---------------------------------------------------------------------------
# Expense type inference
# ---------------------------------------------------------------------------
# The expense report is the view a cost owner reads, and it needs a natural
# classification the general ledger usually does not carry as a column. Where
# there is no column, the account name is the only evidence available - so it is
# used, and the fact that it was a guess is reported.
#
# The types are the ones in the chart of accounts, so an inferred file groups
# exactly like a native one.

EXPENSE_TYPE_KEYWORDS: list[tuple[str, tuple[str, ...]]] = [
    ("Depreciation & amortisation", ("ertekcsokken", "ecs", "amortiz", "depreciat")),
    ("Financing & bank", ("kamat", "banki", "bankkolts", "penzugyi muveletek",
                          "interest", "bank charge", "bank fee", "financing")),
    ("Payroll benefits", ("bergarulek", "jarulek", "szocialis hozzajarulas",
                          "szochо", "payroll tax", "social contribution",
                          "national insurance", "pension", "benefit")),
    ("Employee incentives", ("premium", "jutalom", "bonusz", "osztonzo",
                             "bonus", "incentive", "commission")),
    ("External staff / contractors", ("alvallalkozo", "megbizasi", "kolcsonzott",
                                      "contractor", "external staff", "agency",
                                      "temporary staff", "freelance")),
    ("Direct labour", ("kozvetlen ber", "direct labour", "direct labor")),
    ("Salaries & wages", ("berkoltseg", "berek", "ber ", "szemelyi jellegu",
                          "munkaber", "salary", "salaries", "wage", "payroll",
                          "personnel cost", "staff cost")),
    ("Marketing & advertising", ("marketing", "hirdetes", "reklam", "promocio",
                                 "advertis", "promotion", "campaign", "brand")),
    ("External professional fees", ("konyveles", "ugyved", "jogi", "tanacsadas",
                                    "konyvvizsgal", "audit", "legal", "consult",
                                    "professional fee", "advisory")),
    ("IT & software", ("szoftver", "informatik", "software", "saas", "it cost",
                       "subscription")),
    ("Third-party licences", ("licenc", "jogdij", "licence", "license",
                              "royalty", "third-party")),
    ("Hosting & infrastructure", ("hosting", "szerver", "felho", "cloud",
                                  "infrastructure", "datacent", "aws", "azure")),
    ("Facilities & office", ("berleti", "iroda", "rezsi", "kozuzemi", "takaritas",
                             "rent", "office", "utilit", "facilit", "cleaning",
                             "premises")),
    ("Travel & entertainment", ("utazas", "kikuldetes", "reprezentacio", "travel",
                                "entertainment", "hotel", "mileage", "subsistence")),
    ("Insurance", ("biztositas", "insurance")),
    ("Materials & components", ("anyagkoltseg", "anyagjellegu", "alapanyag",
                                "alkatresz", "anyag", "material", "component",
                                "raw ", "consumable")),
    ("Logistics & fees", ("szallitas", "fuvar", "logisztika", "logistics",
                          "freight", "shipping", "delivery", "courier",
                          "customs")),
]

#: Where a cost cannot be recognised. Named, not blank, so the line still
#: appears on the expense report instead of vanishing from a spend view.
UNCLASSIFIED_EXPENSE = "Other expenses"


def expense_type_from_name(name, category=None) -> str | None:
    """The natural expense type an account name implies, or None.

    Revenue accounts get nothing: the expense report is a spend view and a
    revenue line has no place on it.
    """
    if category == "Revenue":
        return None
    text = _fold(name)
    if not text.strip():
        return None
    for etype, words in EXPENSE_TYPE_KEYWORDS:
        if any(w in text for w in words):
            return etype
    return None


def infer_expense_types(names, categories) -> tuple[list[str], dict]:
    """Classify every cost account by natural expense type.

    Returns the types and a summary, because a guess the user cannot see is
    worse than no guess at all.
    """
    out, recognised = [], 0
    for name, category in zip(names, categories):
        if category == "Revenue":
            out.append("")
            continue
        etype = expense_type_from_name(name, category)
        if etype:
            recognised += 1
            out.append(etype)
        else:
            out.append(UNCLASSIFIED_EXPENSE)
    costs = sum(1 for c in categories if c != "Revenue")
    return out, {"recognised": recognised, "costs": costs,
                 "unclassified": costs - recognised}


def expense_type_issues(info) -> list[str]:
    if not info["costs"]:
        return []
    notes = [
        f"No expense-type column was supplied, so expense types were inferred "
        f"from the account names: {info['recognised']} of {info['costs']} cost "
        f"accounts were recognised. Check the Expense Report before relying on it."
    ]
    if info["unclassified"]:
        notes.append(
            f"{info['unclassified']} cost accounts could not be recognised and "
            f"are grouped under \"{UNCLASSIFIED_EXPENSE}\". Add an expense-type "
            "column to classify them yourself."
        )
    return notes


def infer_category(code: str) -> str:
    """The single-account fallback, kept for callers with no name to hand.

    Prefer `infer_categories`, which reads the names too and can say when it is
    guessing. This one cannot.
    """
    return {"4": "Revenue", "5": "COGS", "6": "OpEx",
            "7": "Other"}.get(str(code).strip()[:1], "OpEx")


# Debit markers across the systems that use them: SAP writes S/H (Soll/Haben),
# English exports write D/C or DR/CR, Hungarian ones T/K (tartozik/követel).
_DEBIT_MARKERS = {"d", "dr", "debit", "s", "t", "+", "1"}
_CREDIT_MARKERS = {"c", "cr", "credit", "h", "k", "-", "2"}


def _dc_sign(col: pd.Series) -> pd.Series:
    """Turn a debit/credit indicator column into +1 / -1."""
    marks = col.astype(str).str.strip().str.lower()
    return marks.map(lambda m: -1.0 if m in _CREDIT_MARKERS else 1.0)


def normalise_signs(std: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """Report every P&L line as a positive magnitude.

    A general ledger credits revenue, so a real extract carries revenue as a
    negative number while costs are positive (or the reverse, depending on the
    system). The engine expects each category as a positive magnitude and
    applies the P&L sign itself, so a whole category arriving negative is
    flipped here rather than silently producing a negative revenue line.

    Only a category that is consistently negative is flipped: a single credit
    note sitting inside an otherwise normal expense category is real data and is
    left alone.
    """
    out = std.copy()
    notes: list[str] = []
    if "category" not in out.columns:
        return out, notes

    money = [c for c in ("actual", "budget", "prior_year") if c in out.columns]
    for category, rows in out.groupby("category"):
        for col in money:
            vals = pd.to_numeric(rows[col], errors="coerce").fillna(0.0)
            nonzero = vals[vals != 0]
            if len(nonzero) < 2 or vals.sum() >= 0:
                continue
            # Consistently negative, not one odd credit among positives.
            if (nonzero < 0).mean() < 0.7:
                continue
            out.loc[rows.index, col] = -vals
            notes.append(
                f"{category} arrived as a credit balance in '{col}'; the sign was "
                "normalised so the line reports as a positive magnitude."
            )
    return out, notes


def apply_mapping(df: pd.DataFrame, matches: list[ColMatch]) -> pd.DataFrame:
    """Produce the standardised account-level frame the engine expects."""
    src = {m.field: m.source for m in matches if m.source}
    missing = [f for f in REQUIRED if f not in src]
    if not (set(src) & set(AMOUNT_FIELDS)):
        missing.append("actual (or budget)")
    if missing:
        raise ValueError(f"Missing required fields after mapping: {missing}")

    out = pd.DataFrame()
    out["account_code"] = df[src["account_code"]].astype(str).str.strip()
    out["account_name"] = df[src["account_name"]].astype(str).str.strip()

    if "actual" in src:
        actual = df[src["actual"]].map(_coerce_amount)
        # One amount column plus a per-line debit/credit marker: the column
        # holds magnitudes and the marker carries the sign.
        if "dc_indicator" in src:
            actual = actual.abs() * _dc_sign(df[src["dc_indicator"]])
    elif {"debit", "credit"} & set(src):
        # Trial-balance shape: two columns that net to the movement.
        debit = (df[src["debit"]].map(_coerce_amount).fillna(0.0)
                 if "debit" in src else 0.0)
        credit = (df[src["credit"]].map(_coerce_amount).fillna(0.0)
                  if "credit" in src else 0.0)
        actual = debit - credit
    else:
        actual = 0.0
    out["actual"] = actual

    out["budget"] = (df[src["budget"]].map(_coerce_amount)
                     if "budget" in src else 0.0)
    out["prior_year"] = (df[src["prior_year"]].map(_coerce_amount)
                         if "prior_year" in src else 0.0)
    if "category" in src:
        out["category"] = df[src["category"]].astype(str).str.strip().str.title()
        out["category"] = out["category"].replace({"Cogs": "COGS", "Opex": "OpEx"})
    else:
        # Both readings, code and name, and a record of which one was trusted -
        # kept on the frame so `ingest` can put the guess in front of the user.
        cats, info = infer_categories(out["account_code"], out["account_name"])
        out["category"] = cats
        out.attrs["category_inference"] = info
    if "expense_type" not in src:
        etypes, einfo = infer_expense_types(out["account_name"], out["category"])
        if einfo["recognised"]:
            out["expense_type"] = etypes
            out.attrs["expense_type_inference"] = einfo
    for opt in ("period", "expense_type", "department", "cost_centre", "entity"):
        if opt in src:
            col = df[src[opt]]
            out[opt] = col.where(col.notna(), "").astype(str).str.strip()
            out[opt] = out[opt].replace({"nan": "", "None": ""})

    out = out.dropna(subset=["actual", "budget"], how="all")
    out[["actual", "budget", "prior_year"]] = out[["actual", "budget", "prior_year"]].fillna(0.0)
    return out.reset_index(drop=True)


def unmapped_columns(df: pd.DataFrame, matches: list[ColMatch]) -> list[str]:
    """Source columns Flux did not map. A real ledger export carries many fields
    the report does not need (document numbers, texts, tax codes); they are
    simply ignored rather than treated as an error."""
    used = {m.source for m in matches if m.source}
    return [c for c in df.columns if c not in used]


def mapping_report(matches: list[ColMatch]) -> pd.DataFrame:
    return pd.DataFrame([{
        "field": m.field, "mapped_from": m.source or "(none)",
        "method": m.method, "confidence": m.score,
        "required": m.field in REQUIRED,
    } for m in matches])


def _header_score(values: list) -> float:
    """How much does this row look like a header row?"""
    cells = [v for v in values if v is not None and str(v).strip() != "" and str(v) != "nan"]
    if len(cells) < 2:
        return -1.0
    text = [c for c in cells if not _looks_numeric(c)]
    frac_text = len(text) / len(cells)
    known = 0
    for c in cells:
        n = _normalize(c)
        for syns in _SYNONYMS.values():
            if n in syns or any(_fuzzy(n, s) >= _FUZZY_CUTOFF for s in syns):
                known += 1
                break
    return len(cells) * 0.1 + frac_text * 2.0 + known * 1.5


def _looks_numeric(v) -> bool:
    if isinstance(v, (int, float)):
        return True
    s = str(v).strip()
    return bool(s) and not pd.isna(_coerce_amount(s))


def _promote_header(raw: pd.DataFrame, scan_rows: int = 15) -> pd.DataFrame:
    """Find the real header row when a file has title/blank rows above it."""
    best_i, best_score = 0, -1.0
    for i in range(min(scan_rows, len(raw))):
        score = _header_score(list(raw.iloc[i].values))
        if score > best_score:
            best_i, best_score = i, score
    header = [str(v).strip() if v is not None and str(v) != "nan" else f"col_{j}"
              for j, v in enumerate(raw.iloc[best_i].values)]
    out = raw.iloc[best_i + 1:].copy()
    out.columns = header
    out = out.dropna(axis=1, how="all").dropna(axis=0, how="all")
    return out.reset_index(drop=True)


def _sniff_delimiter(text: str) -> str:
    """Pick the delimiter that yields the most consistent column count.

    Sniffing on the first line alone fails when a file starts with free-text
    title rows, so score each candidate across all lines: the best delimiter is
    the one whose most common field count is highest and most frequent.
    """
    lines = [ln for ln in text.splitlines() if ln.strip()][:50]
    best, best_score = ",", (0, 0)
    for cand in [",", ";", "\t", "|"]:
        counts = {}
        for ln in lines:
            n = len(ln.split(cand))
            if n > 1:
                counts[n] = counts.get(n, 0) + 1
        if not counts:
            continue
        n, freq = max(counts.items(), key=lambda kv: (kv[1], kv[0]))
        score = (freq, n)
        if score > best_score:
            best, best_score = cand, score
    return best


def load_file(path_or_buffer, sheet_name=0) -> pd.DataFrame:
    """Read a CSV/Excel file, locating the header row even if titles sit above it."""
    name = getattr(path_or_buffer, "name", str(path_or_buffer)).lower()
    if name.endswith((".xlsx", ".xls")):
        raw = pd.read_excel(path_or_buffer, sheet_name=sheet_name, header=None)
    else:
        if hasattr(path_or_buffer, "read"):
            data = path_or_buffer.read()
            text = data.decode("utf-8-sig", errors="replace") if isinstance(data, bytes) else data
        else:
            text = Path(path_or_buffer).read_text(encoding="utf-8-sig", errors="replace")
        sep = _sniff_delimiter(text)
        ncols = max((len(ln.split(sep)) for ln in text.splitlines() if ln.strip()), default=1)
        raw = pd.read_csv(StringIO(text), sep=sep, header=None,
                          names=range(ncols), skip_blank_lines=False, engine="python")
    return _promote_header(raw)


def _llm_suggest(headers: list[str], sample: pd.DataFrame, unmatched: list[str]) -> dict:
    """Ask the model to map still-unmatched canonical fields to a source column.

    Lazy import; only reached when a key is present and fields remain. Returns
    {canonical_field: source_header}. Any failure raises to the caller, which
    ignores it (graceful degradation).
    """
    import json
    import anthropic

    client = anthropic.Anthropic()
    preview = sample.head(5).to_dict(orient="records")
    prompt = (
        "Map spreadsheet columns to a fixed set of finance fields. "
        f"Fields still needing a column: {unmatched}. "
        f"Available column headers: {headers}. "
        f"Sample rows: {preview}. "
        "Return ONLY a JSON object mapping each needed field to the single best "
        "header, or omit a field if no column fits. No prose, no code fences."
    )
    msg = client.messages.create(
        model="claude-haiku-4-5-20251001", max_tokens=300,
        messages=[{"role": "user", "content": prompt}],
    )
    text = "".join(b.text for b in msg.content if getattr(b, "type", None) == "text")
    text = text.strip().removeprefix("```json").removeprefix("```").removesuffix("```").strip()
    result = json.loads(text)
    return {k: v for k, v in result.items() if k in CANONICAL and v in headers}


def ingest(source, overrides: dict | None = None, use_llm: bool | None = None):
    """Ingest a file path or DataFrame -> (standardised_df, report_df, issues)."""
    df = source if isinstance(source, pd.DataFrame) else load_file(source)
    matches = match_columns(list(df.columns), overrides)
    _content_detect(df, matches)

    # Layer 4: optional LLM fallback for unmatched required fields.
    unmatched_req = [m.field for m in matches if m.field in REQUIRED and m.source is None]
    use_llm = HAS_LLM if use_llm is None else use_llm
    if unmatched_req and use_llm:
        try:
            used = {m.source for m in matches if m.source}
            free = [c for c in df.columns if c not in used]
            suggest = _llm_suggest(free, df, unmatched_req)
            for m in matches:
                if m.field in suggest and m.source is None:
                    m.source = suggest[m.field]; m.method = "llm"; m.score = 0.7
        except Exception:
            pass  # graceful degradation

    issues = []
    unmatched_req = [m.field for m in matches if m.field in REQUIRED and m.source is None]
    src_now = {m.field for m in matches if m.source}
    if not (src_now & set(AMOUNT_FIELDS)):
        unmatched_req = unmatched_req + ["actual (or budget)"]
    if unmatched_req:
        issues.append(f"Unmatched required fields (need manual mapping): {unmatched_req}")

    std = None
    if not unmatched_req:
        std = apply_mapping(df, matches)
        std, sign_notes = normalise_signs(std)
        issues.extend(sign_notes)
        src = {m.field: m.source for m in matches if m.source}

        if "budget" not in src and "actual" in src:
            issues.append(
                "No budget column found. This looks like an actuals-only extract, so the "
                "report will show actuals without budget variance. Add a budget column, or "
                "upload the budget as a second file."
            )
        if "actual" in src and std["actual"].sum() == 0:
            issues.append("The actual total is zero; check which column holds the amount.")

        # An unread chart of accounts produces a finished-looking pack with the
        # bottom line inverted, so the classification is reported whether it went
        # well or badly - it is the one guess the user cannot check by eye.
        cat_info = std.attrs.get("category_inference")
        if cat_info:
            issues.extend(category_issues(list(std["account_code"]),
                                          list(std["account_name"]),
                                          list(std["category"]), cat_info))
        exp_info = std.attrs.get("expense_type_inference")
        if exp_info:
            issues.extend(expense_type_issues(exp_info))

        # Multi-currency ledger mapped to a local-currency column: amounts in
        # different currencies must not be added together.
        ccy_col = next((c for c in df.columns if _normalize(c) in
                        {"currency", "ccy", "curr", "devizanem", "deviza", "penznem"}), None)
        if ccy_col is not None and df[ccy_col].nunique(dropna=True) > 1:
            chosen = _normalize(src.get("actual", ""))
            if "eur" not in chosen and "report" not in chosen:
                issues.append(
                    f"The file holds several currencies ({df[ccy_col].nunique()}). Map the "
                    "amount column that is already translated into the reporting currency, "
                    "otherwise the totals mix currencies."
                )
    return std, mapping_report(matches), issues


def aggregate_to_accounts(std: pd.DataFrame) -> pd.DataFrame:
    """Collapse posting lines while keeping the analytical dimensions.

    A raw ledger holds many lines per account. Reporting needs them summed, but
    grouping on the account alone would destroy the department, cost-centre and
    entity detail that the departmental and consolidation views rely on, so the
    dimensions present in the file are part of the grouping key.
    """
    dims = [c for c in ("period", "expense_type", "department", "cost_centre", "entity")
            if c in std.columns]
    keys = ["account_code", "account_name", "category"] + dims
    money = ["actual", "budget", "prior_year"]
    for m in money:
        if m not in std.columns:
            std = std.assign(**{m: 0.0})
    agg = std.groupby(keys, as_index=False, dropna=False)[money].sum()
    if "period" in agg.columns:
        agg["period_no"] = agg["period"].map(period_key)
    return agg.sort_values(["account_code"] + dims).reset_index(drop=True)


def period_key(period: str) -> int | None:
    """Turn a period label into a sortable YYYYMM key.

    Accepts the shapes real exports use: 2025-06, 2025/06, 202506, 06.2025,
    2025-06-30. Returns None when nothing date-like can be read.
    """
    if period is None:
        return None
    digits = re.findall(r"\d+", str(period))
    if not digits:
        return None
    joined = "".join(digits)
    if len(joined) >= 6:
        head = joined[:6]
        y, m = int(head[:4]), int(head[4:6])
        if 1 <= m <= 12 and 1900 <= y <= 2999:
            return y * 100 + m
        # month-first shapes such as 06.2025
        m2, y2 = int(joined[:2]), int(joined[2:6])
        if 1 <= m2 <= 12 and 1900 <= y2 <= 2999:
            return y2 * 100 + m2
    return None


def reporting_period(std: pd.DataFrame) -> tuple[str | None, list[str]]:
    """The period a pack should report: the latest month that has actuals.

    Finance reports the most recent closed month and shows the year to date
    beside it, so the choice follows from the data rather than from a setting.
    A full-year budget must not drag the reporting month into the future: months
    that only carry a plan are not closed yet.
    """
    found = periods_in(std)
    if not found:
        return None, []
    order = lambda p: (period_key(p) is None, period_key(p) or 0, p)
    ordered = sorted(found, key=order)

    if "actual" in std.columns:
        posted = std.loc[pd.to_numeric(std["actual"], errors="coerce").fillna(0) != 0, "period"]
        closed = sorted({str(p).strip() for p in posted if str(p).strip()}, key=order)
        if closed:
            return closed[-1], ordered
    return ordered[-1], ordered


def periods_in(std: pd.DataFrame) -> list[str]:
    """Distinct accounting periods present in an ingested frame."""
    if "period" not in std.columns:
        return []
    vals = (std["period"].dropna().astype(str).str.strip())
    return sorted(v for v in vals.unique() if v and v.lower() != "nan")


def filter_period(std: pd.DataFrame, period: str) -> pd.DataFrame:
    """Keep one accounting period. Files spanning several months must not be
    summed together into a single 'month'."""
    if "period" not in std.columns or not period:
        return std
    out = std[std["period"].astype(str).str.strip() == str(period)]
    return out.reset_index(drop=True) if len(out) else std


def merge_budget(actuals: pd.DataFrame, budget: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    """Join a separate budget file onto actuals.

    Budgets usually come from a planning system rather than the ledger, so they
    arrive as their own file. The join uses the account plus whichever
    dimensions both files share, falling back to the account alone.
    """
    notes = []
    bud = budget.copy()
    if bud["budget"].abs().sum() == 0 and bud["actual"].abs().sum() > 0:
        bud["budget"] = bud["actual"]
        notes.append("Budget file: its amount column was read as the budget.")

    dims = [c for c in ("period", "expense_type", "department", "cost_centre", "entity")
            if c in actuals.columns and c in bud.columns]
    keys = ["account_code"] + dims
    if dims:
        notes.append("Budget joined on account plus " + ", ".join(dims) + ".")

    b = bud.groupby(keys, as_index=False, dropna=False)[["budget"]].sum()
    out = actuals.drop(columns=["budget"], errors="ignore").merge(b, on=keys, how="outer")

    # The actuals may be finer than the plan (a ledger split by entity against a
    # company-level budget). A plain join would repeat the budget on every
    # matching line and inflate the total, so the plan is allocated across those
    # lines in proportion to actuals - evenly when there are no actuals to weight.
    grp = out.groupby(keys, dropna=False)
    if grp.size().max() > 1:
        weight = out["actual"].abs()
        total = grp["actual"].transform(lambda x: x.abs().sum())
        count = grp["actual"].transform("size")
        share = (weight / total).where(total > 0, 1.0 / count)
        out["budget"] = out["budget"] * share
        notes.append("Budget allocated across finer actual lines in proportion to actuals.")

    # Budget-only lines (a plan for an account with no postings yet) arrive from
    # the outer join with the descriptive columns empty; fill them so the report
    # can still classify and label the line.
    for col in ("actual", "budget", "prior_year"):
        if col not in out.columns:
            out[col] = 0.0
        out[col] = pd.to_numeric(out[col], errors="coerce").fillna(0.0)
    for col in ["account_name", "category"] + dims:
        if col in out.columns:
            out[col] = out[col].where(out[col].notna(), "").astype(str).str.strip()
            out[col] = out[col].replace({"nan": "", "None": ""})
    if "category" in out.columns:
        blank = out["category"].eq("")
        if blank.any():
            out.loc[blank, "category"] = out.loc[blank, "account_code"].map(infer_category)
    if "account_name" in out.columns:
        blank = out["account_name"].eq("")
        out.loc[blank, "account_name"] = out.loc[blank, "account_code"].astype(str)
    if "period" in out.columns:
        out["period_no"] = out["period"].map(period_key)

    # A year-to-date view needs a plan that covers the same months.
    if "period" in actuals.columns and "period" in bud.columns:
        a_per, b_per = set(periods_in(actuals)), set(periods_in(bud))
        gap = sorted(a_per - b_per)
        if gap:
            notes.append(
                f"Budget covers {len(b_per)} of the {len(a_per)} periods in the actuals; "
                f"year-to-date budget is incomplete (missing {', '.join(gap[:4])}"
                + (" …" if len(gap) > 4 else "") + ")."
            )

    only_bud = int((out["actual"] == 0).sum())
    only_act = int((out["budget"] == 0).sum())
    if only_bud:
        notes.append(f"{only_bud} line(s) budgeted with no actual postings.")
    if only_act:
        notes.append(f"{only_act} line(s) with actuals but no budget.")
    return out.reset_index(drop=True), notes


TEMPLATE_COLUMNS = ["account_code", "account_name", "category", "cost_centre",
                    "actual", "budget", "prior_year"]


def write_template_csv(path: str) -> None:
    """Minimal CSV template. The full styled Excel template is built by
    scripts/make_template.py into data/input_template.xlsx."""
    sample = pd.DataFrame([
        {"account_code": "4100", "account_name": "Subscription revenue", "category": "Revenue",
         "cost_centre": "Commercial", "actual": 512000, "budget": 480000, "prior_year": 445000},
        {"account_code": "5000", "account_name": "Direct materials", "category": "COGS",
         "cost_centre": "Operations", "actual": 168000, "budget": 160000, "prior_year": 150000},
        {"account_code": "6110", "account_name": "Advertising & digital", "category": "OpEx",
         "cost_centre": "Marketing", "actual": 128000, "budget": 90000, "prior_year": 82000},
    ], columns=TEMPLATE_COLUMNS)
    sample.to_csv(path, index=False)


if __name__ == "__main__":
    # Simulate a messy client export: odd headers, EU number formatting, no
    # category, and negatives written the way real systems write them.
    messy = pd.DataFrame({
        "GL Code": ["4100", "4000", "5000", "6110", "6300"],
        "Description": ["Subscription rev", "Product rev", "Materials",
                        "Digital ads", "G&A salaries"],
        "Actual Amount (EUR)": ["512.000,00-", "298.000,00-", "168.000,00",
                                "128.000,00", "88.000,00"],
        "Plan": ["480.000,00-", "340.000,00-", "160.000,00", "90.000,00", "90.000,00"],
        "Last Year": ["(445.000)", "(331.000)", "150.000", "82.000", "86.000"],
        "Dept": ["Commercial", "Commercial", "Operations", "Marketing", "G&A"],
    })
    std, report, issues = ingest(messy)
    print("MAPPING REPORT")
    print(report.to_string(index=False))
    print("\nISSUES:", issues or "none")
    print("\nSTANDARDISED (fed to the engine):")
    print(std.to_string(index=False))

    from .engine import build_report
    print("\nP&L from ingested data:")
    rep = build_report(std)
    print(rep[["line", "actual", "budget", "var_bud", "fav_unfav", "material"]]
          .to_string(index=False))

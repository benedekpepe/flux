<img src="assets/flux-logo.svg" alt="Flux" width="220">

# Flux — FP&A reporting automation

[![CI](https://github.com/benedekpepe/flux/actions/workflows/ci.yml/badge.svg)](https://github.com/benedekpepe/flux/actions/workflows/ci.yml)

**Live demo: https://flux-reporting.streamlit.app/** — open the *Use sample data*
tab to generate the full pack with no sign-up, or upload your own ledger.

A management reporting engine that turns a general ledger into a **management
P&L with variance analysis** — budget and prior-year variances, favourable /
unfavourable classification, materiality flagging, and consolidation by legal
entity, department and cost centre — delivered as a **live Excel pack**, a
**PowerPoint management deck** and a written commentary, behind a Streamlit app.

The name refers to *flux analysis*: the FP&A practice of explaining the movement
between actual results and budget or prior periods. That is the work this engine
automates.

> **Portfolio project.** Built from a controlling background to show the
> reporting logic, not to replace a consolidation system. The demo company and
> its ledger are generated; no client data is included or required.

## Two ways in

- **The input template** *(recommended)* — one row per account, columns already
  named. The same engine runs either way and the numbers are identical where the
  inference is right; what the template removes is the inference itself. Nothing
  to map, no chart of accounts to recognise, no expense types read off account
  names, and no printed subtotals to filter out.
- **Upload your own export** — a GL extract or trial balance, optionally with a
  separate budget file. Flux reads it as it is: it maps the columns, classifies
  the accounts, drops the printed furniture, and **shows you every assumption it
  made** so you can correct any of them before the pack is built.
- **Sample company** — three legal entities in EUR, USD and HUF, transaction-level
  actuals and a full-year budget. Generates the complete pack with no upload.

The upload path is the interesting one and it is where most of the work went;
the template is what to use when the report has to be right the first time and
nobody has time to check four assumptions.

## What it does

- **Rolls up** leaf accounts to category subtotals and computed subtotals —
  Gross profit, EBIT, Net income — from a single source table.
- **Variances** against budget and prior year, in absolute and percentage terms.
- **Favourable / unfavourable by account type** — revenue and profit lines are
  "higher is better", cost lines are "lower is better", so an overspend never
  reads as good news because the number went up.
- **Two-condition materiality, per timeframe** — an item is flagged only when it
  clears both an absolute EUR floor and a percentage floor, so a 300% variance on
  a €4k line doesn't crowd out the real movers. A near-zero base reads **n/m** and
  is judged on the amount instead. The flag names *which* timeframe cleared both
  floors — **MONTH**, **YTD** or **BOTH** — because a one-off overspend and a
  pattern that has been building all year need different answers.
- **One layout on every sheet** — the month, the year to date, and the full-year
  plan with the run rate against it. Each of the three horizons gets the same
  four columns (actual, plan, variance, variance %), then F/U and the flag. The sheets
  differ in what they cut the ledger by, never in how they report it, so a reader
  who has learnt the P&L has learnt the pack.
- **Written commentary** on every roll-up row and for the pack as a whole,
  generated from the actual drivers. Each comment covers the year to date *and*
  the month and says which of them clears the materiality floors — so it
  explains the flag beside it rather than restating a column.
- **Reads real ledger files** — accounting sign conventions, EU and US number
  formats, English and Hungarian headers, debit/credit columns.

## Screenshots

**Management P&L — the month headlined, the year to date beneath it, and the pack's three lever cells**

![Management P&L with month and YTD variances, the run rate against the full-year plan, F/U badges, materiality flags and per-line commentary](docs/pnl-report.png)

**Expense Report — the same columns, cut by expense type**

![Expense report by expense type on the shared column layout](docs/expense-report.png)

**By Entity — net income per legal entity, consolidating to the group P&L**

![Net income by legal entity consolidating to the group total](docs/by-entity.png)

**Departments & Cost Centres — spend roll-up with cost centres nested**

![Department spend variance with cost centres nested beneath each department](docs/departments.png)

**PowerPoint deck — the same numbers as seven slides, for sending upward**

![Four slides of the generated management deck: cover, result, P&L and year to date](docs/pptx-pack.png)

## How it works

1. **Ingestion** — an arbitrary export is mapped onto the internal schema in
   layers: a synonym dictionary (English + Hungarian), fuzzy header matching,
   content detection from the data itself, and an optional LLM fallback for the
   hard leftovers. Accounts are then classified by category and expense type
   from the codes and the names together. Both the mapping and the
   classification are shown for approval and can be overridden.
2. **Normalisation** — amounts are parsed across accounting sign conventions,
   posting lines are aggregated to account level keeping the analytical
   dimensions, and a credit-balance category is flipped to a positive magnitude.
3. **Engine** — the P&L structure is applied, variances computed, F/U resolved
   from account type, and materiality flagged against the two floors.
4. **Commentary** — the material drivers are turned into a narrative: a headline
   on the bottom line, revenue and gross profit, then the accounts that moved.
5. **Excel pack** — the workbook is written with **live formulas over an input
   sheet**, so an edited input recalculates the whole pack. The sheets that get
   built depend on the dimensions present in the data.
6. **Deck** — the same numbers as seven slides: the result, the P&L, the year to
   date, where the year lands, what moved it, and where the money went. Native PowerPoint charts and
   tables, not images, so the recipient can edit them. The cumulative slide
   compares the month against the average month so far, which is the difference
   between a bad month and a trend; a single-period file drops that slide rather
   than repeat the month under a cumulative heading.
7. **App** — a Streamlit front end for upload, mapping confirmation, preview and
   download of either output.

### The materiality rule

Both floors must be cleared:

| Variance | Amount | Percentage | Material? |
| --- | --- | --- | --- |
| Large amount, small share of a big budget | €30,000 | 3% | No |
| Large share of a tiny budget | €6,000 | 200% | No |
| Clears both | €130,000 | 30% | Yes |
| Near-zero budget | €150,000 | n/m | Yes — judged on the amount |

The test runs twice on every line, once on the month and once on the year to
date, and the flag says which one cleared:

| Flag | Reads as |
| --- | --- |
| **MONTH** | it moved this month, but the year to date is still on plan |
| **YTD** | the month looks fine; the gap has been building since January |
| **BOTH** | material this month *and* cumulatively — look here first |
| *(blank)* | not material on either timeframe |

A single MATERIAL badge could not tell those apart, and the difference is the
difference between a bad month and a trend.

### The lever cells

Three assumptions sit as **editable cells on the P&L sheet**, and every other
sheet points at them rather than holding its own copy:

| Cell | Lever |
| --- | --- |
| `C9` | materiality floor, EUR |
| `G9` | materiality floor, % |
| `K9` | months elapsed, which drives every run rate in the pack |

Months elapsed is **counted from the data** — the number of months that actually
carry postings up to the reporting month — not read off the reporting month's
number. An extract that starts in March still reports June as month six, and
dividing a four-month year to date by six would understate every projection in
the pack by a third. It stays editable, because only the reader knows whether a
gap is a quiet month or a missing export.

Every sheet also carries the same line under the masthead — `Month 2025-06 · YTD
through 2025-06 · FY 2025 · €` — because every sheet now reports the same three
horizons. "YTD 2025-06" on its own would not say whether that is the year through
June or the month of June, so the window is named.

The two floors and the month count are echoed **read-only** on every other
sheet, under the masthead, so a reader who sees `BOTH` on the expense report can
see what threshold produced it without leaving the sheet. They are formulas
pointing at the P&L, not a second editable copy: two editable copies of one
assumption is two sources of truth, which is the problem the levers exist to
solve.

The KPI cards headline the **year to date**, with the month muted beneath it. A
single month is noisy; cumulative performance against the annual plan is what a
management P&L cover is read for. Both figures are in the card, so nobody has to
choose. The green or red on the delta comes from a conditional-formatting rule
rather than a colour fixed at build time — the workbook is live, and a colour
baked in at build time survived an edit that moved the number underneath it.

Change one number and the whole pack re-flags or re-projects. That is enforced by
the test suite, which reads the saved workbook back and checks that each sheet's
flag and run-rate formulas reference the P&L's cells.

## Where the year lands

Every reporting sheet carries the full-year plan and the **run rate** against it:
the year to date ÷ months elapsed × 12, and the gap between that and the budget.
It used to live on a separate *Outlook* sheet, which meant the one question a
reader asks after "how is the month?" was two clicks away from every cut of the
data. Now every line answers it in place.

Months elapsed is the `K9` lever on the P&L, so a reader who disagrees with the
assumption changes one cell and watches the whole pack move.

The deck goes further and shows a second projection beside the run rate:

| Projection | Assumption |
| --- | --- |
| **Run rate** | the rest of the year behaves like the year so far — year to date ÷ months elapsed × 12 |
| **Plan for the rest** | the remaining months hit budget — actual to date + the unspent part of the plan |

The outcome usually sits between them, and the gap is itself the message: it is
the size of what the remaining months have to make up. The workbook carries the
run rate only, because it repeats it on every line of every sheet and a second
projected column on all of them would cost more width than it earns.

Neither is a forecast. There is no seasonality, no pipeline and no assumed
management action — they are arithmetic under a stated assumption, which is what
makes them arguable rather than authoritative.

## Actuals with no budget

A general ledger holds actuals; the plan usually lives in another system. Flux
accepts either shape — one file with both, or a ledger plus a budget export
joined on account code.

**If only actuals are supplied, the variance, F/U and materiality columns are
left empty rather than measured against zero**, and the pack says so. A ledger
compared against nothing would otherwise report every line as a 100% overspend
and flag all of them. The run rate is the exception: it is built from actuals and
the months lever alone, so it still projects. Prior-year actuals are carried on
the GL Input sheet, beside the figures they belong to.

## A real export is a printed report, not a table

An export carries a title block, blank lines, a subtotal after each group, a
grand total at the bottom and a "printed on" footer. Read as data, **the subtotal
rows count the whole ledger twice** — quietly, because the pack still builds and
still looks finished.

So rows that name a total rather than an account are left out, and so are rows
with no amount in any column. Both are reported: a row silently dropped changes
the totals as surely as a subtotal silently counted.

The filter wants two signals before discarding a row that has a real account
code, because a chart of accounts legitimately holds *Net interest expense* and
*Gross margin adjustment*. A row with no code needs only the word.

## Reading a chart of accounts it has never seen

The account code alone cannot say what a line is. `5` opens material cost in a
US chart, every cost by nature in a Hungarian one, and an expense account in
SAP; `8` is revenue in SAP and a cost in Hungary. Reading the first digit and
hoping reported Hungarian revenue as an operating expense and swapped SAP's
revenue and materials — a finished-looking pack with the bottom line inverted,
which is the worst way to be wrong.

So two independent readings are taken and made to agree:

| Reading | From |
| --- | --- |
| **Account name** | keywords in English, Hungarian and German, deliberately narrow — it returns nothing when unsure |
| **Account code** | candidate chart styles: `4-5-6-7`, SAP, SKR03, Hungarian statutory |

Each chart style is scored on how often it agrees with the names. The winner has
to clear a floor, so an unrecognised chart falls back to the names rather than
being forced into a shape it is not. Whichever way it goes, **the pack says what
it based the reading on** — the chart it recognised and how much of the file
corroborates it, every account where the name and the code disagreed, and the
loud cases: no revenue at all, or an account whose name reads as revenue and was
classified otherwise.

A confident, fully corroborated match says nothing, because a warning that always
fires is a warning nobody reads.

The score is load-bearing, not decorative: SAP and SKR03 both put revenue in the
8000s and disagree about the 4000s, and only the account names can tell them
apart. Where a file's names carry no signal at all, the widest-covering chart is
used and the pack says outright that nothing corroborates it.

**The classification is confirmed in the app the same way the column mapping is** —
an editable table of account, name, category and expense type, with the warnings
above it. It is the second guess in the pipeline, and it was the one nobody was
being shown.

### Expense types, where the file has none

The expense report needs a natural classification the ledger usually does not
carry as a column. Where there is none, it is inferred from the account names
against the same seventeen types the demo uses, so an inferred file groups
exactly like a native one. Anything unrecognised lands under *Other expenses* —
named, not blank, so the line still appears on a spend view rather than
vanishing from it. The pack reports how many accounts were recognised and how
many were not.

## Reading real ledger files

A real export writes a negative in at least four ways, and getting any of them
wrong either flips the sign or drops the row while the report still looks
finished:

| Written as | Convention |
| --- | --- |
| `-1.234,50` | leading minus |
| `(1.234,50)` | accounting parentheses |
| `1.234,50-` | trailing minus — the SAP default |
| `1,234.50 CR` | credit marker |

EU and US separators are resolved from the rightmost separator, and currency
symbols, currency codes and non-breaking spaces are stripped.

Sign conventions are handled at three levels: **separate debit and credit
columns** are netted; **one amount column plus a D/C marker** is signed from the
marker (SAP `S`/`H`, English `D`/`C` and `DR`/`CR`, Hungarian `T`/`K`); and **a
whole category arriving as a credit balance** — a ledger credits revenue, so
revenue often arrives negative — is normalised and reported. A single credit note
inside an otherwise positive category is real data and is left alone.

Monetary columns are never guessed from an unnamed column: a document number or
a fiscal year must not be read as an amount, so Flux asks rather than assumes.

## The pack adapts to your data

| Sheet | Needs |
| --- | --- |
| P&L Report, Drivers, GL Input | always |
| Expense Report | an expense-type column |
| Departments & Cost Centres | a department column (cost centre nests under it) |
| By Entity | more than one entity |

The department-to-cost-centre hierarchy is read from the uploaded data, not from
a fixed chart of accounts, so a client's own structure is respected. A period
column keeps months apart: the reporting month is the **latest period that
carries postings**, so a full-year budget beside six months of actuals still
reports June. Whichever sheets get built, they all carry the same columns; a
file holding a single period says so on the P&L rather than leaving the reader
to work out why the month and the year to date are identical.

## Analysis

Under every reporting sheet, and on its own slide in the deck, is a short
analysis of the lines that carry a flag. Four findings, in the order a reader
wants them:

| Finding | Answers |
| --- | --- |
| **Concentration** | is it one line or forty — and which ones |
| **Persistence** | how many of the months so far went the wrong way, and whether the gap is widening |
| **Full year** | where the run rate lands against the plan, and what the remaining months would have to do |
| **Ask** | the question the shape of the variance makes worth putting to someone |

> **Other operating costs · BOTH**
> **Concentration** — Effectively all of the movement sits in 3 lines: marketing
> & advertising (€47.0k), IT & software (€319) and facilities & office (€149).
> **Persistence** — Adverse in 6 of 6 months and the gap has been widening.
> **Full year** — At the current run rate the year lands €511.2k over a €3.5m
> plan. Holding the plan leaves €1.5m for the remaining 6 months, against
> €333.4k a month so far.
> **Ask** — Adverse in 6 of 6 months, which is a level rather than an event: the
> plan and the actual activity disagree. Worth asking the budget owner whether
> the plan was set before the current activity level.

### What it will not do

**It does not say why.** A ledger does not carry causes — a line text in a real
extract reads `Invoice 88213` or `Reclass to 5211`, not "campaign overspend" —
so any sentence beginning *because* would be invented, and an invented cause in a
management pack is the one thing the reader cannot check against the numbers
beside it.

The **Ask** finding is the honest form of "what should I do". A single month out
of line and a six-month drift are the same variance and a different problem: one
is a timing question for the accountant (cut-off, an accrual released, an invoice
in the wrong period), the other a planning question for the budget owner (the
plan and the activity level disagree, and one of them has to move). Which it is
follows from the data, so the pack says it — and names who is likely to hold the
answer rather than guessing at it.

Blocks are written only for the lines the flag already picked out. One under
every row would bury the two that matter. The test suite asserts that no cell in
the pack contains a causal phrasing.

## Commentary

Every comment has the same shape: the year to date, then the month, then the
verdict, then what moved it.

> YTD €255.7k (21.8%) above budget, unfavourable. Month €45.8k (22.1%) above
> budget. Material on both timeframes. Driven by brand & content, demand
> generation and events.

The year to date leads because it is the trend; the month follows because it is
the latest point. Reversing them makes every comment read as news even when the
line has been drifting since January.

Commentary is written on **roll-up rows only** — a P&L line over its accounts, a
department over its cost centres, an expense group over its types, an entity over
its departments. There it can name *which* of the lines below moved the total,
which is not on the row. On a leaf row the comment could only restate the
variance and percentage in the columns beside it, so the Drivers sheet and the
cost-centre lines carry none.

A group holding a single line counts as a roll-up too — it is drawn as the group
row, so a blank comment beside its filled neighbours reads as a gap. Only the
"driven by" clause is dropped there, because it would name the row itself.

Grand totals count as roll-ups: they have the whole sheet underneath them and
are the row a reader looks at first. A spend sheet's total comment is built on a
revenue-free frame, matching the figures on the row — otherwise it agrees with
the individual rows by luck and contradicts the total.

Every roll-up gets one, not only the material ones: a blank cell is ambiguous —
nothing moved, or nothing was generated? — and "neither timeframe clears the
floors" is an answer. The Flag column is what points a reader at the rows worth
stopping on; the commentary explains them.

- **Template mode** (default) — pure Python, no API key, no cost. Describes
  magnitude and direction faithfully; it does not invent root causes.
- **LLM mode** (optional) — with `ANTHROPIC_API_KEY` set, the template narrative
  is enriched via the Anthropic API. Any missing key or failure falls back to the
  template, so a public demo never depends on a paid key.

## Tech stack

Python · pandas · NumPy · openpyxl · python-pptx · Streamlit. No heavyweight dependencies: the
reporting logic is plain and fast, and the Excel layer writes formulas rather
than values.

## Project structure

```
flux/
├── app.py                       # Streamlit app: upload -> map -> download
├── src/flux/
│   ├── coa.py                   # chart of accounts, org hierarchy, P&L structure
│   ├── engine.py                # variance engine: P&L, drivers, departments, entities
│   ├── commentary.py            # narrative generation (optional LLM enrichment)
│   ├── analysis.py              # findings: concentration, persistence, outlook, the question
│   ├── ingest.py                # arbitrary client files -> the internal schema
│   ├── synthetic_data.py        # seeded generator for the demo company
│   └── reporting/
│       ├── styling.py           # palette, type, number formats, sheet chrome
│       ├── formulas.py          # formula builders, lever cells, column layout
│       ├── rows.py              # the report row every sheet writes the same way
│       ├── demo_pack.py         # the multi-entity showcase, from generated data
│       ├── client_pack.py       # the pack built from an ingested client file
│       └── pptx_pack.py         # the seven-slide management deck
├── scripts/
│   ├── demo.py                  # console P&L, drivers and commentary
│   └── make_template.py         # builds data/input_template.xlsx
├── tests/audit.py               # end-to-end verification
├── assets/                      # logo and mark (SVG)
├── data/                        # input template
├── docs/                        # screenshots used in this README
└── .github/workflows/ci.yml     # lint, verification, pack build, app smoke test
```

`styling`, `formulas` and `rows` are shared by both packs, so the demo and the
client pack render the same visual language and reach the same materiality
verdict without either importing the other. A sheet supplies only the five
figures it alone can know — the month actual and budget, the year-to-date pair
and the full-year plan — and `rows` writes everything that follows. That is what
stops the two packs reporting the same numbers two different ways, which is
exactly what they had started doing.

The demo chart of accounts covers 41 accounts, 17 expense types, 7 departments,
19 cost centres, 3 legal entities and 3 currencies.

## Getting started

### Prerequisites

- **Python 3.10+**
- No database, no API key, no configuration — the demo generates its own data.

### 1. Create a virtual environment and install dependencies

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

```powershell
# Windows (PowerShell)
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

If PowerShell refuses to run the activation script, allow it for the current
terminal only: `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`.

### 2. Check the install

```bash
python tests/audit.py
```

`PASSED all 271 checks` means everything is wired up correctly.

### 3. Run the app

```bash
streamlit run app.py
```

Open the **Use sample data** tab to see the full pack immediately, or upload your
own GL on the first tab. The hosted version at
[flux-reporting.streamlit.app](https://flux-reporting.streamlit.app/) runs the
same code if you would rather not install anything.

The app follows your system light/dark preference. Both palettes are built from
the same navy, brass and ivory the Excel packs use, so the app and the workbook
read as one product.

## Command line

The engine runs without the app. These entry points add `src` to the import
path themselves, so they work as they are:

```bash
python scripts/demo.py             # console P&L + material drivers
python scripts/make_template.py    # rebuild data/input_template.xlsx
python tests/audit.py              # the verification suite
```

Building a pack straight from the module needs `src` on the import path, since
the package lives under `src/flux/`:

```bash
PYTHONPATH=src python -m flux.reporting.demo_pack      # -> output/flux_demo_pack.xlsx
PYTHONPATH=src python -m flux.reporting.client_pack    # -> output/flux_client_pack.xlsx
PYTHONPATH=src python -m flux.reporting.pptx_pack      # -> output/flux_management_pack.pptx
```

```powershell
# Windows (PowerShell) — set it once per terminal session
$env:PYTHONPATH="src"
python -m flux.reporting.demo_pack
python -m flux.reporting.client_pack
```

Or use it as a library:

```python
from flux import build_report, generate_commentary
from flux.reporting import build_client_pack

report = build_report(gl)                    # gl: account-level DataFrame
print(generate_commentary(report, gl))
build_client_pack(gl, "2025-06", "pack.xlsx")
```

## Tests

```bash
python tests/audit.py
```

**271 checks** covering the P&L arithmetic and roll-up, F/U logic per account
type, the two-condition materiality rule and the not-meaningful escape, number
and period parsing (including all four negative conventions), sign
normalisation, column mapping (including the guard that stops a document number
being read as an amount), aggregation and the proportional budget join,
reporting-period derivation, expense grouping, cross-view reconciliation, the
structure of both generated workbooks and of the deck, the shared column layout,
the actuals-only contract, and edge cases — including a guard that the modules
defining a dataclass still import on a host that execs them without registering
them in `sys.modules`. Exits non-zero on any failure.

Two groups are the load-bearing ones. The **reconciliation** checks: the entity,
department, cost-centre and expense-type views must all tie to the same total as
the P&L, and the P&L must tie to the source transactions. The **layout** checks
read the saved workbooks back and assert that every reporting sheet in both packs
carries the same thirteen columns in the same order, and that its flag and
run-rate formulas point at the P&L's lever cells rather than a local copy — the
drift they catch is invisible in the source, because each sheet builds its own
formulas and any one of them can wander off alone.

They also run automatically on every push and pull request via GitHub Actions
(Python 3.11 and 3.12), together with the linter, an import smoke test, an
end-to-end build of all three packs, and a headless render of the Streamlit app — see
the CI badge at the top. The generated packs are uploaded as a build artifact, so
every green run leaves a downloadable workbook.

## Notes and limitations

- **Not a consolidation system.** There is no intercompany elimination, no
  currency translation adjustment and no journal posting. It reports what the
  ledger says.
- **The commentary describes, it does not diagnose.** It names what moved and by
  how much; it does not claim to know why, because the ledger doesn't say.
- **FX** is applied at the rate carried on each transaction, so the demo shows
  translated EUR alongside local currency but does not model revaluation.
- The synthetic company is seeded and deterministic, which makes the demo stable
  and the tests reproducible — but its variances are generated, not real.

## Roadmap

- [x] Core variance engine with roll-up, F/U logic and materiality flagging.
- [x] Excel export with live formulas and conditional formatting.
- [x] Variance commentary, with optional LLM enrichment.
- [x] Ingestion layer: smart column mapping for arbitrary client files.
- [x] Streamlit app: upload, confirm mapping, download the pack.
- [x] Accounting sign conventions and the actuals-only path.
- [x] Verification suite and CI.
- [x] Public deploy (Streamlit Community Cloud).
- [x] PowerPoint management pack (python-pptx).
- [x] Year-to-date view and a run-rate outlook, in the workbook and the deck.
- [x] One column layout across every reporting sheet, off shared lever cells.
- [ ] Period-over-period trend view, from the periods already in the file.

## License

MIT — see [LICENSE](LICENSE).

## Author

Built by **Péter Benedek** — B.P. Studio
[Portfolio](https://benedekpeter.netlify.app/) · [GitHub](https://github.com/benedekpepe) · [LinkedIn](https://www.linkedin.com/in/benedek-d-peter/)

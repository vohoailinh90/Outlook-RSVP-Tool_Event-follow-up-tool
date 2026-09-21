# Target structure and the staged plan to reach it

What the repository looks like now, what it should look like, and the order of the moves
in between. Read this before proposing a structural change.

## Where the code is today

Measured, not estimated:

| | |
|---|---|
| Files | 4 Python files, 7,806 lines |
| `rsvp_app.py` | 5,779 lines; `RSVPApp(tk.Tk)` is **4,673 of them, in 108 methods** |
| Tkinter-free code already at module level in `rsvp_app.py` | **766 lines, 23 functions** |
| Logic-only methods inside the class | 21 methods, 275 lines |
| `except` blocks | ~108 repo-wide, 10 ending in a bare `pass` |

Importability, which is what decides testability:

| Module | Imports on Linux? | Needs |
|---|---|---|
| `db.py` | yes | stdlib `sqlite3` at module level; `openpyxl` lazily, inside the migration |
| `history.py` | no | `openpyxl` |
| `outlook_com.py` | **no** | `pythoncom` — Windows + a signed-in Outlook |
| `rsvp_app.py` | **no** | `tkinter`, and `outlook_com` transitively |

The last row is the whole problem. Those 766 platform-neutral lines are untestable not
because of anything they do, but because line 54 of the file they live in says
`import outlook_com`.

## Target layering

```
rsvp/
  domain/    money, roster, vote resolution, amount parsing   deps: none
  i18n/      message/subject/body builders                    deps: none
  storage/   sqlite persistence                               deps: stdlib sqlite3
  ports/     OutlookPort (Protocol)                           deps: none
  adapters/  outlook COM implementation of OutlookPort        deps: pywin32  [Windows]
  export/    openpyxl workbook building                       deps: openpyxl
  ui/        tkinter, one module per tab                      deps: tkinter
app.py
```

Rule: `domain/`, `i18n/`, `storage/` and `ports/` import nothing that needs a display,
Outlook, or openpyxl. `adapters/` depends on `ports/`, never the reverse.

`scripts/check_layering.py` enforces this, and it is deliberately live **before** the code
it guards exists — it covers `db.py` and `scripts/` today and gains each package as that
package lands. A rule that only describes a future state enforces nothing.

One honest caveat, because the guard's own message says so: `db.py` today is not purely
`storage/`. Its one-time Excel migration needs `openpyxl`, imported inside the function.
That import is declared in the guard's `LAZY_ALLOWED` with a reason and an expiry — Phase 2
moves the migration into `export/`, and the entry must be deleted then. Deferring an import
keeps the module importable; it does **not** mean the layer has no dependency, and the
guard refuses any *undeclared* lazy import precisely so that "move it inside a function"
cannot become a way around the boundary.

## Phases

Each phase is independently shippable and leaves CI green.

**Phase 0 — guardrails (done).** `.gitignore`, personal data untracked, three guards with
negative tests, dependency manifests, CI on `windows-latest`, the characterization harness,
the agent roster, this document. No application code changed.

**Phase 1 — extract `i18n/` (done).** 797 lines left `rsvp_app.py` (5,779 → 4,995) for
`rsvp/i18n/`: `langs.py`, `prompts.py`, `cleanup.py`, `messages.py`. The code was moved by
exact line range rather than retyped, so the function bodies are byte-identical.

The scope was the *message* code, not all 766 lines that happen to be Tkinter-free:
`parse_amount_from_text` is money and `read_gift_contribution_rows` needs openpyxl, so both
stayed behind for Phase 2. i18n is a layer, not a bucket for whatever compiles without a
display.

`rsvp_app.py` re-exports all 34 public names, so none of the 108 methods that call them
changed. That is a compatibility shim, not the destination: new code should import
`rsvp.i18n` directly.

Proof the move changed nothing: `tests/golden/i18n_snapshot.json` holds the rendered output
of 152 builder calls — every builder × every language, plus two fallback codes and eight
paste-cleanup inputs — captured *before* the extraction. `tests/test_i18n_parity.py` renders
them again and compares. Both paths reproduce the pre-extraction hash exactly. A
property-based test can pass while a transposed line changes what a recipient receives; a
full-output comparison cannot.

The payoff is concrete: `tests/test_message_builders.py` used to import `rsvp_app`, so it
skipped on any machine without a display and Outlook — a skipping test protects nothing.
It now imports `rsvp.i18n` and runs everywhere. CI stays on Windows regardless.

**Phase 2 — extract `domain/` (done).** `rsvp/domain/` now holds `money.py` (amount
parsing plus the arithmetic behind the totals) and `roster.py` (group expansion and
de-duplication, with the Outlook expander injected rather than imported). The Excel
migration moved from `db.py` to `rsvp/export/legacy_excel.py`, so `db.py` is stdlib-only
(581 → 406 lines) and `check_layering.py`'s `LAZY_ALLOWED` is empty rather than
grandfathering an exemption forever.

The suite no longer skips anything: 72 tests run identically on Linux and on Windows CI,
where before phase 1 six of them skipped off Windows.

**Money behaviour changed here, deliberately.** `parse_amount_from_text` took the first
number in the string. Every case below silently produced a wrong figure that the UI then
displayed with full confidence:

| Input | Was | Now | Why it was wrong |
|---|---|---|---|
| `2026 year-end party, 3000 JPY` | `2026.0` | `3000.0` | took the year |
| `5 people x 3000 JPY` | `5.0` | `3000.0` | took the headcount |
| `3.000` | `3.0` | `3000.0` | read a Vietnamese/European thousands separator as a decimal point — a thousand-fold understatement |
| `1.234.567` | `0.0` | `1234567.0` | `float()` raised and the handler returned zero |
| `approx 3000-4000 JPY` | `3000.0` | `4000.0` | ambiguous range; now follows the stated rule |

The rule is now written down rather than emergent: **a number next to a currency marker
wins over one that is not**, and with no marker anywhere the first number is still used, so
a plain `3000` behaves exactly as before. `tests/test_domain_money.py` splits the cases into
unchanged, deliberately changed, and still-wrong-on-purpose (negatives lose their sign;
amounts are floats, not `Decimal`).

U4 is fixed too: `_on_amount_paid_changed` swallowed every exception, so a locked or
unwritable database lost the figure while the UI still showed it as entered. It now warns —
once per run, because it fires on every keystroke.

**Phase 3 — `ports/` + the adapter seam.** Define `OutlookPort`, make `outlook_com.py`
implement it, and give the tests a fake. This is what finally makes code above the seam
testable without Outlook. Judge it by whether a test can exercise a send path against the
fake — a directory rename that leaves everything still calling `outlook_com` directly has
bought nothing.

**Phase 4 — split `ui/`.** Seven tabs, seven modules, out of the 4,673-line class. Largest
and last, because it is worth least until the layers beneath it are real.

### What the Phase 2 review found

An independent review of the money change, plus an exhaustive old-vs-new diff over 1,998
generated budget strings, turned up one genuine defect and one overstatement:

- **`0.500` became `500.0`** — wrong by 1000x *upward*, a worse failure than the
  understatement the thousands rule exists to fix. A leading zero before the separator is a
  decimal signal; nobody writes `0.500` for five hundred. Fixed.
- **The rule is locale-blind**, so `$3.500` reads as three thousand five hundred. That is
  correct for JPY and VND, the two currencies this tool is used for, and wrong for a USD
  amount written with a trailing zero. Making it currency-dependent would need a currency to
  be present, and the commonest input of all (`3000`) has none — the ambiguity would just
  move somewhere less visible. Kept, and pinned by `TestKnownAmbiguousCases` so it is a
  decision on the record rather than an accident.
- **The docstring overstated the fix**, claiming the first number is used when no marker
  appears "anywhere". Adjacency is required, so `JPY quota is 10 max, paid 3000` yields
  `10.0`. Corrected; widening adjacency is deliberately not done, since a marker in one
  clause would then capture a number from another.

The diff itself is now `tests/golden/money_snapshot.json`: 2,260 generated inputs and their
parsed values. Any future change to amount parsing surfaces as a concrete list of figures
that would be shown differently, rather than as a surprise on someone's screen.

Worth noting how the defect was found. The review agent was originally asked to do the
old-vs-new diff, which was the wrong instruction: comparing two implementations over a
corpus is a computation, and this repository's own rule is that computable checks belong in
a script. Run as a script it took one command and found the `0.500` case immediately. The
agent's time was better spent on what a script cannot settle — whether the ambiguity is
acceptable, and whether a modal dialog can re-enter the Tk event loop.

### What the Codex review taught the tests

An external review of phase 1 found four real defects, three of them in the *evidence*
rather than in the code — the tests were claiming more than they checked:

- **The golden file was not reproducible.** It and its harness both first appear in the
  extraction commit, and the harness imports the extracted package, so regenerating it from
  the new code would have recorded a regression as the expected result. Fixed by
  `scripts/verify_golden_baseline.py`, which reconstructs the builders from `rsvp_app.py`
  at the parent commit, straight from git, and re-derives every entry. All 162 reproduce
  exactly. The fixture is now checkable rather than assertable.
- **A truncated golden file passed.** The parity test subtracted key sets one way only, so a
  101-entry subset of a 152-entry snapshot cleared both the length check and every
  comparison. Up to 51 cases could vanish in silence. Now exact equality, both directions.
- **The isolation test asserted nothing.** It read
  `assert forbidden not in sys.modules or True` — unconditionally true. It also could not
  have worked in-process, since another test importing tkinter first would poison
  `sys.modules`, and on Windows CI every forbidden module is installed. It now runs in a
  fresh interpreter and inspects the real module graph.
- **The lazy-import exemption was keyed on (file, module).** An exemption justified for one
  function licensed that import anywhere in the file — the laundering path the surrounding
  comment claimed to close. Now keyed on the enclosing function too.

The pattern worth keeping: every one of these was a check that passed while measuring less
than it appeared to. That is the failure mode this repository keeps finding, and the reason
each guard ships with a test that breaks the thing it defends.

### What phase 2 taught the guards

`scripts/check_names_resolve.py` is new, and it exists because of a gap phase 1's review
found rather than a rule anyone wrote up front. The staged extraction leaves re-export
shims behind, and **no test constructs `RSVPApp`** — doing so needs a display and a
signed-in Outlook. So a name dropped from a shim passes the entire suite and surfaces as a
`NameError` in front of a user, mid-send. Demonstrated: with `build_subject` removed from
the shim, all 30 tests passed and the guard failed immediately.

It then caught a real defect within the hour — `rsvp/export/legacy_excel.py` was moved out
of `db.py` without carrying `from datetime import datetime` with it. That is a crash in the
migration path, introduced and caught inside the same change.

`check_no_pii.py` also changed twice here. It excluded `example.com` case-sensitively, so
`Alice@Example.com` in a new test read as a real person; and it scanned only *tracked*
files, meaning a new file full of addresses passed right up until the commit that
introduced the leak. It now scans what `git add -A` would stage.

### What phase 1 taught the guards

Two guards were wrong in ways only a real move could expose, and both were the same class
of fault — a check that kept passing while its coverage quietly fell away:

- `check_i18n_matrix.py` pointed at `rsvp_app.py`. When 20 of the 24 language tables moved
  into `rsvp/i18n/`, it reported OK while checking only the 4 left behind. It now scans the
  whole application tree, and prints per-file counts so a drop is visible.
- `tests/test_guards.py` mutated `GREETING` in `rsvp_app.py` to prove the matrix guard
  fires. After the move that mutation applied to nothing — caught only because the test
  asserts its own mutation took effect. A mutation test that silently mutates nothing is
  the purest form of false confidence.

Neither was found by reading the diff. Both were found by running the checks against a tree
that had actually changed shape, which is the argument for doing the extraction in small
phases rather than one large one.

## Personal data (D4)

Two leaks, same commit, same severity:

- `rsvp_data.db` — 3 events, 5 recipients, 24 responses, 23 gift rows: real names, real
  `@{vn,jp,bcn}.bosch.com` addresses, real contribution amounts.
- `RSVP_tool.pdf` — 28 pages, 125 raster screenshots of the live app, marked
  `Internal C-SC1 | © Robert Bosch GmbH`. Page 20 shows ten employees with full names, org
  codes and addresses. **No text scan can see this**: the file has no `@` in its text layer
  at all. The first version of `check_no_pii.py` read it as clean, which is exactly the
  false confidence a guard is supposed to remove. The guard now refuses to pass while any
  tracked binary document is unreviewed.

Both are untracked and ignored. Both remain on disk.

**Still open, and it is the user's call: git history.** Untracking stops the next commit;
it does not remove either file from the commits already published. The repository has
exactly one commit, so the purge is mechanically cheap — but it rewrites published history
and cannot be undone, so it is escalated rather than executed. For a repo that is or was
public, treat the data as already disclosed and rotate what can be rotated.

## Known defects, characterized not fixed

Recorded here so a refactor cannot quietly absorb them. Each has a test that locks in
current behavior; fixing them is a behavior change belonging to Phase 2.

- **U3** — `parse_amount_from_text` returns the first numeric run in the string:
  `"2026 year-end party, 3000 JPY"` yields `2026.0`. Feeds gift and attendance totals.
  Also `float`, for money.
- **U4** — `_on_amount_paid_changed` wraps its database write in `except Exception: pass`.
  A failed save is invisible to the user.
- **U5** — `README.md` documents `send_scheduled_reminders.py` and
  `run_scheduled_reminders.bat`. Neither exists in the repository. An agent will believe
  the README.
- **U6** — `history.py` is half-dead: its Excel history is superseded by `db.py`, but it
  still owns the live JSON override files. The module name no longer describes it.

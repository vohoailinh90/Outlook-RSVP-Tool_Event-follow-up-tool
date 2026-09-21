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

**Phase 2 — extract `domain/`.** Money, roster building, vote resolution. This is where
U3 and U4 below get fixed, because fixing them changes behavior and needs the tests Phase 1
establishes. Move the Excel migration out of `db.py` and delete its `LAZY_ALLOWED` entry.

**Phase 3 — `ports/` + the adapter seam.** Define `OutlookPort`, make `outlook_com.py`
implement it, and give the tests a fake. This is what finally makes code above the seam
testable without Outlook. Judge it by whether a test can exercise a send path against the
fake — a directory rename that leaves everything still calling `outlook_com` directly has
bought nothing.

**Phase 4 — split `ui/`.** Seven tabs, seven modules, out of the 4,673-line class. Largest
and last, because it is worth least until the layers beneath it are real.

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

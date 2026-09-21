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

**Phase 1 — extract `i18n/`.** The 766 lines. Zero Tkinter coupling, and every
`outlook_com.*` call site in `rsvp_app.py` is already confined to `RSVPApp` methods, so the
only thing blocking extraction is the top-level import — not functional coupling. Highest
leverage per unit of risk in the repository, and the characterization tests in
`tests/test_message_builders.py` already lock in current behavior so the move can be proven
to change nothing. After this phase the message builders are testable on any machine, which
shortens the local edit-test loop; CI itself stays on Windows either way.

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

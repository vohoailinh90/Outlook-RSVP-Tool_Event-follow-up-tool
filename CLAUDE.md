# Outlook RSVP / Event follow-up tool

Instructions for Claude Code working in this repository.

## Target platform — Windows only

This tool runs on Windows and only on Windows: `outlook_com.py` drives Outlook
through COM, which does not exist off Windows. Linux and macOS are not
supported, not tested and not targeted.

- **Do not spend analysis, implementation, review or verification effort on
  Linux/macOS-only behavior** — POSIX paths, case-sensitive filesystems,
  permission bits and `chmod`, symlinks, signals, `sh`/`bash` quoting, POSIX
  locale and encoding defaults, LF-vs-CRLF differences that only appear off
  Windows.
- **A defect that can only occur off Windows is out of scope.** If you notice
  one, record it in a single line and move on: never widen scope to fix it,
  never let it block a review or a merge. Portability is not a requirement, so
  it never produces a review finding on its own.
- **Windows behavior itself is fully in scope.** COM object lifetime and
  early/late binding, console encoding (`cp1252` vs UTF-8), CRLF line endings,
  file locking while Outlook holds an item, drive-letter paths and path-length
  limits are Windows concerns, not portability concerns — they are where this
  tool actually breaks.
- **Existing cross-platform code stays as it is.** This rule stops new effort;
  it is not a licence to strip working non-Windows branches out of code that
  already has them. Removing them is a behavior change that buys nothing here.

A requirement that explicitly asks for non-Windows support is a change to this
constraint: treat it as a requirement change and confirm with the user before
doing the work.

## Agent roster

Three agents, and the reason there are only three is the point: an agent earns its place
only by supplying judgement the main session cannot supply for itself. A second context
that re-derives what the main session already knows costs more than it returns.

| Agent | Use it for |
|---|---|
| `rsvp-reviewer` | Default review + verification for routine changes |
| `rsvp-reviewer-critical` | Outlook COM, sending/scanning mail, personal data, money, schema changes |
| `architecture-critic` | Challenges a design **before** it is implemented |

The main session does the analysis, the design and the implementation. It does not review
its own diff: the session that wrote the code cannot be its own only verification.

**Deliberately not in the roster yet.** Each has a named trigger, so the team grows on
evidence rather than on anticipation:

- `outlook-com-specialist` — add when Phase 3 needs COM *design authorship*, not review.
  Until then its checklist lives inside `rsvp-reviewer-critical`, where it is already used.
- `test-engineer` — escalation only, when `rsvp-reviewer-critical` formally reports that
  verification needs designing (failure injection, migration rehearsal).
- A refactor implementer — not planned. For long mechanical work the answer is a written
  handoff that survives a restart, not a subagent whose context does not.

## Deterministic work is not agent work

If the answer to a check is computable, compute it. A model comparing two values can be
wrong; a script cannot. Never ask an agent to decide something in `scripts/`:

```bash
python scripts/check_no_pii.py        # no personal data is tracked in git
python scripts/check_i18n_matrix.py   # every message table has en/ja/vi
python scripts/check_layering.py      # layers stay importable without heavy deps
python -m pytest tests/ -q            # characterization + guard self-tests
```

Every guard has a test in `tests/test_guards.py` that **breaks the thing it defends and
requires the guard to notice**. A guard that passes on a clean tree proves nothing on its
own — it would also pass if it were blind. When you add a guard, add its negative test in
the same change. When you fix a bug, add the mutation that would have caught it.

## Personal data

Recipients are real colleagues; the repository is public. Real names, email addresses and
contribution amounts have been committed here before — once as `rsvp_data.db`, once as
screenshots inside `RSVP_tool.pdf` that no text scan could see. Both are untracked now and
`check_no_pii.py` fails the build if either class returns.

- Never put real recipient data in a test fixture, a commit, an artifact or a finding.
  Build example data instead.
- A tracked binary (PDF, image, spreadsheet) must be listed in `.pii-allowlist` by a human
  who opened it. Do not add a line there to make the build green.
- `rsvp_data.db` stays on disk and out of git. It is the user's live data: never delete it.

## Structure

`docs/agentic/ARCHITECTURE.md` holds the target layering and the staged plan to get there.
The short version: `rsvp_app.py` is 5,779 lines, of which 766 are already free of Tkinter,
and those come out first. Read that document before proposing a structural change.

## Verification

Implementation is not complete because code was written.

- Routine change: guards + tests + an independent diff review by `rsvp-reviewer`.
- High-risk change: the above via `rsvp-reviewer-critical`, plus an explicit statement of
  what happens to an existing `rsvp_data.db` and how to roll back.
- COM code cannot be executed anywhere. Say "verified by reading" and mean it, rather than
  implying you ran something.

## Sending email is irreversible

This tool sends real mail to real colleagues and creates real calendar invites. Before
changing any send path, be specific about whether the change could send to the wrong
audience, send twice, or send on a path that previously only drafted a message. `Display()`
opens a draft; `Send()` is gone the moment it runs. Prefer drafting.

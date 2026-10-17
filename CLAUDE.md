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

### What these guards cannot see

Stated plainly, because a guard's blind spots are where your judgement is still required.
All five of these were found by attacking the guards and are now caught; the limits below
are what remains after that.

- **Names without an address.** A name is personal data and no pattern recognises one. The
  guard does not pretend to: it refuses to pass while a tracked `.csv`/`.tsv` or binary
  document is unaccounted for, and makes a person say what is in it. A name sitting in a
  `.md` or `.txt` file still gets through.
- **An address at a non-ASCII domain.** The email pattern is ASCII, so `user@` followed by
  a Unicode (IDN) domain is not seen. The Punycode form (`xn--...`) is.
- **A computed module name.** `importlib.import_module(name)` with a variable is beyond
  static analysis. Constant arguments are checked.
- **A message table built at runtime** rather than written as a dict literal is invisible
  to the AST check.
- **A wrong translation.** The matrix check proves a language is *present*, never that it
  is correct.
- **What the Outlook wiring does at run time.** `check_layering.py` fails every ordinary way
  for `rsvp_app.py` to reach `outlook_com` except through `self.outlook`: an attribute call,
  an alias, a from-import, a second binding, the module name as a string. It also pins the
  one allowed use to `self.outlook = outlook if outlook is not None else outlook_com` in
  `RSVPApp.__init__`. It checks the shape of that line, not the data flow into it. Code that
  deliberately defeats it, for example `outlook = None` just above, or a module name built
  from pieces, is beyond static analysis and belongs to human review.
- **When a closure is called.** `check_names_resolve.py` lets a nested function read any
  name its enclosing function binds, anywhere, because a closure resolves it at call time.
  If the closure is called *before* that binding runs (`if flag: return inner()` above
  `import os`), Python raises `NameError` and the guard does not see it. Telling the two
  apart needs control-flow analysis, which is beyond this check.

## Personal data

Recipients are real colleagues; the repository is public. Real names, email addresses and
contribution amounts have been committed here before — once as `rsvp_data.db`, once as
screenshots inside `RSVP_tool.pdf` that no text scan could see. Both are untracked now and
`check_no_pii.py` fails the build if either class returns.

- Never put real recipient data in a test fixture, a commit, an artifact or a finding.
  Build example data instead.
- A tracked binary (PDF, image, spreadsheet) must be listed in `.pii-allowlist` by a human
  who opened it, with the `blob=` digest of the exact file they saw. Do not add a line
  there to make the build green.
- `rsvp_data.db` stays on disk and out of git. It is the user's live data: never delete it.

## Structure

The root holds only what a user runs: `rsvp_app.py`, README, the requirements files and
`how_to_vote.png`. Every module lives in `rsvp/` (storage in `rsvp/storage/`, the Outlook
COM adapter in `rsvp/adapters/`, Excel in `rsvp/export/`), tests in `tests/`, guards in
`scripts/`. `python scripts/layout_check.py` is the verdict and CI runs it; when a change
adds a Python file, test or module, use `.claude/skills/repo-layout/SKILL.md`.

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

## Pull requests: request the Codex review yourself

Automatic Codex review is turned off on this repository. Codex reviews a pull request only
when someone asks, so asking is part of opening one:

- Right after opening a PR, comment `@codex review` on it.
- After every push that answers Codex's findings, comment `@codex review` again: it does not
  look at a new head on its own.
- Codex answers with review comments, one finding each, or with "Didn't find any major
  issues" (or a 👍 reaction). Keep going until the review is clean as *Pull request
  lifecycle* below defines it (a P2/P3 finding whose rejection with evidence stands does not
  keep the loop going).
- A finding is a claim to verify, not an order. If it is real, fix it with a test that fails
  without the fix. If not, reply on its thread with the reason.

## Sending email is irreversible

This tool sends real mail to real colleagues and creates real calendar invites. Before
changing any send path, be specific about whether the change could send to the wrong
audience, send twice, or send on a path that previously only drafted a message. `Display()`
opens a draft; `Send()` is gone the moment it runs. Prefer drafting.

## Pull request lifecycle

Every pull request you create or drive for a requirement goes through the loop below, whatever the size of the change. It is a standing rule, not a judgement call. It starts when a PR exists; it does not authorize opening one, so a PR is still opened only when the user asks for it or the task's own flow calls for it.

1. **Open, then Codex review — immediately.** The first action after the PR exists is the Codex review: a PR comment whose whole body is `@codex review`, and `@codex security review` as well when the change touches authentication, authorization, secrets, payments or personal data. Codex does not review a PR on its own unless the repository turns that on, so ask for it. A PR with no Codex review is not done and is never merged.
2. **Validate every finding before acting.** Check it against the code. A confirmed finding is fixed; a wrong one is answered on its thread with the evidence — never silently skipped, never "fixed" speculatively.
3. **Fix, then have the new head reviewed.** Fix each confirmed finding — with a regression test where the defect has testable behavior; a prose-only fix keeps the repository's normal checks instead — run the repository's own checks, push, and request the next review on the new head: each review command once per head commit, one request at a time, never before the previous response was validated.
4. **Stop when the review is clean.** Clean means every review command the change requires (the standard review, and the security review where it applies) has a review of the *current* head with no finding (other than a P2/P3 whose recorded rejection stands, see the first bullet below), every earlier finding was fixed or rejected with evidence, CI and the required checks are green, and there is no merge conflict. A fix pushed after the last review is unreviewed, so it is not clean — review it with every command the change requires.
5. **Merge it yourself — where the repository allows it.** A clean PR is merged without asking for approval, scoped to the validated head commit, provided the base branch requires branches to be up to date and the merging identity cannot bypass that rule (read the branch protection to confirm both). Where either is missing, merging is a genuine blocker: leave the PR open, say it is clean and why it was not merged, and do not merge. Record the merge commit and a one-line summary on the PR.
6. **Then continue.** Do not stop after the merge to ask what is next: start the next outstanding task from the latest default branch (a merged branch is finished; never stack new commits on it). An outstanding task is one the user explicitly asked for that is not done yet; never invent work or pick up an item the user did not ask for. A task that needs the code of a PR that is not merged waits: skip it rather than stack on that branch. When none remains, report what merged and stop.

What bounds the loop:

- **No review cap.** The loop runs until every review command the change requires is clean on the current head. A review is still bounded by step 3's rules. Two exits remain: a finding you reject with evidence — a P2/P3 stands at once and blocks nothing, even if Codex raises it again; a P0/P1 is a disagreement unless a review drops it: when other fixes make a new head, push it and let the next review decide (raised again, unchanged or reworded or moved when it is substantially the same issue, means escalate; not raised again means closed), and when no code change is pending escalate at once — report it to the repository owner and do not merge until they decide — and a genuine blocker.
- **Codex unreachable.** If Codex gives no reaction or review after a reasonable wait, or answers with an error or a usage-limit notice, the PR is BLOCKED: report it once, leave the PR open and continue with the next task that does not depend on it. Retry at most once on an unchanged head, and never call Codex another way (API, Action, CLI, SDK, browser automation).
- **The trigger string.** No PR description, comment or commit message carries the literal trigger string except the intentional review requests of steps 1 and 3: the standard review, the security review and the single retry.
- **Who may stop the loop.** A human is needed only for a genuine blocker — the task cannot proceed without a decision only a human can make — or a change to the original requirement.

<!-- team-onboarding:start -->
## Team agent process

`vohoailinh90/Outlook-RSVP-Tool_Event-follow-up-tool` is registered with the free-model team: `team` (Claude leads) and `team-codex` (Codex leads), with free models writing the tests, the code and a pre-review. The details are in `docs/TEAM.md`; the settings that differ per repository live in `repos.conf` on the VPS, not here.

- Base branch `main`; sensitivity `public`; gates setup `python3 -m venv --system-site-packages .venv && (grep -viE '^[[:space:]]*(pywin32|pywinauto|comtypes|wmi|pythonnet|pyinstaller[-_.]+hooks[-_.]+contrib[-_.]+win|winshell)([^A-Za-z0-9._-]|$)' requirements.txt > .team-linux-requirements.txt || true) && (grep -viE '^[[:space:]]*(-r|--requirement)[[:space:]=]*["'"'"']?(\./)?(requirements\.txt)["'"'"']?[[:space:]]*(#.*)?$' requirements-dev.txt > .team-linux-requirements-dev.txt || true) && .venv/bin/pip install -q -r .team-linux-requirements.txt -r .team-linux-requirements-dev.txt`; suggested test commands `.venv/bin/python scripts/check_i18n_matrix.py`, `.venv/bin/python scripts/check_layering.py`, `.venv/bin/python scripts/layout_check.py`, `.venv/bin/python -m pytest -q`.
- If you are a team worker (tester, coder or reviewer), follow the handoff files you are given and touch only the files the task names. Never edit CI (`.github/`), agent instructions (`CLAUDE.md`, `AGENTS.md`, `.claude/`, `.codex/`), lockfiles or `repos.conf` unless the task says so.

If you are Claude Code or Codex working on this repository directly (a spec, an audit or an ordinary change), this is the pull-request process:

- When your work here is ready to ship, open the pull request yourself; do not wait to be asked, unless the user asked for a local change only or this repository's own rules say a pull request is opened only on request.
- As soon as it is open, ask Codex for a review with one comment on the PR: `@codex review` (add `@codex security review` where this repository's rules, or the change itself, call for one: authentication, authorization, secrets, payments, personal data). Codex never reviews on its own here. Ask once per head commit for each review command, never while a request is unanswered, and never mention Codex in the PR description or a commit message. Where the repository has `scripts/codex_round_count.py`, run it before each request.
- Check each finding against the code. Confirmed: fix it, with a regression test where the defect has testable behavior (a prose-only fix keeps the repository's normal checks instead), push, and ask again on the new head. Wrong: answer on its thread with the evidence. There is no cap on rounds; a rejected P2/P3 stands and blocks nothing; a P0/P1 you cannot settle goes to the repository owner and the PR is not merged.
- Merge it yourself, without waiting for approval, once all of these hold on the current head: every review this change requires (the standard review, plus the security review where it touches authentication, authorization, secrets, payments or personal data) has been requested of that exact commit and has no finding left (a bare thumbs-up is not a review, unless this repository's own rules accept one for that exact request), CI is green, there is no merge conflict, and the base branch's protection (or a merge queue) requires branches to be up to date and applies to admins, and the identity that merges has no bypass of it (ruleset bypass list, custom role or app privilege). Immediately before merging, re-read the PR head and the base head: if the PR head moved, review it again; if the base moved, merge it in and re-validate. Merge scoped to the validated head SHA. If that protection is missing, do not merge: leave the PR open, say it is clean and why it was not merged, and tell the owner.
- If Codex gives no sign of life within 10 minutes, retry once per head; an explicit error or usage-limit reply needs no wait before that one retry. If it acknowledges but never delivers a review, or its task fails, do not retry: the PR is BLOCKED. The PR is also BLOCKED when the retry gets no answer. A BLOCKED PR stays open and is reported once; never merge it unreviewed, and never use another route (API, CLI, Action) to get a review.

Exceptions in this repository, which replace the matching rule above:

- The repository's agent roster (CLAUDE.md "Agent roster") is rsvp-reviewer, rsvp-reviewer-critical and architecture-critic, and the main session implements. The coder and tester of the free-model team are not roster agents: their branch is a proposal that the main session takes over, and a roster reviewer still reviews and verifies the diff before it is merged.
- The worker image has no Windows, pywin32 or tkinter, so on the team gates the tests that use the monolith fixture are skipped (tests/conftest.py skips them unless RSVP_REQUIRE_APP_IMPORT=1). A green team gate does not cover them: the Windows CI job sets that variable, and its result on the pushed head is the evidence. scripts/check_no_pii.py needs a git checkout and is not part of the gate.

Where this repository documents its own rule for one of these steps elsewhere (who merges, what gates a merge, how the review is requested, a cap on rounds), that rule governs that step; this process covers whatever the repository leaves unsaid.
<!-- team-onboarding:end -->

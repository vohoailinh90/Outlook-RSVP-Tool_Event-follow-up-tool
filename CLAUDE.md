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

## Scope of this file

This file records the platform constraint and nothing else. It deliberately
does not define a routing policy, agent roster or review process — this
repository has not adopted one, and inventing rules here that nobody agreed to
would be worse than the silence it replaces. Follow the conventions already in
the code and in `README.md`.

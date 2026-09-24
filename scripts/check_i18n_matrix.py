#!/usr/bin/env python3
"""Fail if a translatable string table is missing a language.

Why this is a script and not a review question: every language table in
rsvp_app.py is read through `TABLE.get(lang_code, TABLE["en"])`. A table that
is missing "vi" does not raise — it silently serves English to Vietnamese
recipients. Nothing in the running app ever reports it, and a reviewer reading
a diff cannot see the absence of a key that was never typed. A script can.

Parses the AST rather than importing: rsvp_app.py imports tkinter and
outlook_com, so it cannot be imported in an environment without a display or
without Outlook. The tables are module-level dict literals, which the AST gives
us exactly.

Exit 0 = every table complete. Exit 1 = a language is missing. Exit 2 = the
file could not be parsed.
"""
from __future__ import annotations

import ast
import sys
from pathlib import Path

REQUIRED = {"en", "ja", "vi"}
# "bilingual" is composed from en+ja at build time rather than stored, so it is
# a legal key but never a required one.
KNOWN_LANGS = REQUIRED | {"bilingual"}
ROOT = Path(__file__).resolve().parents[1]

# Every application source file, not one hardcoded path. The phase 1 extraction
# moved 20 of the 24 language tables from rsvp_app.py into rsvp/i18n/messages.py
# and this check - then pointed at rsvp_app.py alone - reported OK while
# silently covering only the 4 tables left behind. Coverage that follows the
# code cannot drift that way. scripts/ and tests/ are excluded: they contain
# deliberately broken tables as fixtures.
TARGET_GLOBS = ("rsvp_app.py", "rsvp/**/*.py")


def targets() -> list[Path]:
    seen: list[Path] = []
    for pattern in TARGET_GLOBS:
        for path in sorted(ROOT.glob(pattern)):
            if path.is_file() and path not in seen:
                seen.append(path)
    return seen

# A dict literal is a language table when every one of its keys is a language
# code we recognise and at least one is a required language.
#
# An earlier version demanded TWO required languages before it would look at a
# table. That made the guard blind to the most likely real mistake: adding a
# new message and writing only the English line. A table of {"en": ...} alone
# was not "incomplete", it was invisible. Keys being a SUBSET of the known
# codes is what makes this both narrow (LANG_LABELS-style maps have other keys,
# so they are skipped) and complete (an en-only table is caught).
def _is_lang_table(node: ast.Dict) -> bool:
    if not node.keys:
        return False
    keys = [k.value for k in node.keys
            if isinstance(k, ast.Constant) and isinstance(k.value, str)]
    if len(keys) != len(node.keys):
        return False
    return set(keys) <= KNOWN_LANGS and bool(REQUIRED.intersection(keys))


def main() -> int:
    paths = targets()
    if not paths:
        print("i18n-matrix: FAIL - no application source found", file=sys.stderr)
        return 2

    failures = []
    checked = 0
    per_file = []
    for target in paths:
        try:
            tree = ast.parse(target.read_text(encoding="utf-8"))
        except SyntaxError as exc:
            print(f"i18n-matrix: cannot parse {target.name}: {exc}", file=sys.stderr)
            return 2

        found = 0
        for node in ast.walk(tree):
            if not isinstance(node, ast.Dict) or not _is_lang_table(node):
                continue
            checked += 1
            found += 1
            keys = {k.value for k in node.keys}
            missing = REQUIRED - keys
            if missing:
                failures.append((target, node.lineno, sorted(missing), sorted(keys)))
        if found:
            per_file.append((target.relative_to(ROOT), found))

    if not checked:
        # A guard that checks nothing must fail loudly, not pass quietly.
        print("i18n-matrix: FAIL - no language tables found anywhere; the "
              "detector has drifted from the code it guards", file=sys.stderr)
        return 1

    for target, lineno, missing, keys in failures:
        print(f"{target.relative_to(ROOT)}:{lineno}: language table missing {missing} "
              f"(has {keys}) - .get(lang, T['en']) will silently serve English",
              file=sys.stderr)

    if failures:
        print(f"\ni18n-matrix: FAIL - {len(failures)} of {checked} tables incomplete",
              file=sys.stderr)
        return 1
    summary = ", ".join(f"{p} ({n})" for p, n in per_file)
    print(f"i18n-matrix: OK - {checked} language tables across {len(per_file)} "
          f"file(s), all have {sorted(REQUIRED)}")
    print(f"             {summary}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

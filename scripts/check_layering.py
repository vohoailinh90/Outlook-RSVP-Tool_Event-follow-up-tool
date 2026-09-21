#!/usr/bin/env python3
"""Fail if a module imports across a layer boundary it is not allowed to cross.

The decomposition in docs/agentic/ARCHITECTURE.md is only real if something
enforces it. This is that something.

The rule that matters: the pure layers (domain, i18n) and the storage layer
must not import tkinter, pywin32 or openpyxl. The moment they do, they stop
being testable without a display, without Outlook, and the seam is gone -
which is precisely how rsvp_app.py's 766 lines of platform-neutral template
code ended up untestable: they sit in a module that imports outlook_com at
the top.

This guard is written to be meaningful TODAY (db.py is already stdlib-only and
must stay that way) and to keep being meaningful as each extraction phase
lands, by adding the new package to LAYERS. A guard that only describes a
future state enforces nothing.

Exit 0 clean, 1 violation, 2 could not run.
"""
from __future__ import annotations

import ast
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

HEAVY = {
    "tkinter": "a display",
    "tkcalendar": "a display",
    "win32com": "Windows + Outlook",
    "pythoncom": "Windows + Outlook",
    "pywintypes": "Windows + Outlook",
    "openpyxl": "the openpyxl dependency",
}

# Function-local imports of a heavy dependency are NOT automatically fine.
# Deferring an import keeps the module importable, but the layer still depends
# on the package at runtime, and "just move it inside a function" would
# otherwise be a standing way to launder any future violation past this guard.
# So each one must be declared here, with a reason, and anything undeclared
# fails exactly like a module-level import. This is a ratchet: the existing
# case is grandfathered visibly, new ones are not.
LAZY_ALLOWED: dict[tuple[str, str], str] = {
    ("db.py", "openpyxl"): (
        "one-time Excel->SQLite migration (migrate_from_excel_if_needed). "
        "db.py mixes storage and export concerns; phase 2 moves this out, "
        "after which this entry must be deleted."
    ),
}

# path glob -> set of top-level module names it may NOT import.
# Phase 1 adds "rsvp/i18n/**", phase 2 "rsvp/domain/**", and so on.
LAYERS: list[tuple[str, set[str]]] = [
    ("db.py", set(HEAVY)),
    ("scripts/*.py", set(HEAVY)),
    ("rsvp/domain/**/*.py", set(HEAVY)),
    ("rsvp/i18n/**/*.py", set(HEAVY)),
    ("rsvp/storage/**/*.py", set(HEAVY)),
    ("rsvp/ports/**/*.py", set(HEAVY)),
]


def imported_roots(path: Path) -> tuple[set[str], set[str]]:
    """Return (module_level, function_local) top-level import names.

    The distinction is the whole point. A module-level `import openpyxl`
    makes the layer unimportable wherever openpyxl is absent, which is the
    boundary this guard defends. An import inside a function body - as
    db.py does for the one-time Excel migration - keeps the module
    importable and defers the cost to the caller that actually needs it.
    That is the correct pattern for an optional dependency, so flagging it
    would train people to switch the guard off.
    """
    try:
        tree = ast.parse(path.read_text(encoding="utf-8"))
    except (SyntaxError, OSError):
        return set(), set()

    nested: set[int] = set()
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            for child in ast.walk(node):
                if isinstance(child, (ast.Import, ast.ImportFrom)):
                    nested.add(id(child))

    top: set[str] = set()
    local: set[str] = set()

    # importlib.import_module("tkinter") and __import__("tkinter") are imports
    # that ast.Import never sees. Only a constant argument can be resolved
    # statically - a computed module name is beyond any static check, and this
    # guard does not pretend otherwise.
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call) or not node.args:
            continue
        fn = node.func
        name = (fn.attr if isinstance(fn, ast.Attribute)
                else fn.id if isinstance(fn, ast.Name) else None)
        if name not in {"import_module", "__import__"}:
            continue
        arg = node.args[0]
        if isinstance(arg, ast.Constant) and isinstance(arg.value, str):
            local.add(arg.value.split(".")[0])

    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            names = {a.name.split(".")[0] for a in node.names}
        elif isinstance(node, ast.ImportFrom) and node.level == 0 and node.module:
            names = {node.module.split(".")[0]}
        else:
            continue
        (local if id(node) in nested else top).update(names)
    return top, local


def main() -> int:
    violations: list[str] = []
    notes: list[str] = []
    checked = 0
    for pattern, forbidden in LAYERS:
        for path in sorted(ROOT.glob(pattern)):
            if not path.is_file():
                continue
            checked += 1
            top, local = imported_roots(path)
            for bad in sorted(top & forbidden):
                violations.append(
                    f"{path.relative_to(ROOT)}: imports {bad!r} at module level, "
                    f"which needs {HEAVY[bad]}. This layer must stay importable "
                    f"without it - move it inside the function that needs it."
                )
            rel = path.relative_to(ROOT).as_posix()
            for lazy in sorted(local & forbidden):
                reason = LAZY_ALLOWED.get((rel, lazy))
                if reason is None:
                    violations.append(
                        f"{rel}: lazily imports {lazy!r} inside a function. "
                        f"Deferring the import keeps the module importable but the "
                        f"layer still needs {HEAVY[lazy]} at runtime. Declare it in "
                        f"LAZY_ALLOWED with a reason, or move it out of this layer."
                    )
                else:
                    notes.append(f"{rel}: lazy {lazy!r} allowed - {reason}")

    if not checked:
        print("layering: FAIL - matched no files; LAYERS has drifted from the tree",
              file=sys.stderr)
        return 1
    for n in notes:
        print(f"layering: note: {n}")
    for v in violations:
        print(f"layering: {v}", file=sys.stderr)
    if violations:
        print(f"\nlayering: FAIL - {len(violations)} violation(s)", file=sys.stderr)
        return 1
    print(f"layering: OK - {checked} file(s) respect their layer boundary")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

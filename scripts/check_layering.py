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
    # Not a third-party package, but the same kind of dependency: the COM
    # adapter. Code above the seam takes an OutlookPort (rsvp/ports/) instead.
    "outlook_com": "Windows + Outlook (take an OutlookPort instead)",
}

# Function-local imports of a heavy dependency are NOT automatically fine.
# Deferring an import keeps the module importable, but the layer still depends
# on the package at runtime, and "just move it inside a function" would
# otherwise be a standing way to launder any future violation past this guard.
# So each one must be declared here, with a reason, and anything undeclared
# fails exactly like a module-level import. This is a ratchet: the existing
# case is grandfathered visibly, new ones are not.
# Keyed on (path, ENCLOSING FUNCTION, module). An earlier version keyed only on
# (path, module), so an exemption written for one function silently licensed
# that import anywhere in the file - adding `import openpyxl` to load_history
# would have been accepted by an exemption justified for the migration routine.
# That is the laundering path the comment above claims to close. Reported by
# Codex review of 194af3c.
#
# Empty by design right now: db.py's lazy openpyxl import was the only entry,
# and phase 2 moved that code to rsvp/export/legacy_excel.py, so the storage
# layer is stdlib-only and the exemption is gone rather than grandfathered.
LAZY_ALLOWED: dict[tuple[str, str, str], str] = {}

# path glob -> set of top-level module names it may NOT import.
# Phase 1 adds "rsvp/i18n/**", phase 2 "rsvp/domain/**", and so on.
LAYERS: list[tuple[str, set[str]]] = [
    ("db.py", set(HEAVY)),
    ("scripts/*.py", set(HEAVY)),
    ("rsvp/domain/**/*.py", set(HEAVY)),
    ("rsvp/i18n/**/*.py", set(HEAVY)),
    ("rsvp/storage/**/*.py", set(HEAVY)),
    ("rsvp/ports/**/*.py", set(HEAVY)),
    ("rsvp/services/**/*.py", set(HEAVY)),
]

# The application above the Outlook seam may name outlook_com only to wire it
# in as the default OutlookPort. Any `outlook_com.<attr>` access is a call
# that bypasses the port, so a test's fake would never see it (phase 3).
SEAM_CLIENTS = ["rsvp_app.py"]
SEAM_MODULE = "outlook_com"
WIRING_CLASS = "RSVPApp"


def seam_bypasses(path: Path) -> list[tuple[int, str]]:
    """(line, what) for every way `path` could reach outlook_com without
    going through the injected OutlookPort.

    Allowed: a plain `import outlook_com`, and ONE use of the name - the
    default wiring. Anything else is a bypass: an attribute call, a second
    use of the name (`oc = outlook_com` then `oc.send_...`), an aliased
    import, or `from outlook_com import ...` (Codex review of PR #1: the
    first version looked only for `outlook_com.<attr>`)."""
    try:
        tree = ast.parse(path.read_text(encoding="utf-8"))
    except (SyntaxError, OSError):
        return []
    found: list[tuple[int, str]] = []
    uses: list[ast.Name] = []
    parent: dict[int, ast.AST] = {}
    for node in ast.walk(tree):
        for child in ast.iter_child_nodes(node):
            parent[id(child)] = node
    for node in ast.walk(tree):
        # The module named as a string - import_module("outlook_com"),
        # import_module(name="outlook_com"), __import__, sys.modules[...] -
        # is a way to reach it without an import statement. Any string
        # constant that IS the module name counts, whatever consumes it, so
        # a new call shape cannot slip past (Codex review of PR #8). Prose
        # that merely mentions the module in a longer string is not matched.
        if (isinstance(node, ast.Constant) and isinstance(node.value, str)
                and (node.value == SEAM_MODULE
                     or node.value.startswith(SEAM_MODULE + "."))):
            found.append((node.lineno,
                          f"names {SEAM_MODULE} as a string (a dynamic "
                          f"import)"))
        if isinstance(node, ast.Import):
            for alias in node.names:
                if (alias.name.split(".")[0] == SEAM_MODULE
                        and alias.asname is not None):
                    found.append((node.lineno, f"imports {SEAM_MODULE} as "
                                  f"{alias.asname!r}"))
        elif (isinstance(node, ast.ImportFrom) and node.module
              and node.module.split(".")[0] == SEAM_MODULE):
            found.append((node.lineno, f"imports names from {SEAM_MODULE}"))
        elif isinstance(node, ast.Attribute) and isinstance(
                node.value, ast.Name) and node.value.id == SEAM_MODULE:
            found.append((node.lineno, f"calls {SEAM_MODULE}.{node.attr}"))
        elif (isinstance(node, ast.Name) and node.id == SEAM_MODULE
              and not isinstance(parent.get(id(node)), ast.Attribute)):
            uses.append(node)
    # Exactly one bare use is allowed, and only in the exact shape of the
    # default wiring: the `else` branch of the conditional assigned to
    # self.outlook. Allowing whichever use came first let
    # `(oc := outlook_com)` inside the wiring keep a hidden alias (Codex
    # review of PR #8).
    wiring = [n for n in uses if _is_default_wiring(n, parent)]
    for node in uses:
        if node is not (wiring[0] if wiring else None):
            found.append((node.lineno, f"uses {SEAM_MODULE} outside the "
                                       f"default wiring of self.outlook"))
    return sorted(found)


def _is_default_wiring(node: ast.Name, parent: dict[int, ast.AST]) -> bool:
    """True only for the exact injection expression

        self.outlook = outlook if outlook is not None else outlook_com

    (annotation allowed). Checking only that outlook_com was the else
    branch passed `outlook if False else outlook_com`, which throws the
    injected fake away (Codex review of PR #8)."""
    ifexp = parent.get(id(node))
    if not (isinstance(ifexp, ast.IfExp) and ifexp.orelse is node):
        return False
    test = ifexp.test
    if not (isinstance(ifexp.body, ast.Name) and ifexp.body.id == "outlook"
            and isinstance(test, ast.Compare)
            and isinstance(test.left, ast.Name) and test.left.id == "outlook"
            and len(test.ops) == 1 and isinstance(test.ops[0], ast.IsNot)
            and len(test.comparators) == 1
            and isinstance(test.comparators[0], ast.Constant)
            and test.comparators[0].value is None):
        return False
    assign = parent.get(id(ifexp))
    if isinstance(assign, ast.AnnAssign):
        targets = [assign.target]
    elif isinstance(assign, ast.Assign):
        targets = assign.targets
    else:
        return False
    if assign.value is not ifexp:
        return False
    # self.outlook must be the ONLY target: `self.outlook = oc = ...` kept an
    # alias through the second one (Codex review of PR #8).
    if not (len(targets) == 1 and isinstance(targets[0], ast.Attribute)
            and targets[0].attr == "outlook"
            and isinstance(targets[0].value, ast.Name)
            and targets[0].value.id == "self"):
        return False
    # And it must be the constructor's wiring, taking `outlook` as its
    # parameter: the same line in another method would swap an injected fake
    # for the real adapter whenever that method ran (Codex review of PR #8).
    func = parent.get(id(assign))
    while func is not None and not isinstance(
            func, (ast.FunctionDef, ast.AsyncFunctionDef)):
        func = parent.get(id(func))
    cls = parent.get(id(func)) if func is not None else None
    return (func is not None and func.name == "__init__"
            and isinstance(cls, ast.ClassDef) and cls.name == WIRING_CLASS
            and _outlook_defaults_to_none(func.args))


def _outlook_defaults_to_none(args: ast.arguments) -> bool:
    """`outlook` is a parameter whose default is the literal None.

    The app's entry point calls RSVPApp() with no argument, so losing the
    default would stop the app starting, and any other default would wire a
    non-port in, while the guard stayed green (Codex review of PR #1)."""
    positional = args.posonlyargs + args.args
    defaults = [None] * (len(positional) - len(args.defaults)) + list(args.defaults)
    pairs = list(zip(positional, defaults)) + list(zip(args.kwonlyargs,
                                                       args.kw_defaults))
    return any(a.arg == "outlook" and isinstance(d, ast.Constant)
               and d.value is None for a, d in pairs)


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

    # id(import node) -> name of the function it sits in (innermost wins)
    nested: dict[int, str] = {}
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            for child in ast.walk(node):
                if isinstance(child, (ast.Import, ast.ImportFrom)):
                    nested[id(child)] = node.name
    for node in ast.walk(tree):
        if isinstance(node, ast.ClassDef):
            for child in ast.walk(node):
                if isinstance(child, (ast.Import, ast.ImportFrom)):
                    nested.setdefault(id(child), node.name)

    # Same attribution for importlib/__import__ calls.
    call_scope: dict[int, str] = {}
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            for child in ast.walk(node):
                if isinstance(child, ast.Call):
                    call_scope[id(child)] = node.name

    top: set[str] = set()
    local: set[tuple[str, str]] = set()   # (enclosing function, module)

    def enclosing_of(node) -> str | None:
        return nested.get(id(node))

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
            local.add((call_scope.get(id(node), "<module>"),
                       arg.value.split(".")[0]))

    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            names = {a.name.split(".")[0] for a in node.names}
        elif isinstance(node, ast.ImportFrom) and node.level == 0 and node.module:
            names = {node.module.split(".")[0]}
        else:
            continue
        enclosing = enclosing_of(node)
        if enclosing is None:
            top.update(names)
        else:
            local.update((enclosing, n) for n in names)
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
            for func, lazy in sorted(local):
                if lazy not in forbidden:
                    continue
                reason = LAZY_ALLOWED.get((rel, func, lazy))
                if reason is None:
                    violations.append(
                        f"{rel}: {func}() lazily imports {lazy!r}. Deferring the "
                        f"import keeps the module importable but the layer still "
                        f"needs {HEAVY[lazy]} at runtime. Declare it in "
                        f"LAZY_ALLOWED as ({rel!r}, {func!r}, {lazy!r}) with a "
                        f"reason, or move it out of this layer."
                    )
                else:
                    notes.append(f"{rel}: {func}() lazy {lazy!r} allowed - {reason}")

    for name in SEAM_CLIENTS:
        path = ROOT / name
        if not path.is_file():
            violations.append(f"{name}: listed in SEAM_CLIENTS but missing")
            continue
        checked += 1
        for line, what in seam_bypasses(path):
            violations.append(
                f"{name}:{line}: {what}, bypassing the OutlookPort. Use "
                f"self.outlook, so a test's fake sees the call.")

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

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

This guard is written to be meaningful TODAY (rsvp/storage/db.py is already
stdlib-only and must stay that way) and to keep being meaningful as each extraction phase
lands, by adding the new package to LAYERS. A guard that only describes a
future state enforces nothing.

Exit 0 clean, 1 violation, 2 could not run.
"""
from __future__ import annotations

import ast
import re
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

# The COM adapter lives at rsvp/adapters/outlook_com.py. Its top-level import
# name is `rsvp`, which says nothing, so every import of the adapters package
# (absolute or relative) is reported under the HEAVY key "outlook_com".
ADAPTERS_PACKAGE = "rsvp.adapters"
ADAPTER_MODULE = "rsvp.adapters.outlook_com"


def layer_root(dotted: str) -> str:
    """The HEAVY key an absolute import name falls under: the adapters
    package counts as outlook_com, anything else by its top-level name."""
    if dotted == ADAPTERS_PACKAGE or dotted.startswith(ADAPTERS_PACKAGE + "."):
        return "outlook_com"
    return dotted.split(".")[0]


def absolute_from(path: Path, node: ast.ImportFrom) -> str | None:
    """The absolute module a `from X import ...` reads, resolving a relative
    import (`from ..adapters import outlook_com`) against the file's package.
    Skipping relative imports would let any module inside rsvp/ reach the
    adapter without a name the guard recognises."""
    return resolve_relative(path, node.level, node.module)


def resolve_relative(path: Path, level: int, module: str | None) -> str | None:
    """`level` leading dots plus `module`, read from the package of `path`."""
    if level == 0:
        return module
    try:
        parts = list(path.resolve().relative_to(ROOT).parent.parts)
    except ValueError:
        return module
    if level - 1 > len(parts):
        return None
    base = parts[:len(parts) - (level - 1)]
    return ".".join(base + ([module] if module else [])) or None


def dynamic_root(path: Path, name: str, package: str | None = None) -> str:
    """layer_root of a constant import_module/__import__ name. A relative
    name (`import_module("..adapters.outlook_com", __package__)`) is resolved
    against the calling file's package, like `from .. import`; without that a
    leading dot gave an empty root and the adapter passed (Codex review of
    PR #10).

    Root cause of two rounds of findings: modelling every way Python resolves
    a dynamic name is a race the guard keeps losing. So a name or `package`
    with an `adapters` / `outlook_com` component is the adapter, whatever it
    resolves against; only then is the rest resolved, against a constant
    `package` when one is given (`import_module(".x", "rsvp")`), else the
    caller's own package."""
    components = {part for text in (name, package or "") for part in text.split(".")}
    if components & {"adapters", SEAM_MODULE}:
        return "outlook_com"
    stripped = name.lstrip(".")
    level = len(name) - len(stripped)
    if not level:
        return layer_root(name)
    if package:
        base = package.split(".")
        base = base[:len(base) - (level - 1)] if level - 1 <= len(base) else []
        return layer_root(".".join(base + ([stripped] if stripped else [])))
    return layer_root(resolve_relative(path, level, stripped or None) or "")

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
    ("scripts/*.py", set(HEAVY)),
    ("rsvp/domain/**/*.py", set(HEAVY)),
    ("rsvp/i18n/**/*.py", set(HEAVY)),
    ("rsvp/storage/**/*.py", set(HEAVY)),
    ("rsvp/ports/**/*.py", set(HEAVY)),
    ("rsvp/services/**/*.py", set(HEAVY)),
]

# The application above the Outlook seam may name outlook_com only to wire it
# in as the default OutlookPort, imported exactly as
# `from rsvp.adapters import outlook_com`. Any `outlook_com.<attr>` access is a call
# that bypasses the port, so a test's fake would never see it (phase 3).
SEAM_CLIENTS = ["rsvp_app.py"]
SEAM_MODULE = "outlook_com"
WIRING_CLASS = "RSVPApp"


def seam_bypasses(path: Path) -> list[tuple[int, str]]:
    """(line, what) for every way `path` could reach outlook_com without
    going through the injected OutlookPort.

    Allowed: `from rsvp.adapters import outlook_com`, and ONE use of the
    name - the default wiring. Anything else is a bypass: an attribute call,
    a second use of the name (`oc = outlook_com` then `oc.send_...`), an
    aliased import, `from rsvp.adapters.outlook_com import ...` (Codex
    review of PR #1: the first version looked only for `outlook_com.<attr>`),
    or any other route to the adapters package (`import rsvp.adapters...`,
    `from rsvp import adapters`, `<x>.adapters`, `<x>.outlook_com`)."""
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
                and any(node.value == name or node.value.startswith(name + ".")
                        for name in (SEAM_MODULE, ADAPTERS_PACKAGE))):
            found.append((node.lineno,
                          f"names {node.value!r} as a string (a dynamic "
                          f"import)"))
        if isinstance(node, ast.Import):
            for alias in node.names:
                if layer_root(alias.name) == SEAM_MODULE:
                    found.append((node.lineno, f"imports {alias.name}"
                                  " (only `from rsvp.adapters import "
                                  "outlook_com` is the wiring import)"))
        elif isinstance(node, ast.ImportFrom):
            module = absolute_from(path, node) or ""
            if module == ADAPTERS_PACKAGE:
                for alias in node.names:
                    if alias.name != SEAM_MODULE or alias.asname is not None:
                        found.append((node.lineno, f"imports {alias.name} as "
                                      f"{alias.asname or alias.name!r} from "
                                      f"{ADAPTERS_PACKAGE}"))
            elif (layer_root(module) == SEAM_MODULE
                  or module.split(".")[0] == SEAM_MODULE):
                found.append((node.lineno, f"imports names from {module}"))
            elif module == "rsvp" and any(
                    a.name == "adapters" for a in node.names):
                found.append((node.lineno, "imports the adapters package"))
        elif isinstance(node, ast.Attribute) and isinstance(
                node.value, ast.Name) and node.value.id == SEAM_MODULE:
            found.append((node.lineno, f"calls {SEAM_MODULE}.{node.attr}"))
        elif isinstance(node, ast.Attribute) and node.attr in (
                SEAM_MODULE, "adapters"):
            found.append((node.lineno, f"reaches .{node.attr} as an attribute"))
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


# A string that reads as a module path: dotted identifiers, optional leading
# dots for a relative name. Prose has spaces and is never matched.
MODULE_PATH = re.compile(r"\.*[A-Za-z_]\w*(\.[A-Za-z_]\w*)*")


def adapter_references(path: Path, strings: bool = True) -> list[int]:
    """Lines where a layer names the adapter without an import statement.

    - An attribute: `import rsvp` then `rsvp.adapters.outlook_com.send_...`;
      the import alone names only `rsvp`, which every layer may import.
    - A string constant that is a module path with an `adapters` /
      `outlook_com` component: `load("rsvp.adapters.outlook_com")` after
      `from importlib import import_module as load`, `getattr(rsvp,
      "adapters")`, a package argument. Matching the NAME rather than the
      function that consumes it ends a run of review rounds (Codex, PR #10)
      that each found one more way to spell the call. Docstrings are skipped,
      and so are scripts/ (`strings=False`): the guards there must name the
      adapter to guard it, and none of them runs inside the app.
    """
    try:
        tree = ast.parse(path.read_text(encoding="utf-8"))
    except (SyntaxError, OSError):
        return []
    docstrings = {id(node.value) for node in ast.walk(tree)
                  if isinstance(node, ast.Expr) and isinstance(node.value, ast.Constant)}
    lines = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Attribute) and node.attr in ("adapters", SEAM_MODULE):
            lines.add(node.lineno)
        elif (strings and isinstance(node, ast.Constant) and isinstance(node.value, str)
              and id(node) not in docstrings
              and MODULE_PATH.fullmatch(node.value)
              and {"adapters", SEAM_MODULE} & set(node.value.split("."))):
            lines.add(node.lineno)
    return sorted(lines)


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
        if not isinstance(node, ast.Call):
            continue
        fn = node.func
        name = (fn.attr if isinstance(fn, ast.Attribute)
                else fn.id if isinstance(fn, ast.Name) else None)
        if name not in {"import_module", "__import__"}:
            continue
        keywords = {k.arg: k.value for k in node.keywords if k.arg}
        # The module name may be passed by keyword: import_module(name=...).
        arg = node.args[0] if node.args else keywords.get("name")
        if not (isinstance(arg, ast.Constant) and isinstance(arg.value, str)):
            continue
        module = arg.value
        # __import__(name, globals, locals, fromlist, level): a positive
        # level makes a bare name relative, `__import__("adapters", ..., 1)`.
        level = node.args[4] if name == "__import__" and len(node.args) > 4 else keywords.get("level")
        if (isinstance(level, ast.Constant) and isinstance(level.value, int)
                and level.value > 0 and not module.startswith(".")):
            module = "." * level.value + module
        package = node.args[1] if name == "import_module" and len(node.args) > 1 else keywords.get("package")
        package = (package.value if isinstance(package, ast.Constant)
                   and isinstance(package.value, str) else None)
        local.add((call_scope.get(id(node), "<module>"),
                   dynamic_root(path, module, package)))

    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            names = {layer_root(a.name) for a in node.names}
        elif isinstance(node, ast.ImportFrom):
            module = absolute_from(path, node)
            if not module:
                continue
            # `from rsvp import adapters` names the package in the alias.
            names = {layer_root(f"{module}.{a.name}") if module == "rsvp"
                     else layer_root(module) for a in node.names}
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
            if "outlook_com" in forbidden:
                for line in adapter_references(path, strings=rel.startswith("rsvp/")):
                    violations.append(
                        f"{rel}:{line}: names rsvp.adapters (outlook_com) as an "
                        f"attribute or a module-path string, which needs "
                        f"{HEAVY['outlook_com']}.")
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

#!/usr/bin/env python3
"""Fail if a module references a global name that does not exist.

The staged extraction leaves compatibility shims behind: rsvp_app.py re-exports
names that now live in rsvp/i18n/ and rsvp/domain/. If a shim misses one, every
test still passes - no test constructs RSVPApp or calls its 108 methods, because
doing so needs a display and a signed-in Outlook - and the failure appears as a
NameError in front of a user, on Windows, mid-send.

That gap is exactly why this is a script. "Did the shim cover everything?" is
decided by comparing two sets of names, which a reviewer does by eye and a
computer does exactly.

Resolves each module's own globals: its imports, its definitions, and builtins.
Anything referenced but unresolved is reported. Locals, comprehension targets,
parameters and attribute access are all handled by walking scopes properly.

Exit 0 clean, 1 unresolved name found, 2 could not run.
"""
from __future__ import annotations

import ast
import builtins
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TARGETS = ("rsvp_app.py", "db.py", "history.py", "outlook_com.py", "rsvp/**/*.py")
BUILTINS = set(dir(builtins)) | {"__file__", "__name__", "__doc__", "__package__"}


class _Scope:
    """One enclosing scope during the walk.

    `bound` grows in source order, as the walker passes each binding. `every`
    is every name the scope binds anywhere (precollected, for functions).

    Code runs in two ways, and the lookup follows it:
    - The body being walked, a class body and a comprehension run straight
      away, so they can only see what is `bound` so far. `def f(): print(x);
      x = 1` is an UnboundLocalError, and the guard reports it.
    - A nested def or lambda runs LATER, when it is called, so it sees
      everything its enclosing function binds (`every`). A closure defined
      above a later `import os` is valid (Codex review of PR #3).

    Known limit: WHEN a closure is called is not tracked. One called before
    the enclosing binding runs (`if flag: return inner()` above `import os`)
    raises NameError, and this guard passes it. Deciding that needs
    control-flow analysis; recorded in CLAUDE.md under "What these guards
    cannot see".
    """

    def __init__(self, kind: str, every: set[str] | None = None) -> None:
        self.kind = kind            # "function", "class" or "comprehension"
        self.bound: set[str] = set()
        self.every = every if every is not None else self.bound


class ScopeWalker(ast.NodeVisitor):
    """Collect names loaded at module scope that the module never binds."""

    def __init__(self, module_globals: set[str]) -> None:
        self.globals = module_globals
        self.unresolved: list[tuple[str, int]] = []
        self.scopes: list[_Scope] = []

    # --- scope helpers -------------------------------------------------
    def _bound(self, name: str) -> bool:
        immediate = True    # still inside code that runs straight away
        for scope in reversed(self.scopes):
            # A method does not see its class body's names.
            if scope.kind == "class" and not immediate:
                continue
            if name in scope.bound:
                return True
            if scope.kind == "function":
                if name in scope.every:
                    # Local to this function. Read before its binding, that is
                    # an UnboundLocalError even when a global of the same name
                    # exists: `x = 0; def f(): print(x); x = 1` (Codex review
                    # of PR #3). From a nested function, it resolves at call
                    # time.
                    return not immediate
                immediate = False   # outer scopes are seen at call time
        return name in self.globals or name in BUILTINS

    def _bind_target(self, node, scope: set[str]) -> None:
        for n in ast.walk(node):
            if isinstance(n, ast.Name):
                scope.add(n.id)

    def _enter_function(self, node) -> None:
        params: set[str] = set()
        args = node.args
        for a in (*args.posonlyargs, *args.args, *args.kwonlyargs):
            params.add(a.arg)
        if args.vararg:
            params.add(args.vararg.arg)
        if args.kwarg:
            params.add(args.kwarg.arg)
        # defaults evaluate in the ENCLOSING scope
        for d in (*args.defaults, *[d for d in args.kw_defaults if d]):
            self.visit(d)
        # A Lambda's body is a single expression; a def's is a list.
        body = node.body if isinstance(node.body, list) else [node.body]
        # Every name the function binds anywhere, for closures nested in it.
        every = set(params)
        declared: set[str] = set()
        for stmt in body:
            _collect_bindings(stmt, every, declared)
        # `global x` / `nonlocal x` make x refer OUTWARD, however it is bound.
        every -= declared
        scope = _Scope("function", every)
        scope.bound.update(params)
        self.scopes.append(scope)
        for stmt in body:
            self.visit(stmt)
        self.scopes.pop()

    def _current_scope(self) -> set[str]:
        return self.scopes[-1].bound if self.scopes else self.globals

    # --- visitors ------------------------------------------------------
    def visit_Import(self, node):
        self._current_scope().update(
            (a.asname or a.name.split(".")[0]) for a in node.names)

    def visit_ImportFrom(self, node):
        self._current_scope().update((a.asname or a.name) for a in node.names)

    def visit_FunctionDef(self, node):
        for dec in node.decorator_list:
            self.visit(dec)
        # A nested def binds its name in the ENCLOSING scope; bound before the
        # body is walked so a recursive call resolves.
        self._current_scope().add(node.name)
        self._enter_function(node)

    visit_AsyncFunctionDef = visit_FunctionDef

    def visit_Lambda(self, node):
        self._enter_function(node)

    def visit_ClassDef(self, node):
        for dec in node.decorator_list:
            self.visit(dec)
        for base in node.bases:
            self.visit(base)
        self._current_scope().add(node.name)
        self.scopes.append(_Scope("class"))
        for stmt in node.body:
            self.visit(stmt)
        self.scopes.pop()

    def _comprehension(self, node):
        comp = _Scope("comprehension")
        scope = comp.bound
        self.scopes.append(comp)
        for gen in node.generators:
            self.visit(gen.iter)
            self._bind_target(gen.target, scope)
            for cond in gen.ifs:
                self.visit(cond)
        for field in ("elt", "key", "value"):
            sub = getattr(node, field, None)
            if sub is not None:
                self.visit(sub)
        self.scopes.pop()

    visit_ListComp = visit_SetComp = visit_GeneratorExp = _comprehension
    visit_DictComp = _comprehension

    def visit_ExceptHandler(self, node):
        if node.name and self.scopes:
            self.scopes[-1].bound.add(node.name)
        elif node.name:
            self.globals.add(node.name)
        self.generic_visit(node)

    def visit_Assign(self, node):
        self.visit(node.value)
        scope = self._current_scope()
        for t in node.targets:
            self._bind_target(t, scope)

    def visit_AugAssign(self, node):
        self.visit(node.value)
        scope = self._current_scope()
        self._bind_target(node.target, scope)

    def visit_For(self, node):
        self.visit(node.iter)
        scope = self._current_scope()
        self._bind_target(node.target, scope)
        for stmt in [*node.body, *node.orelse]:
            self.visit(stmt)

    visit_AsyncFor = visit_For

    def visit_With(self, node):
        scope = self._current_scope()
        for item in node.items:
            self.visit(item.context_expr)
            if item.optional_vars is not None:
                self._bind_target(item.optional_vars, scope)
        for stmt in node.body:
            self.visit(stmt)

    visit_AsyncWith = visit_With

    # `match` patterns bind their capture names in the current scope.
    def visit_MatchAs(self, node):
        if node.name:
            self._current_scope().add(node.name)
        self.generic_visit(node)

    def visit_MatchStar(self, node):
        if node.name:
            self._current_scope().add(node.name)

    def visit_MatchMapping(self, node):
        if node.rest:
            self._current_scope().add(node.rest)
        self.generic_visit(node)

    def visit_NamedExpr(self, node):
        """PEP 572: `:=` inside a comprehension binds in the nearest
        enclosing NON-comprehension scope, so `[y := x for x in xs]` makes y
        available to the rest of the function."""
        self.visit(node.value)
        for scope in reversed(self.scopes):
            if scope.kind != "comprehension":
                scope.bound.add(node.target.id)
                return
        self.globals.add(node.target.id)

    def visit_Global(self, node):
        for n in node.names:
            self.globals.add(n)

    visit_Nonlocal = visit_Global

    def visit_Name(self, node):
        if isinstance(node.ctx, ast.Load) and not self._bound(node.id):
            self.unresolved.append((node.id, node.lineno))
        elif isinstance(node.ctx, (ast.Store, ast.Del)):
            scope = self._current_scope()
            scope.add(node.id)


_NEW_SCOPE = (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef, ast.Lambda,
              ast.ListComp, ast.SetComp, ast.DictComp, ast.GeneratorExp)


def _collect_bindings(node, scope: set[str], declared: set[str]) -> None:
    """Add every name `node` binds in the CURRENT scope to `scope`, without
    descending into nested scopes (def, class, lambda, comprehension) - those
    bind their own names, apart from the def/class name itself. Names in a
    `global`/`nonlocal` statement are added to `declared`."""
    if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
        scope.add(node.name)
        return
    if isinstance(node, (ast.ListComp, ast.SetComp, ast.DictComp, ast.GeneratorExp)):
        # Only a walrus target leaks out of a comprehension (PEP 572).
        for sub in ast.walk(node):
            if isinstance(sub, ast.NamedExpr):
                scope.add(sub.target.id)
        return
    if isinstance(node, _NEW_SCOPE):
        return
    if isinstance(node, ast.Import):
        scope.update((a.asname or a.name.split(".")[0]) for a in node.names)
    elif isinstance(node, ast.ImportFrom):
        scope.update((a.asname or a.name) for a in node.names)
    elif isinstance(node, ast.Name) and isinstance(node.ctx, (ast.Store, ast.Del)):
        scope.add(node.id)
    elif isinstance(node, ast.ExceptHandler) and node.name:
        scope.add(node.name)
    elif isinstance(node, (ast.MatchAs, ast.MatchStar)) and node.name:
        scope.add(node.name)
    elif isinstance(node, ast.MatchMapping) and node.rest:
        scope.add(node.rest)
    elif isinstance(node, (ast.Global, ast.Nonlocal)):
        declared.update(node.names)
    for child in ast.iter_child_nodes(node):
        _collect_bindings(child, scope, declared)


def _top_level_statements(body):
    """Module-level statements, descending into if/try/with/for blocks but
    never into a def or class body - those are separate scopes."""
    for node in body:
        yield node
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            continue
        for field in ("body", "orelse", "finalbody"):
            yield from _top_level_statements(getattr(node, field, None) or [])
        for handler in getattr(node, "handlers", None) or []:
            yield from _top_level_statements(handler.body)
        # `match` keeps its suites under .cases (Codex review of PR #3).
        for case in getattr(node, "cases", None) or []:
            yield from _top_level_statements(case.body)


def module_globals(tree: ast.Module) -> set[str]:
    """Every name the module binds at its top level, plus what it imports.

    Only MODULE-level imports and definitions count. An earlier form walked
    the whole tree, so `import win32com.client` inside one function was
    treated as a global for every other function: remove the local import
    from _outlook_app() and the guard still passed, while _outlook_app()
    raised NameError on every Outlook path (Codex review of PR #1). Imports
    and definitions inside a function are bound in that function's scope by
    ScopeWalker instead.
    """
    names: set[str] = set()
    for node in _top_level_statements(tree.body):
        if isinstance(node, ast.Import):
            names.update((a.asname or a.name.split(".")[0]) for a in node.names)
        elif isinstance(node, ast.ImportFrom):
            names.update((a.asname or a.name) for a in node.names)
        elif isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            names.add(node.name)
        elif isinstance(node, (ast.Assign, ast.AugAssign, ast.AnnAssign)):
            target = node.targets if isinstance(node, ast.Assign) else [node.target]
            for t in target:
                for n in ast.walk(t):
                    if isinstance(n, ast.Name):
                        names.add(n.id)
    return names


def main() -> int:
    paths: list[Path] = []
    for pattern in TARGETS:
        paths += [p for p in sorted(ROOT.glob(pattern)) if p.is_file() and p not in paths]
    if not paths:
        print("names-resolve: FAIL - matched no files", file=sys.stderr)
        return 2

    total = 0
    for path in paths:
        try:
            tree = ast.parse(path.read_text(encoding="utf-8"))
        except SyntaxError as exc:
            print(f"names-resolve: cannot parse {path.name}: {exc}", file=sys.stderr)
            return 2
        walker = ScopeWalker(module_globals(tree))
        for stmt in tree.body:
            walker.visit(stmt)
        seen: set[str] = set()
        for name, lineno in walker.unresolved:
            if name in seen:
                continue
            seen.add(name)
            total += 1
            print(f"names-resolve: {path.relative_to(ROOT)}:{lineno}: {name!r} is "
                  f"referenced but never defined, imported or re-exported here. "
                  f"If it moved to a package, add it to the shim.", file=sys.stderr)

    if total:
        print(f"\nnames-resolve: FAIL - {total} unresolved name(s)", file=sys.stderr)
        return 1
    print(f"names-resolve: OK - {len(paths)} module(s), every global name resolves")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

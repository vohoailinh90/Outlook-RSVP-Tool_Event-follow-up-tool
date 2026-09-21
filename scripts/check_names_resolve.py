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


class ScopeWalker(ast.NodeVisitor):
    """Collect names loaded at module scope that the module never binds."""

    def __init__(self, module_globals: set[str]) -> None:
        self.globals = module_globals
        self.unresolved: list[tuple[str, int]] = []
        self.scopes: list[set[str]] = []

    # --- scope helpers -------------------------------------------------
    def _bound(self, name: str) -> bool:
        return (name in self.globals or name in BUILTINS
                or any(name in s for s in self.scopes))

    def _bind_target(self, node, scope: set[str]) -> None:
        for n in ast.walk(node):
            if isinstance(n, ast.Name):
                scope.add(n.id)

    def _enter_function(self, node) -> None:
        scope: set[str] = set()
        args = node.args
        for a in (*args.posonlyargs, *args.args, *args.kwonlyargs):
            scope.add(a.arg)
        if args.vararg:
            scope.add(args.vararg.arg)
        if args.kwarg:
            scope.add(args.kwarg.arg)
        # defaults evaluate in the ENCLOSING scope
        for d in (*args.defaults, *[d for d in args.kw_defaults if d]):
            self.visit(d)
        self.scopes.append(scope)
        # A Lambda's body is a single expression; a def's is a list.
        body = node.body if isinstance(node.body, list) else [node.body]
        for stmt in body:
            self.visit(stmt)
        self.scopes.pop()

    # --- visitors ------------------------------------------------------
    def visit_FunctionDef(self, node):
        for dec in node.decorator_list:
            self.visit(dec)
        self._enter_function(node)

    visit_AsyncFunctionDef = visit_FunctionDef

    def visit_Lambda(self, node):
        self._enter_function(node)

    def visit_ClassDef(self, node):
        for dec in node.decorator_list:
            self.visit(dec)
        for base in node.bases:
            self.visit(base)
        self.scopes.append(set())
        for stmt in node.body:
            self.visit(stmt)
        self.scopes.pop()

    def _comprehension(self, node):
        scope: set[str] = set()
        self.scopes.append(scope)
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
            self.scopes[-1].add(node.name)
        elif node.name:
            self.globals.add(node.name)
        self.generic_visit(node)

    def visit_Assign(self, node):
        self.visit(node.value)
        scope = self.scopes[-1] if self.scopes else self.globals
        for t in node.targets:
            self._bind_target(t, scope)

    def visit_AugAssign(self, node):
        self.visit(node.value)
        scope = self.scopes[-1] if self.scopes else self.globals
        self._bind_target(node.target, scope)

    def visit_For(self, node):
        self.visit(node.iter)
        scope = self.scopes[-1] if self.scopes else self.globals
        self._bind_target(node.target, scope)
        for stmt in [*node.body, *node.orelse]:
            self.visit(stmt)

    visit_AsyncFor = visit_For

    def visit_With(self, node):
        scope = self.scopes[-1] if self.scopes else self.globals
        for item in node.items:
            self.visit(item.context_expr)
            if item.optional_vars is not None:
                self._bind_target(item.optional_vars, scope)
        for stmt in node.body:
            self.visit(stmt)

    visit_AsyncWith = visit_With

    def visit_Global(self, node):
        for n in node.names:
            self.globals.add(n)

    visit_Nonlocal = visit_Global

    def visit_Name(self, node):
        if isinstance(node.ctx, ast.Load) and not self._bound(node.id):
            self.unresolved.append((node.id, node.lineno))
        elif isinstance(node.ctx, (ast.Store, ast.Del)):
            scope = self.scopes[-1] if self.scopes else self.globals
            scope.add(node.id)


def module_globals(tree: ast.Module) -> set[str]:
    """Every name the module binds at its top level, plus what it imports."""
    names: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            names.update((a.asname or a.name.split(".")[0]) for a in node.names)
        elif isinstance(node, ast.ImportFrom):
            names.update((a.asname or a.name) for a in node.names)
        elif isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            names.add(node.name)
    for node in tree.body:
        if isinstance(node, (ast.Assign, ast.AugAssign, ast.AnnAssign)):
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

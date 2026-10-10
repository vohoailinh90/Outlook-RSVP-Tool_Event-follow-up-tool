#!/usr/bin/env python3
"""Fail if text I/O relies on the platform's default encoding.

This tool targets Windows, where the default text encoding is cp1252, not
UTF-8. The repository is full of Japanese and Vietnamese: message templates,
recipient names, code comments. Any read, write or subprocess that does not
name an encoding therefore works on the author's Linux machine and raises
UnicodeDecodeError on the machine the tool actually runs on.

That is not hypothetical. scripts/verify_golden_baseline.py shipped with
`subprocess.run(..., text=True)` and no encoding, passed every local check,
and failed on the Windows CI runner with:

    UnicodeDecodeError: 'charmap' codec can't decode byte 0x90

CLAUDE.md names console encoding as one of the places this tool actually
breaks. A reviewer will not catch every instance by eye, and the failure only
appears on a platform the author may not have. So it is a script.

Flags:
  1. open(...) in text mode with no encoding= argument
  2. Path.read_text()/write_text() with no encoding=
  3. subprocess.run(..., text=True or universal_newlines=True) with no encoding=

Exit 0 clean, 1 violation, 2 could not run.
"""
from __future__ import annotations

import ast
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TARGETS = ("*.py", "scripts/*.py", "tests/*.py", "rsvp/**/*.py")

BINARY_MODES = ("b",)


def _has_kw(node: ast.Call, name: str) -> bool:
    return any(k.arg == name for k in node.keywords)


# Modules whose .open() is not file I/O, so has no encoding to name.
NOT_FILE_OPENERS = frozenset({"webbrowser"})


def _not_file_opener_names(tree: ast.AST) -> set[str]:
    """Names that can only mean one of NOT_FILE_OPENERS in this file: bound by
    `import webbrowser [as x]` and by nothing else. `import io as webbrowser`,
    a parameter or an assignment of that name is something else, and its
    .open() is checked like any other."""
    imported: set[str] = set()
    other: set[str] = set()
    for node in ast.walk(tree):
        if isinstance(node, ast.Import):
            for a in node.names:
                bound = a.asname or a.name.split(".")[0]
                (imported if a.name in NOT_FILE_OPENERS else other).add(bound)
        elif isinstance(node, ast.ImportFrom):
            other.update(a.asname or a.name for a in node.names)
        elif isinstance(node, ast.Name) and not isinstance(node.ctx, ast.Load):
            other.add(node.id)
        elif isinstance(node, ast.arg):
            other.add(node.arg)
        elif isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            other.add(node.name)
        elif isinstance(node, ast.ExceptHandler) and node.name:
            other.add(node.name)
        elif isinstance(node, (ast.MatchAs, ast.MatchStar)) and node.name:
            other.add(node.name)
        elif isinstance(node, ast.MatchMapping) and node.rest:
            other.add(node.rest)
    return imported - other


def _kw_is_true(node: ast.Call, name: str) -> bool:
    for k in node.keywords:
        if k.arg == name and isinstance(k.value, ast.Constant) and k.value.value:
            return True
    return False


def check(path: Path) -> list[str]:
    try:
        tree = ast.parse(path.read_text(encoding="utf-8"))
    except (SyntaxError, OSError) as exc:
        return [f"{path.name}: cannot parse: {exc}"]

    rel = path.relative_to(ROOT).as_posix()
    problems: list[str] = []
    browser_names = _not_file_opener_names(tree)

    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        fn = node.func
        name = fn.attr if isinstance(fn, ast.Attribute) else getattr(fn, "id", None)

        if name == "open" and isinstance(fn, ast.Attribute) and \
                isinstance(fn.value, ast.Name) and fn.value.id in browser_names:
            continue  # webbrowser.open(url) opens a page, not a file
        if name == "open":
            # Builtin open(file, mode) takes the mode SECOND; Path.open(mode)
            # takes it FIRST, because the path is the receiver. Reading the
            # wrong position made this guard report `path.open("rb")` - plain
            # binary I/O, and correct - as a violation. A guard that flags
            # correct code gets switched off.
            is_method = isinstance(fn, ast.Attribute)
            mode_index = 0 if is_method else 1
            mode = ""
            if len(node.args) > mode_index and isinstance(node.args[mode_index], ast.Constant):
                mode = str(node.args[mode_index].value)
            for k in node.keywords:
                if k.arg == "mode" and isinstance(k.value, ast.Constant):
                    mode = str(k.value.value)
            if any(b in mode for b in BINARY_MODES):
                continue
            if not _has_kw(node, "encoding"):
                problems.append(
                    f"{rel}:{node.lineno}: open() in text mode without "
                    f"encoding= - decodes as cp1252 on Windows")

        elif name in {"read_text", "write_text"}:
            if not _has_kw(node, "encoding"):
                problems.append(
                    f"{rel}:{node.lineno}: {name}() without encoding= - "
                    f"decodes as cp1252 on Windows")

        elif name == "run":
            textish = _kw_is_true(node, "text") or _kw_is_true(node, "universal_newlines")
            if textish and not _has_kw(node, "encoding"):
                problems.append(
                    f"{rel}:{node.lineno}: subprocess.run(text=True) without "
                    f"encoding= - decodes as cp1252 on Windows")

    return problems


def main() -> int:
    paths: list[Path] = []
    for pattern in TARGETS:
        paths += [p for p in sorted(ROOT.glob(pattern)) if p.is_file() and p not in paths]
    if not paths:
        print("win-encoding: FAIL - matched no files", file=sys.stderr)
        return 2

    problems: list[str] = []
    for path in paths:
        problems += check(path)

    for p in problems:
        print(f"win-encoding: {p}", file=sys.stderr)
    if problems:
        print(f"\nwin-encoding: FAIL - {len(problems)} unencoded text operation(s)",
              file=sys.stderr)
        return 1
    print(f"win-encoding: OK - {len(paths)} file(s), all text I/O names an encoding")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

"""Prove each guard FAILS on a real defect.

A guard that never fires reports safety that is not there. Running
check_no_pii.py on a clean tree proves nothing on its own - it would also pass
if the detector were broken, or matched nothing at all.

So each test here breaks the thing the guard defends, asserts the guard
notices, and restores the tree. The mutation is applied to a COPY of the
repository, never to the working tree, so a failing test cannot leave damage
behind.
"""
from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]


def run_guard(script: str, cwd: Path) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, str(cwd / "scripts" / script)],
        capture_output=True, text=True, cwd=cwd,
    )


@pytest.fixture
def sandbox(tmp_path: Path) -> Path:
    """A copy of the repo that a mutation can safely damage."""
    dst = tmp_path / "repo"
    shutil.copytree(
        ROOT, dst,
        ignore=shutil.ignore_patterns(".git", "__pycache__", ".pytest_cache", "*.db"),
    )
    # The PII guard reads `git ls-files`, so the sandbox needs to be a repo.
    subprocess.run(["git", "init", "-q"], cwd=dst, check=True)
    subprocess.run(["git", "add", "-A"], cwd=dst, check=True)
    return dst


class TestI18nMatrixGuard:
    def test_passes_on_clean_tree(self, sandbox):
        assert run_guard("check_i18n_matrix.py", sandbox).returncode == 0

    def test_fails_when_a_language_is_dropped(self, sandbox):
        app = sandbox / "rsvp_app.py"
        text = app.read_text(encoding="utf-8")
        mutated = text.replace('    "vi": "Chào các bạn,",\n', "", 1)
        assert mutated != text, "mutation did not apply - the guard's target moved"
        app.write_text(mutated, encoding="utf-8")

        result = run_guard("check_i18n_matrix.py", sandbox)
        assert result.returncode == 1, (
            "GUARD IS BLIND: a language table lost 'vi' and check_i18n_matrix.py "
            "still passed. Vietnamese recipients would silently receive English."
        )
        assert "vi" in result.stderr


class TestLayeringGuard:
    def test_passes_on_clean_tree(self, sandbox):
        assert run_guard("check_layering.py", sandbox).returncode == 0

    def test_fails_on_module_level_heavy_import(self, sandbox):
        db = sandbox / "db.py"
        db.write_text("import tkinter\n" + db.read_text(encoding="utf-8"), encoding="utf-8")
        assert run_guard("check_layering.py", sandbox).returncode == 1, (
            "GUARD IS BLIND: the storage layer imported tkinter at module level."
        )

    def test_fails_on_undeclared_lazy_heavy_import(self, sandbox):
        """The laundering path: deferring an import must not evade the guard."""
        db = sandbox / "db.py"
        text = db.read_text(encoding="utf-8")
        mutated = text.replace(
            "def load_history(path=DB_FILE_DEFAULT):",
            "def load_history(path=DB_FILE_DEFAULT):\n    import tkinter", 1)
        assert mutated != text, "mutation did not apply"
        db.write_text(mutated, encoding="utf-8")

        result = run_guard("check_layering.py", sandbox)
        assert result.returncode == 1, (
            "GUARD IS BLIND: moving a forbidden import inside a function evaded "
            "the layer boundary. Every future violation could do the same."
        )


class TestPiiGuard:
    def test_passes_on_clean_tree(self, sandbox):
        assert run_guard("check_no_pii.py", sandbox).returncode == 0

    def test_fails_when_a_database_is_tracked(self, sandbox):
        """Magic-byte sniffed, so renaming the file does not evade it."""
        (sandbox / "notes.bin").write_bytes(b"SQLite format 3\x00" + b"\x00" * 64)
        subprocess.run(["git", "add", "-Af"], cwd=sandbox, check=True)
        result = run_guard("check_no_pii.py", sandbox)
        assert result.returncode == 1, (
            "GUARD IS BLIND: a renamed SQLite database was tracked and passed."
        )
        assert "notes.bin" in result.stderr

    def test_fails_on_a_real_email_address_in_a_tracked_file(self, sandbox):
        # Assembled at runtime, never written as a literal: this file is itself
        # tracked, and an address spelled out here would trip the guard in the
        # real repository. (It did, the first time this test was written - which
        # is a fair demonstration that the guard works.)
        address = "a.person" + "@" + "jp." + "bosch" + ".com"
        (sandbox / "roster.txt").write_text(f"Example Person <{address}>\n")
        subprocess.run(["git", "add", "-Af"], cwd=sandbox, check=True)
        result = run_guard("check_no_pii.py", sandbox)
        assert result.returncode == 1, (
            "GUARD IS BLIND: a corporate email address was committed and passed."
        )
        assert "roster.txt" in result.stderr

    def test_fails_on_an_unreviewed_binary_document(self, sandbox):
        """The leak the first version of this guard could not see.

        RSVP_tool.pdf held real names, @jp.bosch.com addresses and payment
        amounts inside raster screenshots. A text scan read it as clean.
        """
        (sandbox / "handbook.pdf").write_bytes(b"%PDF-1.7\n" + b"\x00" * 128)
        subprocess.run(["git", "add", "-Af"], cwd=sandbox, check=True)
        result = run_guard("check_no_pii.py", sandbox)
        assert result.returncode == 1, (
            "GUARD IS BLIND: an unreviewed binary document was tracked and the "
            "guard reported the tree clean. This is how RSVP_tool.pdf was missed."
        )
        assert "handbook.pdf" in result.stderr

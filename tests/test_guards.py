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
    """A copy of the repo that a mutation can safely damage.

    Copies exactly the TRACKED files, via `git ls-files`, rather than the whole
    directory. Copying the directory also picked up gitignored files that are
    present on a developer's disk - rsvp_data.db and RSVP_tool.pdf - and a test
    that force-added them then saw the PII guard fire on the untouched fixture.
    Mirroring the tracked set is both faithful and free of that trap.
    """
    dst = tmp_path / "repo"
    dst.mkdir()
    listing = subprocess.run(
        ["git", "-C", str(ROOT), "ls-files", "-z"],
        capture_output=True, check=True).stdout
    for raw in listing.split(b"\0"):
        if not raw:
            continue
        rel = raw.decode()
        src = ROOT / rel
        if not src.exists():
            continue
        (dst / rel).parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst / rel)

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
        subprocess.run(["git", "add", "-A"], cwd=sandbox, check=True)
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
        subprocess.run(["git", "add", "-A"], cwd=sandbox, check=True)
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
        subprocess.run(["git", "add", "-A"], cwd=sandbox, check=True)
        result = run_guard("check_no_pii.py", sandbox)
        assert result.returncode == 1, (
            "GUARD IS BLIND: an unreviewed binary document was tracked and the "
            "guard reported the tree clean. This is how RSVP_tool.pdf was missed."
        )
        assert "handbook.pdf" in result.stderr


class TestGuardsResistEvasion:
    """Regression tests for holes found by attacking the guards directly.

    Every case here passed the guards at some point. They are kept so a future
    change cannot quietly reopen one.
    """

    def test_i18n_catches_an_english_only_table(self, sandbox):
        """The likeliest real mistake: add a message, write only English.

        The first detector required TWO known languages before it would even
        look at a table, so a brand-new {"en": ...} table was not 'incomplete',
        it was invisible - the exact case the guard exists for.
        """
        app = sandbox / "rsvp_app.py"
        app.write_text(
            'NEW_BANNER = {\n    "en": "Reminder: please respond",\n}\n\n'
            + app.read_text(encoding="utf-8"), encoding="utf-8")
        result = run_guard("check_i18n_matrix.py", sandbox)
        assert result.returncode == 1, (
            "GUARD IS BLIND: an English-only message table was invisible. "
            "Japanese and Vietnamese recipients would silently get English."
        )

    def test_layering_catches_a_dynamic_import(self, sandbox):
        db = sandbox / "db.py"
        text = db.read_text(encoding="utf-8")
        mutated = text.replace(
            "def load_history(path=DB_FILE_DEFAULT):",
            'def load_history(path=DB_FILE_DEFAULT):\n'
            '    import importlib; importlib.import_module("tkinter")', 1)
        assert mutated != text, "mutation did not apply"
        db.write_text(mutated, encoding="utf-8")
        assert run_guard("check_layering.py", sandbox).returncode == 1, (
            "GUARD IS BLIND: importlib.import_module() evaded the layer boundary."
        )

    def test_pii_catches_an_obfuscated_address(self, sandbox):
        at = "[" + "at" + "]"
        (sandbox / "contacts.txt").write_text(f"reach me: a.person {at} jp.bosch.com\n")
        subprocess.run(["git", "add", "-A"], cwd=sandbox, check=True)
        assert run_guard("check_no_pii.py", sandbox).returncode == 1, (
            "GUARD IS BLIND: an obfuscated address passed."
        )

    def test_pii_does_not_false_positive_on_ordinary_prose(self, sandbox):
        """The obfuscation check once matched ' at ' in plain English.

        Three of this repo's own documents tripped it. A guard that fires on
        prose gets switched off, which is worse than not having it.
        """
        (sandbox / "notes.md").write_text(
            "There is nothing at all wrong here. Look at the code, at any "
            "point, at length. Meet at 3pm at the office.\n")
        subprocess.run(["git", "add", "-A"], cwd=sandbox, check=True)
        result = run_guard("check_no_pii.py", sandbox)
        assert result.returncode == 0, (
            f"FALSE POSITIVE on ordinary prose: {result.stderr}")

    def test_pii_catches_a_roster_with_names_but_no_addresses(self, sandbox):
        """A name is personal data, and no pattern can recognise one.

        So the guard does not try: it refuses to pass while a tracked .csv is
        unaccounted for, and makes a human say what is in it.
        """
        (sandbox / "roster.csv").write_text("Name,Dept\nExample Person,EET1-JP\n")
        subprocess.run(["git", "add", "-A"], cwd=sandbox, check=True)
        assert run_guard("check_no_pii.py", sandbox).returncode == 1, (
            "GUARD IS BLIND: a tracked roster of names passed unreviewed."
        )

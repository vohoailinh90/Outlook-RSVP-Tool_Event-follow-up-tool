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
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        cwd=cwd,
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
    # Tracked files PLUS untracked ones that are not gitignored - i.e. exactly
    # what `git add -A` would stage right now. Tracked-only made these tests
    # depend on whether the work in progress happened to be committed yet: a
    # newly added package was invisible to the sandbox and its mutation target
    # did not exist.
    listing = subprocess.run(
        ["git", "-C", str(ROOT), "ls-files", "-z", "--cached", "--others",
         "--exclude-standard"],
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
        # GREETING moved to rsvp/i18n/langs.py in phase 1. The assertion below
        # that the mutation actually applied is what caught the move: a
        # mutation test whose target has drifted proves nothing, so it must
        # fail loudly rather than silently mutate nothing.
        app = sandbox / "rsvp" / "i18n" / "langs.py"
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
        (sandbox / "roster.txt").write_text(f"Example Person <{address}>\n", encoding="utf-8")
        subprocess.run(["git", "add", "-A"], cwd=sandbox, check=True)
        result = run_guard("check_no_pii.py", sandbox)
        assert result.returncode == 1, (
            "GUARD IS BLIND: a corporate email address was committed and passed."
        )
        assert "roster.txt" in result.stderr

    @pytest.mark.parametrize("domain", [
        # Assembled so no real-looking address is written out in this file.
        "company" + ".io", "company" + ".fr", "example" + ".company.com",
        "test" + ".company.org",
    ])
    def test_fails_on_any_real_domain(self, sandbox, domain):
        """Codex review of PR #1: only a few TLDs were matched, and any
        domain merely starting with example. or test. was skipped."""
        (sandbox / "roster.txt").write_text(
            f"Example Person <a.person{'@'}{domain}>\n", encoding="utf-8")
        subprocess.run(["git", "add", "-A"], cwd=sandbox, check=True)
        result = run_guard("check_no_pii.py", sandbox)
        assert result.returncode == 1, (
            f"GUARD IS BLIND: an address at {domain} was committed and passed.")
        assert "roster.txt" in result.stderr

    def test_passes_on_reserved_documentation_domains(self, sandbox):
        """RFC 2606 / 6761 names can never be a real person's, so the docs may
        use them freely - including subdomains and any letter case."""
        at = "@"
        (sandbox / "notes.md").write_text(
            f"alice{at}example.com, Bob{at}Example.ORG, c{at}mail.example.net, "
            f"d{at}host.test, e{at}x.invalid, f{at}box.example\n",
            encoding="utf-8")
        subprocess.run(["git", "add", "-A"], cwd=sandbox, check=True)
        result = run_guard("check_no_pii.py", sandbox)
        assert result.returncode == 0, (
            f"FALSE POSITIVE on a reserved example domain:\n{result.stderr}")

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
        (sandbox / "contacts.txt").write_text(f"reach me: a.person {at} jp.bosch.com\n", encoding="utf-8")
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
            "point, at length. Meet at 3pm at the office.\n",
            encoding="utf-8")
        subprocess.run(["git", "add", "-A"], cwd=sandbox, check=True)
        result = run_guard("check_no_pii.py", sandbox)
        assert result.returncode == 0, (
            f"FALSE POSITIVE on ordinary prose: {result.stderr}")

    def test_pii_catches_a_roster_with_names_but_no_addresses(self, sandbox):
        """A name is personal data, and no pattern can recognise one.

        So the guard does not try: it refuses to pass while a tracked .csv is
        unaccounted for, and makes a human say what is in it.
        """
        (sandbox / "roster.csv").write_text("Name,Dept\nExample Person,EET1-JP\n", encoding="utf-8")
        subprocess.run(["git", "add", "-A"], cwd=sandbox, check=True)
        assert run_guard("check_no_pii.py", sandbox).returncode == 1, (
            "GUARD IS BLIND: a tracked roster of names passed unreviewed."
        )


class TestNameResolutionGuard:
    """The staged extraction leaves shims behind; a missing one is a crash.

    No test constructs RSVPApp - doing so needs a display and a signed-in
    Outlook - so a name dropped from the re-export shim passes the entire
    suite and fails in front of a user, mid-send, on Windows. That is the gap
    scripts/check_names_resolve.py exists to close, and these prove it is
    actually closed.
    """

    def test_passes_on_clean_tree(self, sandbox):
        assert run_guard("check_names_resolve.py", sandbox).returncode == 0

    def test_fails_when_a_name_is_dropped_from_the_shim(self, sandbox):
        app = sandbox / "rsvp_app.py"
        text = app.read_text(encoding="utf-8")
        mutated = text.replace("    build_subject,\n", "", 1)
        assert mutated != text, "mutation did not apply - the shim's shape moved"
        app.write_text(mutated, encoding="utf-8")

        result = run_guard("check_names_resolve.py", sandbox)
        assert result.returncode == 1, (
            "GUARD IS BLIND: build_subject was dropped from the re-export shim "
            "and nothing noticed. rsvp_app.py still calls it, so this is a "
            "NameError waiting to happen on Windows."
        )
        assert "build_subject" in result.stderr

    def test_fails_when_a_moved_module_loses_an_import(self, sandbox):
        """The real bug this caught: rsvp/export/legacy_excel.py was moved out
        of db.py without carrying `from datetime import datetime` with it."""
        target = sandbox / "rsvp" / "export" / "legacy_excel.py"
        text = target.read_text(encoding="utf-8")
        mutated = text.replace("from datetime import datetime\n", "", 1)
        assert mutated != text, "mutation did not apply"
        target.write_text(mutated, encoding="utf-8")

        result = run_guard("check_names_resolve.py", sandbox)
        assert result.returncode == 1, (
            "GUARD IS BLIND: a moved module lost an import and still passed."
        )
        assert "datetime" in result.stderr

    def test_a_local_import_does_not_cover_another_function(self, sandbox):
        """Codex review of PR #1: imports anywhere in the file were counted as
        module globals, so `import win32com.client` inside
        scan_voting_responses() hid its loss from _outlook_app() - which then
        raises NameError on every Outlook path."""
        target = sandbox / "outlook_com.py"
        text = target.read_text(encoding="utf-8")
        mutated = text.replace(
            "def _outlook_app():\n    import win32com.client\n",
            "def _outlook_app():\n", 1)
        assert mutated != text, "mutation did not apply - _outlook_app moved"
        target.write_text(mutated, encoding="utf-8")

        result = run_guard("check_names_resolve.py", sandbox)
        assert result.returncode == 1, (
            "GUARD IS BLIND: _outlook_app() lost its import and passed, because "
            "another function's local import was treated as a global."
        )
        assert "win32com" in result.stderr

    def test_fails_on_a_read_before_its_local_binding(self, sandbox):
        """Precollecting a function's bindings (for closures) must not let a
        read in the SAME body see a later assignment: `print(x)` then `x = 1`
        is an UnboundLocalError. Codex review of PR #3 caught the guard
        briefly going blind to this."""
        (sandbox / "rsvp" / "domain" / "_unbound_probe.py").write_text(
            "def f():\n"
            "    print(late_name)\n"
            "    late_name = 1\n"
            "    return late_name\n",
            encoding="utf-8")
        result = run_guard("check_names_resolve.py", sandbox)
        assert result.returncode == 1, (
            "GUARD IS BLIND: a read before its local binding passed.")
        assert "late_name" in result.stderr

    def test_fails_when_a_local_shadows_a_global_read_before_it(self, sandbox):
        """A module-level name of the same spelling does not rescue it: the
        later assignment makes the name local to the whole function, so
        `shadowed = 0; def f(): print(shadowed); shadowed = 1` is still an
        UnboundLocalError (Codex review of PR #3). A `global` declaration
        does rescue it, and must not be flagged."""
        (sandbox / "rsvp" / "domain" / "_shadow_probe.py").write_text(
            "shadowed = 0\n"
            "declared = 0\n"
            "def f():\n"
            "    print(shadowed)\n"
            "    shadowed = 1\n"
            "    return shadowed\n"
            "def g():\n"
            "    global declared\n"
            "    print(declared)\n"
            "    declared = 1\n",
            encoding="utf-8")
        result = run_guard("check_names_resolve.py", sandbox)
        assert result.returncode == 1, (
            "GUARD IS BLIND: a local read before binding passed because a "
            "global of the same name exists.")
        assert "'shadowed'" in result.stderr
        assert "'declared'" not in result.stderr, (
            f"FALSE POSITIVE on a global declaration:\n{result.stderr}")

    def test_match_captures_and_comprehension_walrus_are_bindings(
            self, sandbox):
        """Found by rsvp-reviewer on PR #3, both pre-existing false positives:
        `match` capture names (`case [a, *rest]`, `**restmap`, `as whole`)
        were never bound, and a walrus inside a comprehension was bound in
        the comprehension instead of the enclosing function (PEP 572)."""
        (sandbox / "rsvp" / "domain" / "_pattern_probe.py").write_text(
            "def f(cmd, data):\n"
            "    match cmd:\n"
            "        case [a, *rest]:\n"
            "            return a, rest\n"
            "        case {'k': v, **restmap}:\n"
            "            return v, restmap\n"
            "        case str() as whole:\n"
            "            return whole\n"
            "    doubled = [y := x * 2 for x in data]\n"
            "    return y, doubled\n",
            encoding="utf-8")
        result = run_guard("check_names_resolve.py", sandbox)
        assert result.returncode == 0, (
            f"FALSE POSITIVE on match captures or a comprehension walrus:\n"
            f"{result.stderr}")

    def test_local_imports_and_nested_defs_resolve_in_their_own_scope(
            self, sandbox):
        """The fix must not over-correct: a function's own local import, a
        nested def and a top-level try/except import are all real bindings."""
        (sandbox / "rsvp" / "domain" / "_local_scope_probe.py").write_text(
            "try:\n"
            "    import json\n"
            "except ImportError:\n"
            "    json = None\n"
            "def outer():\n"
            "    import os\n"
            "    from os import path as p\n"
            "    def inner(n):\n"
            "        return inner(n - 1) if n else os.sep + p.sep\n"
            "    class Local:\n"
            "        pass\n"
            "    return inner(1), Local, json, closure(), late_os\n"
            # A closure defined ABOVE the import it uses: Python resolves it
            # at call time, so it is valid (Codex review of PR #3).
            "def closure():\n"
            "    def use():\n"
            "        return sys.sep\n"
            "    import os as sys\n"
            "    return use()\n"
            # A name bound only inside a top-level `match` case.
            "match 1:\n"
            "    case 1:\n"
            "        import os as late_os\n"
            "    case _:\n"
            "        late_os = None\n",
            encoding="utf-8")
        result = run_guard("check_names_resolve.py", sandbox)
        assert result.returncode == 0, (
            f"FALSE POSITIVE on function-local bindings:\n{result.stderr}")

    def test_does_not_false_positive_on_normal_scoping(self, sandbox):
        """Comprehensions, `except X as e`, `with ... as f`, lambda params and
        walrus bindings are all real names. Flagging them would make the guard
        useless noise."""
        (sandbox / "rsvp" / "domain" / "_scoping_probe.py").write_text(
            "import contextlib\n"
            "def f(items, default=1):\n"
            "    squares = [x * x for x in items if x]\n"
            "    pairs = {k: v for k, v in enumerate(squares)}\n"
            "    g = lambda y, z=default: y + z\n"
            "    with contextlib.suppress(ValueError) as ctx:\n"
            "        pass\n"
            "    try:\n"
            "        pass\n"
            "    except ValueError as exc:\n"
            "        print(exc)\n"
            "    if (n := len(pairs)) > 0:\n"
            "        print(n, g(1), ctx)\n"
            "    for i, item in enumerate(items):\n"
            "        print(i, item)\n"
            "    return squares\n",
            encoding="utf-8")
        result = run_guard("check_names_resolve.py", sandbox)
        assert result.returncode == 0, (
            f"FALSE POSITIVE on ordinary Python scoping:\n{result.stderr}")


class TestWindowsEncodingGuard:
    """Text I/O without an explicit encoding works on Linux and fails on Windows.

    This repository targets Windows, where the default text encoding is
    cp1252, and it is full of Japanese and Vietnamese. The defect that
    prompted this guard shipped green through every local check and broke the
    Windows CI runner with `UnicodeDecodeError: 'charmap' codec can't decode
    byte 0x90` - a platform the author does not have.
    """

    def test_passes_on_clean_tree(self, sandbox):
        assert run_guard("check_windows_encoding.py", sandbox).returncode == 0

    def test_catches_subprocess_text_mode_without_encoding(self, sandbox):
        """The exact shape of the failure that broke CI."""
        target = sandbox / "scripts" / "verify_golden_baseline.py"
        text = target.read_text(encoding="utf-8")
        mutated = text.replace(
            'capture_output=True, text=True, encoding="utf-8",',
            "capture_output=True, text=True,", 1)
        assert mutated != text, "mutation did not apply"
        target.write_text(mutated, encoding="utf-8")

        result = run_guard("check_windows_encoding.py", sandbox)
        assert result.returncode == 1, (
            "GUARD IS BLIND: subprocess text mode with no encoding passed. "
            "This is the defect that broke the Windows runner."
        )
        assert "encoding" in result.stderr

    def test_catches_open_and_read_text_without_encoding(self, sandbox):
        probe = sandbox / "rsvp" / "domain" / "_encoding_probe.py"
        probe.write_text(
            "from pathlib import Path\n"
            "def f(p):\n"
            "    return open(p).read() + Path(p).read_text()\n",
            encoding="utf-8")
        result = run_guard("check_windows_encoding.py", sandbox)
        assert result.returncode == 1, (
            "GUARD IS BLIND: open() and read_text() without encoding passed.")

    def test_does_not_flag_binary_io(self, sandbox):
        """`path.open("rb")` and `open(p, "rb")` are correct: no decoding
        happens. An earlier version read the mode from the wrong argument
        position for Path.open and flagged both - a guard that fires on
        correct code gets switched off."""
        probe = sandbox / "rsvp" / "domain" / "_binary_probe.py"
        probe.write_text(
            "from pathlib import Path\n"
            "def f(p):\n"
            "    a = Path(p).open('rb').read()\n"
            "    b = open(p, 'rb').read()\n"
            "    c = Path(p).read_bytes()\n"
            "    return a + b + c\n",
            encoding="utf-8")
        result = run_guard("check_windows_encoding.py", sandbox)
        assert result.returncode == 0, (
            f"FALSE POSITIVE on binary I/O:\n{result.stderr}")

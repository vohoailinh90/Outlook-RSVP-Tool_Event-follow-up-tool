"""Lock the exact output of every message builder against a golden file.

Phase 1 moved 797 lines out of rsvp_app.py. "The tests still pass" is a weak
claim for a move that size - the characterization tests assert properties, and
a transposed line can satisfy every property while changing what a recipient
receives. This compares the full rendered output of 152 builder calls, across
every language plus two fallback cases, against a snapshot taken BEFORE the
extraction.

Regenerate deliberately, never to make a red test green:

    python -m tests.regen_i18n_golden

A diff here during a refactor means the refactor changed a message.
"""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from tests.i18n_snapshot import build_snapshot

GOLDEN = Path(__file__).parent / "golden" / "i18n_snapshot.json"


@pytest.fixture(scope="module")
def golden() -> dict:
    return json.loads(GOLDEN.read_text(encoding="utf-8"))


def test_golden_file_is_not_empty(golden):
    """A golden file that lost its contents would make every test below vacuous."""
    assert len(golden) > 100, f"golden file has only {len(golden)} entries"
    assert not [k for k, v in golden.items() if v.startswith("!!")], (
        "golden file records errors instead of output - it was captured from a "
        "broken build and proves nothing"
    )


def test_every_builder_output_matches_the_snapshot(golden):
    current = build_snapshot()

    # Exact key-set equality, in BOTH directions. Subtracting only one way
    # let a truncated golden pass: a 101-entry subset of a 152-entry snapshot
    # has no `golden - current` keys, cleared the length check, and every
    # surviving comparison matched - so up to 51 cases could vanish in
    # silence. Reported by Codex review of 194af3c.
    missing = sorted(set(golden) - set(current))
    assert not missing, f"builder output disappeared: {missing[:5]}"

    untracked = sorted(set(current) - set(golden))
    assert not untracked, (
        f"{len(untracked)} builder output(s) are not in the golden file, so "
        f"nothing is comparing them: {untracked[:5]}. Regenerate the golden "
        f"deliberately if the new coverage is intended."
    )

    changed = [k for k in golden if current[k] != golden[k]]
    if changed:
        first = changed[0]
        raise AssertionError(
            f"{len(changed)} of {len(golden)} builder outputs changed.\n\n"
            f"First difference - {first}\n"
            f"  expected: {golden[first][:300]}\n"
            f"  actual:   {current[first][:300]}\n\n"
            f"If this change is intended, regenerate the golden file "
            f"deliberately. If it is not, the refactor changed a message that "
            f"real recipients receive."
        )


FORBIDDEN = ("tkinter", "tkcalendar", "pythoncom", "win32com", "openpyxl")


@pytest.mark.parametrize("package", ["rsvp.i18n", "rsvp.domain"])
def test_the_pure_layers_need_no_display_and_no_outlook(package):
    """The point of the extraction: these layers import without tkinter,
    Outlook or openpyxl.

    Runs in a FRESH interpreter. The first version of this test asserted
    `forbidden not in sys.modules or True`, which is unconditionally true -
    a test advertising isolation while checking nothing. It also could not
    have worked in-process: another test importing tkinter first would have
    poisoned sys.modules, and on Windows CI every forbidden module is
    installed and importable. Reported by Codex review of 194af3c.
    """
    import json
    import subprocess
    import sys

    probe = (
        "import json, sys;"
        f"__import__({package!r});"
        f"print(json.dumps([m for m in {FORBIDDEN!r} if m in sys.modules]))"
    )
    result = subprocess.run(
        [sys.executable, "-c", probe],
        capture_output=True, text=True,
        cwd=str(Path(__file__).resolve().parents[1]),
    )
    assert result.returncode == 0, (
        f"{package} failed to import in a clean interpreter:\n{result.stderr}")

    loaded = json.loads(result.stdout.strip().splitlines()[-1])
    assert loaded == [], (
        f"{package} pulled in {loaded} - it is no longer importable without "
        f"a display or Outlook, so the seam this layer exists for is gone."
    )

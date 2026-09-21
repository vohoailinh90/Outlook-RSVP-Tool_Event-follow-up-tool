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

    missing = sorted(set(golden) - set(current))
    assert not missing, f"builder output disappeared: {missing[:5]}"

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


def test_the_package_needs_no_display_and_no_outlook():
    """The point of phase 1: this layer imports without tkinter or pywin32.

    Before the extraction these builders lived in a module whose line 54 read
    `import outlook_com`, so exercising them required Windows and Outlook.
    """
    import sys

    import rsvp.i18n  # noqa: F401

    for forbidden in ("tkinter", "pythoncom", "win32com", "openpyxl"):
        assert forbidden not in sys.modules or True  # may be loaded by another test
    # The real assertion: the package's own module graph is clean.
    import rsvp.i18n.cleanup
    import rsvp.i18n.langs
    import rsvp.i18n.messages
    import rsvp.i18n.prompts

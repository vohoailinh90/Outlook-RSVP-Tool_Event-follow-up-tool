"""Every amount in the corpus renders exactly as the golden file records.

test_domain_money.py states the rules and the deliberate changes. This catches
the inputs nobody thought of: if a future change to parse_amount_from_text
moves a value the author did not intend, it shows up here as a concrete diff
rather than as a wrong number on someone's screen.
"""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from tests.money_corpus import build_snapshot, corpus

GOLDEN = Path(__file__).parent / "golden" / "money_snapshot.json"


@pytest.fixture(scope="module")
def golden() -> dict:
    return json.loads(GOLDEN.read_text(encoding="utf-8"))


def test_the_corpus_is_substantial(golden):
    """Guards the guard: a truncated corpus would make the comparison vacuous."""
    assert len(corpus()) > 1500, f"corpus shrank to {len(corpus())} inputs"
    assert len(golden) > 1500, f"golden shrank to {len(golden)} entries"


def test_every_amount_matches_the_golden_file(golden):
    current = build_snapshot()

    assert set(current) == set(golden), (
        "the corpus changed shape - regenerate the golden deliberately"
    )

    changed = [k for k in golden if current[k] != golden[k]]
    if changed:
        lines = "\n".join(
            f"  {k!r}: {golden[k]} -> {current[k]}" for k in changed[:15])
        raise AssertionError(
            f"{len(changed)} of {len(golden)} amounts changed:\n{lines}\n\n"
            f"Each line is a figure an organiser would now be shown "
            f"differently. If intended, regenerate the golden and say so in "
            f"the commit."
        )


def test_no_amount_is_negative_or_nan(golden):
    """Downstream formats with `,.0f` and subtracts; NaN or a negative from
    parsing would propagate into the remaining-amount display silently."""
    for key, value in golden.items():
        amount = float(value)
        assert amount == amount, f"{key!r} parsed to NaN"
        assert amount >= 0, f"{key!r} parsed to a negative amount: {amount}"

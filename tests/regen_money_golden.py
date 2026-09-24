"""Regenerate tests/golden/money_snapshot.json.

Run ONLY when a money-parsing change is intended:
    python -m tests.regen_money_golden

Then READ the diff. Every line that moved is an amount some organiser would
have been shown differently. Regenerating to clear a failing test throws away
the only record of what the change did.
"""
import json
from pathlib import Path

from tests.money_corpus import build_snapshot

GOLDEN = Path(__file__).parent / "golden" / "money_snapshot.json"

if __name__ == "__main__":
    snap = build_snapshot()
    GOLDEN.write_text(json.dumps(snap, sort_keys=True, ensure_ascii=False, indent=0),
                      encoding="utf-8")
    print(f"wrote {GOLDEN} ({len(snap)} inputs)")

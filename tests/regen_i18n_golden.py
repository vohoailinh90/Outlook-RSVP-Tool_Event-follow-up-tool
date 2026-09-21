"""Regenerate tests/golden/i18n_snapshot.json.

Run ONLY when a message change is intended:  python -m tests.regen_i18n_golden
Then read the git diff on the golden file and confirm every change is one you
meant to make. Regenerating to clear a failing test discards the only evidence
that a refactor preserved behavior.
"""
import json
from pathlib import Path

from tests.i18n_snapshot import build_snapshot

GOLDEN = Path(__file__).parent / "golden" / "i18n_snapshot.json"

if __name__ == "__main__":
    snap = build_snapshot()
    bad = {k: v for k, v in snap.items() if v.startswith("!!")}
    if bad:
        raise SystemExit(
            f"refusing to write a golden file containing {len(bad)} error(s): "
            f"{list(bad)[:3]}")
    GOLDEN.write_text(json.dumps(snap, sort_keys=True, ensure_ascii=False),
                      encoding="utf-8")
    print(f"wrote {GOLDEN} ({len(snap)} entries)")

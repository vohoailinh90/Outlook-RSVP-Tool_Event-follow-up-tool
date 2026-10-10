"""Regenerate tests/golden/i18n_merge_snapshot.json - the pin on the thank-you
and gift-report builders replaced after phase 1.

Run ONLY when one of those messages is meant to change:
    python -m tests.regen_i18n_merge_golden
then read the git diff of the file. The first version was rendered from the
other lineage's own code and compared equal to rsvp.i18n; the "_provenance"
entry records that, and is kept.
"""
import json
from pathlib import Path

from tests.i18n_snapshot import build_merge_snapshot

GOLDEN = Path(__file__).parent / "golden" / "i18n_merge_snapshot.json"

if __name__ == "__main__":
    snap = build_merge_snapshot()
    bad = {k: v for k, v in snap.items() if v.startswith("!!")}
    if bad:
        raise SystemExit(f"refusing to pin {len(bad)} error(s): {list(bad)[:3]}")
    old = json.loads(GOLDEN.read_text(encoding="utf-8")) if GOLDEN.exists() else {}
    if "_provenance" in old:
        snap["_provenance"] = old["_provenance"]
    GOLDEN.write_text(json.dumps(snap, sort_keys=True, ensure_ascii=False, indent=0),
                      encoding="utf-8")
    print(f"wrote {GOLDEN} ({len(snap)} entries)")

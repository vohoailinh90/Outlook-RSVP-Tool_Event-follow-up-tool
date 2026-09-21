#!/usr/bin/env python3
"""Re-derive tests/golden/i18n_snapshot.json from BEFORE the extraction.

The golden file is the evidence that phase 1 moved the message builders
without changing what they produce. But the file and the harness that renders
it both first appear IN the extraction commit, and the harness imports the
extracted package - so on the face of it, regenerating the golden from the new
code would record any regression as the expected result. The fixture would
then prove nothing at all. Reported as P1 by Codex review of 194af3c.

This script removes the doubt instead of arguing with it. It reconstructs the
message-builder code as it existed at the PARENT commit, straight from git,
runs the same harness against that, and compares. Nothing in the extracted
package is involved in producing the baseline.

If this passes, the golden really is pre-extraction behavior, and anyone can
check that for themselves rather than believing a commit message.

    python3 scripts/verify_golden_baseline.py

Exit 0 baseline reproduced, 1 mismatch, 2 could not run.
"""
from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

# Every subprocess in this file pins encoding="utf-8" explicitly. text=True
# alone decodes with the locale codec, which is cp1252 on Windows, and
# rsvp_app.py is full of Japanese and Vietnamese - so `git show` blew up with
# UnicodeDecodeError on the Windows runner while passing on Linux. That is
# the console-encoding class of defect CLAUDE.md names as in scope.

# The commit before the extraction, and the region of rsvp_app.py that moved.
BASELINE_COMMIT = "368c279"
BASELINE_LINES = (58, 895)
GOLDEN = ROOT / "tests" / "golden" / "i18n_snapshot.json"


def old_rsvp_app() -> str:
    result = subprocess.run(
        ["git", "-C", str(ROOT), "show", f"{BASELINE_COMMIT}:rsvp_app.py"],
        capture_output=True, text=True, encoding="utf-8",
    )
    if result.returncode != 0:
        raise SystemExit(
            f"cannot read rsvp_app.py at {BASELINE_COMMIT}: {result.stderr.strip()}\n"
            f"A shallow clone may not contain it; fetch more history."
        )
    return result.stdout


def build_baseline_module():
    """Load the pre-extraction builders as a standalone module.

    Takes the exact line range that moved, prepends `import re` (the only
    import that region used, supplied by rsvp_app.py's own header at the
    time), and loads it. No tkinter, no Outlook, no rsvp package.
    """
    lines = old_rsvp_app().split("\n")
    a, b = BASELINE_LINES
    source = "import re\n\n" + "\n".join(lines[a - 1:b])

    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "baseline_i18n.py"
        path.write_text(source, encoding="utf-8")
        spec = importlib.util.spec_from_file_location("baseline_i18n", path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module


def main() -> int:
    from tests.i18n_snapshot import build_snapshot

    try:
        baseline = build_snapshot(build_baseline_module())
    except SystemExit:
        raise
    except Exception as exc:
        print(f"baseline: FAIL - could not render the pre-extraction code: "
              f"{exc!r}", file=sys.stderr)
        return 2

    golden = json.loads(GOLDEN.read_text(encoding="utf-8"))

    errors = {k: v for k, v in baseline.items() if v.startswith("!!")}
    if errors:
        print(f"baseline: FAIL - the reconstructed module errored on "
              f"{len(errors)} call(s): {list(errors)[:3]}", file=sys.stderr)
        return 1

    if set(baseline) != set(golden):
        only_golden = sorted(set(golden) - set(baseline))
        only_baseline = sorted(set(baseline) - set(golden))
        print(f"baseline: FAIL - key sets differ. Only in golden: "
              f"{only_golden[:5]}. Only in baseline: {only_baseline[:5]}",
              file=sys.stderr)
        return 1

    changed = [k for k in golden if golden[k] != baseline[k]]
    if changed:
        first = changed[0]
        print(f"baseline: FAIL - {len(changed)} entr(ies) in the golden file do "
              f"NOT match the pre-extraction code.\n\n  {first}\n"
              f"    golden:   {golden[first][:200]}\n"
              f"    pre-move: {baseline[first][:200]}\n\n"
              f"Either the extraction changed behavior, or the golden file was "
              f"regenerated from the new code.", file=sys.stderr)
        return 1

    print(f"baseline: OK - all {len(golden)} entries reproduce exactly from "
          f"rsvp_app.py at {BASELINE_COMMIT}, before the extraction")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

#!/usr/bin/env python3
"""Fail if personal data is tracked in git.

This repository is public and handles real colleagues' names, email addresses
and gift-contribution amounts. `rsvp_data.db` is written next to the app on
every run, so without a guard it gets committed again the moment someone runs
`git add -A`. That is how it got committed the first time.

Why a script and not a review instruction: "did this diff add PII?" is decided
by a pattern match over tracked files. A reviewer reading a diff will catch a
.db file once and miss it the third time; a script never gets bored. See
CLAUDE.md, "Deterministic work is not agent work".

Checks, over `git ls-files` (tracked files only - a gitignored working copy is
fine, that is the point of the ignore):
  1. No SQLite database is tracked, whatever it is named (magic-byte sniffed,
     so renaming rsvp_data.db to data.bin does not evade it).
  2. No corporate email address appears in a tracked text file.

Exit 0 clean, 1 violation found, 2 the check could not run.
"""
from __future__ import annotations

import codecs
import re
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SQLITE_MAGIC = b"SQLite format 3\x00"

# Any address at any real domain. The docs legitimately contain example.com
# and noreply@ addresses, so those are excluded - but by EXACT reserved domain,
# not by pattern. An earlier form listed a handful of TLDs (com, net, jp, ...)
# and excluded any domain merely BEGINNING with "example." or "test.", so an
# address at a .io or .fr company, or at a real domain whose first label
# happened to be "example", passed (Codex review of PR #1). The domain part now accepts any DNS suffix, and reserved_domain()
# below decides what is documentation.
#
# re.I matters: "Alice@Example.com" is as much a documentation address as
# "alice@example.com", and a case-sensitive exclusion flagged the first one as
# a real person. A guard that fires on RFC 2606 example domains trains people
# to ignore it.
EMAIL = re.compile(
    rb"[A-Za-z0-9._%+-]+@((?:[A-Za-z0-9-]+\.)+[A-Za-z]{2,})\b", re.I)

# RFC 2606 / RFC 6761 names that can never belong to a real person.
RESERVED_DOMAINS = {b"example.com", b"example.net", b"example.org"}
RESERVED_TLDS = {b"example", b"test", b"invalid", b"localhost"}


def reserved_domain(domain: bytes) -> bool:
    """True for a documentation-only domain, or a subdomain of one."""
    d = domain.lower().rstrip(b".")
    if d.rsplit(b".", 1)[-1] in RESERVED_TLDS:
        return True
    return any(d == r or d.endswith(b"." + r) for r in RESERVED_DOMAINS)


ALLOWED_ADDRESSES = {b"noreply@anthropic.com"}

# Bracketed at-forms only. A first attempt also matched a bare " at ", which
# matched ordinary English prose ("nothing at all. The...") in this repo's own
# documentation - three false positives on the first run. A guard that cries
# wolf gets switched off, so it now requires an unambiguous bracketed form.
# (The example spellings are assembled below rather than written out, because
# this file is itself scanned.)
_AT = rb"\[\s*at\s*\]|\(\s*at\s*\)"
_DOT = rb"\[\s*dot\s*\]|\(\s*dot\s*\)|\."
OBFUSCATED = re.compile(
    rb"[A-Za-z0-9._%+-]+\s*(?:" + _AT + rb")\s*"
    rb"[A-Za-z0-9.-]+\s*(?:" + _DOT + rb")\s*[A-Za-z]{2,}", re.I)

# A .csv or .tsv in THIS repository is a recipient roster until a human says
# otherwise. Names alone are personal data, and no regex can recognise a name -
# so rather than pretend to, the guard makes a person account for the file.
ROSTER_SUFFIXES = {".csv", ".tsv"}

# Binary documents can carry personal data as pixels - a screenshot of the app
# showing real recipients is invisible to every regex in this file. So these are
# not "skipped": each one must be named in .pii-allowlist by a human who opened
# it. An unaccounted binary FAILS. This is the difference between a guard that
# cannot see a leak and a guard that admits it cannot see and stops.
#
# Outlook's own formats are listed too: a saved .msg or .oft is an OLE document
# holding recipients and message text, and a .pst/.ost is a whole mailbox
# (Codex review of PR #1).
OPAQUE_SUFFIXES = {".png", ".jpg", ".jpeg", ".gif", ".ico", ".bmp", ".webp",
                   ".pdf", ".xlsx", ".xls", ".docx", ".doc", ".pptx", ".zip",
                   ".msg", ".oft", ".pst", ".ost"}
ALLOWLIST_FILE = ROOT / ".pii-allowlist"


# Windows tools readily save text as UTF-16, where every ASCII character is
# followed by a NUL byte and no address pattern can match the raw bytes: a
# UTF-16 roster.txt passed this check (Codex review of PR #1). BOM-marked
# UTF-16/32 is decoded before scanning. UTF-32 comes first because its
# little-endian BOM begins with UTF-16's.
_BOMS = ((codecs.BOM_UTF32_LE, "utf-32"), (codecs.BOM_UTF32_BE, "utf-32"),
         (codecs.BOM_UTF16_LE, "utf-16"), (codecs.BOM_UTF16_BE, "utf-16"))


def scannable_text(blob: bytes) -> bytes | None:
    """The bytes to run the address patterns over, or None when the file
    cannot be read as text and a human has to account for it instead.

    A NUL byte without a BOM means UTF-16 with no BOM, or a binary format this
    check has no name for; either way a regex over it proves nothing."""
    for bom, codec in _BOMS:
        if blob.startswith(bom):
            try:
                return blob.decode(codec).encode("utf-8")
            except UnicodeDecodeError:
                return None
    return None if b"\0" in blob else blob


def load_allowlist() -> dict[str, str | None]:
    """Path -> the git blob id a human approved, or None when the line
    carries no `blob=`."""
    if not ALLOWLIST_FILE.exists():
        return {}
    entries: dict[str, str | None] = {}
    for raw in ALLOWLIST_FILE.read_text(encoding="utf-8").splitlines():
        line = raw.split("#", 1)[0].strip()
        if not line:
            continue
        path, _, last = line.rpartition(" ")
        if last.startswith("blob=") and path.strip():
            entries[path.strip()] = last[len("blob="):]
        else:
            entries[line] = None
    return entries


def blob_id(path: Path) -> str:
    """The id git gives these contents. git's own clean filters apply, so a
    CRLF checkout on Windows hashes the same as the committed file."""
    return subprocess.run(
        ["git", "-C", str(ROOT), "hash-object", "--", str(path)],
        capture_output=True, check=True, text=True,
        encoding="utf-8").stdout.strip()


def approval_problem(rel: str, path: Path,
                     allowed: dict[str, str | None]) -> str | None:
    """None when a human approved exactly these bytes, else what to do.

    An approval names the CONTENTS, not only the path. By path alone, an
    approved screenshot regenerated in place with real recipients in it still
    passed (Codex review of PR #1)."""
    current = blob_id(path)
    if rel not in allowed:
        return (f"Open it, confirm it holds no real names/addresses/amounts or "
                f"corporate classification markings, then add "
                f"`{rel}  blob={current}  # <who checked, what they saw>` to "
                f".pii-allowlist - or untrack it.")
    approved = allowed[rel]
    if approved is None:
        return (f"It is in .pii-allowlist with no blob= digest, so the approval "
                f"is not tied to the bytes a human saw. Open it again and put "
                f"blob={current} on its line.")
    if approved != current:
        return (f"It CHANGED since it was approved (approved blob={approved}, "
                f"now blob={current}). Open it again, then update blob= on "
                f"its line.")
    return None


def tracked_files() -> list[Path]:
    """Tracked files PLUS untracked ones that are not gitignored.

    Tracked-only meant a brand-new file full of addresses passed this check
    right up until it was committed - the guard cleared the commit that
    introduced the leak. This is the set `git add -A` would stage, so the
    leak is caught while it is still uncommitted.
    """
    out = subprocess.run(
        ["git", "-C", str(ROOT), "ls-files", "-z", "--cached", "--others",
         "--exclude-standard"],
        capture_output=True, check=True,
    ).stdout
    return [ROOT / n.decode() for n in out.split(b"\0") if n]


def main() -> int:
    try:
        files = tracked_files()
    except (subprocess.CalledProcessError, FileNotFoundError) as exc:
        print(f"no-pii: cannot list tracked files: {exc}", file=sys.stderr)
        return 2

    allowed = load_allowlist()
    violations: list[str] = []
    scanned = 0
    opaque = 0

    for path in files:
        if not path.exists():
            continue
        try:
            head = path.open("rb").read(16)
        except OSError:
            continue

        if head.startswith(SQLITE_MAGIC):
            violations.append(
                f"{path.relative_to(ROOT)}: SQLite database is TRACKED IN GIT. "
                f"It holds recipients, responses and contribution amounts. "
                f"Untrack it (git rm --cached) and keep it gitignored."
            )
            continue

        if path.suffix.lower() in ROSTER_SUFFIXES:
            rel = path.relative_to(ROOT).as_posix()
            problem = approval_problem(rel, path, allowed)
            if problem:
                violations.append(
                    f"{rel}: tracked {path.suffix} file. Rosters hold names, and a "
                    f"name is personal data even with no address next to it - which "
                    f"no pattern here can detect. {problem}"
                )
                continue

        if path.suffix.lower() in OPAQUE_SUFFIXES:
            opaque += 1
            rel = path.relative_to(ROOT).as_posix()
            problem = approval_problem(rel, path, allowed)
            if problem:
                violations.append(
                    f"{rel}: tracked binary document that this check CANNOT read. "
                    f"Personal data in a screenshot is invisible to a text scan. "
                    f"{problem}"
                )
            continue

        try:
            blob = scannable_text(path.read_bytes())
        except OSError:
            continue
        if blob is None:
            opaque += 1
            rel = path.relative_to(ROOT).as_posix()
            problem = approval_problem(rel, path, allowed)
            if problem:
                violations.append(
                    f"{rel}: holds NUL bytes, so it is binary or UTF-16 without "
                    f"a byte-order mark, and this check CANNOT read it (re-saving "
                    f"it as UTF-8 also works). {problem}"
                )
            continue
        scanned += 1
        for m in EMAIL.finditer(blob):
            addr = m.group(0)
            if addr.lower() in ALLOWED_ADDRESSES or reserved_domain(m.group(1)):
                continue
            line = blob[: m.start()].count(b"\n") + 1
            violations.append(
                f"{path.relative_to(ROOT)}:{line}: real email address "
                f"{addr.decode(errors='replace')!r} in a tracked file"
            )
        for m in OBFUSCATED.finditer(blob):
            line = blob[: m.start()].count(b"\n") + 1
            violations.append(
                f"{path.relative_to(ROOT)}:{line}: obfuscated email address "
                f"{m.group(0).decode(errors='replace').strip()!r} in a tracked file"
            )

    if not scanned:
        print("no-pii: FAIL - scanned no text files; the guard has drifted",
              file=sys.stderr)
        return 1

    for v in violations:
        print(f"no-pii: {v}", file=sys.stderr)
    if violations:
        print(f"\nno-pii: FAIL - {len(violations)} violation(s)", file=sys.stderr)
        return 1
    print(f"no-pii: OK - {scanned} text file(s) scanned, "
          f"{opaque} binary file(s) human-allowlisted, no PII found")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

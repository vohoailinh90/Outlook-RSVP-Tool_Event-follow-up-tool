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

Both the working copy and the staged copy of each file are checked, because
either one can end up in the next commit.

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

# Files larger than this are never loaded whole, nor is anything opaque by
# name: a mailbox-sized .pst could otherwise exhaust memory before the guard
# said it needs approval (Codex review of PR #5). They are judged by blob id,
# like any opaque file, after a look at their first bytes.
SCAN_LIMIT = 16 * 1024 * 1024


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


def index_blobs() -> tuple[dict[str, str], set[str]]:
    """Path -> blob id of the copy STAGED in the index, which is what the next
    commit will contain; and the set of submodule paths, which hold a commit
    id rather than file contents."""
    out = subprocess.run(
        ["git", "-C", str(ROOT), "ls-files", "-s", "-z"],
        capture_output=True, check=True,
    ).stdout
    blobs: dict[str, str] = {}
    gitlinks: set[str] = set()
    for entry in out.split(b"\0"):
        meta, _, name = entry.partition(b"\t")
        if not name:
            continue
        mode, obj = meta.split()[:2]
        if mode == b"160000":
            gitlinks.add(name.decode())
        else:
            blobs[name.decode()] = obj.decode()
    return blobs, gitlinks


def staged_sizes(ids: set[str]) -> dict[str, int]:
    """Blob id -> size in bytes, without reading any contents. An id git
    cannot read is left out."""
    out = subprocess.run(
        ["git", "-C", str(ROOT), "--no-replace-objects", "cat-file",
         "--batch-check"],
        input="".join(f"{i}\n" for i in sorted(ids)).encode("ascii"),
        capture_output=True, check=True,
    ).stdout
    sizes: dict[str, int] = {}
    for line in out.splitlines():
        header = line.split()
        if len(header) == 3:                 # <id> <type> <size>
            sizes[header[0].decode()] = int(header[2])
    return sizes


def staged_head(blob: str, n: int) -> bytes | None:
    """The first n bytes of a staged blob, or None when git cannot read it.
    Only those bytes are read: the process is stopped once they arrive, so
    a mailbox-sized blob is never loaded to learn it is a database."""
    proc = subprocess.Popen(
        ["git", "-C", str(ROOT), "--no-replace-objects", "cat-file", "blob",
         blob], stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
    head = b""
    try:
        head = proc.stdout.read(n)
    finally:
        proc.kill()             # no-op if git already exited
        proc.stdout.close()
        code = proc.wait()
    # A blob shorter than n ends before the kill, so its exit code is real.
    if len(head) < n and code != 0:
        return None
    return head


class StagedBlobs:
    """Reads staged blobs one at a time, on demand, through one long-lived
    `git cat-file --batch` process. At most one blob is held in memory: all
    blobs read up front in a single call needed their total size, which
    for many blobs just under SCAN_LIMIT ran to gigabytes (Codex review of
    PR #5).

    Staged and working copies are compared by these raw bytes, never by blob
    id. Each earlier shortcut let a staged address through (Codex review of
    PR #5):
      - `git diff-files` trusts assume-unchanged and skip-worktree, and omits
        a path with either flag set.
      - `git hash-object` runs clean filters, so equal ids only prove that
        re-adding the working file reproduces the staged blob; a filter that
        appends an address hashed a harmless file to the blob holding it.
    The ids go in on stdin, so no path ever passes through a line-based
    stream.

    --no-replace-objects: `cat-file` otherwise honours a local refs/replace
    ref and returns its bytes under the ORIGINAL id, so a harmless
    replacement hid a staged address that a clone would receive (Codex
    review of PR #5)."""

    def __init__(self) -> None:
        self.proc: subprocess.Popen | None = None

    def read(self, blob: str) -> bytes | None:
        """The exact bytes the next commit will carry for this blob, or None
        when git cannot read it."""
        try:
            if self.proc is None:
                self.proc = subprocess.Popen(
                    ["git", "-C", str(ROOT), "--no-replace-objects",
                     "cat-file", "--batch"],
                    stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                    stderr=subprocess.DEVNULL)
            self.proc.stdin.write(f"{blob}\n".encode("ascii"))
            self.proc.stdin.flush()
            header = self.proc.stdout.readline().split()
            if len(header) != 3:             # "<id> missing", or git died
                return None
            size = int(header[2])            # <id> <type> <size>
            data = self.proc.stdout.read(size)
            self.proc.stdout.read(1)         # the newline after the contents
            return data if len(data) == size else None
        except (OSError, ValueError):
            return None

    def close(self) -> None:
        if self.proc is not None:
            self.proc.stdin.close()
            self.proc.stdout.close()
            self.proc.wait()


def approval_problem(rel: str, current: str | None,
                     allowed: dict[str, str | None],
                     staged: str | None) -> str | None:
    """None when a human approved exactly these bytes, else what to do.

    `current` is the working copy's blob id, or None when it has no working
    copy and only the staged one will be committed.

    An approval names the CONTENTS, not only the path. By path alone, an
    approved screenshot regenerated in place with real recipients in it still
    passed (Codex review of PR #1).

    Both copies must match: the working tree AND the index. Checking only the
    working tree passed a changed image that was staged and then had its
    working copy restored from HEAD - the commit would carry the unreviewed
    blob (Codex review of PR #4)."""
    ids = {i for i in (current, staged) if i is not None}
    if current is None:
        copies = f"staged blob={staged}"
    else:
        copies = f"working copy blob={current}"
        if staged is not None and staged != current:
            copies += f", staged blob={staged}"
    if rel not in allowed:
        if len(ids) > 1:
            return (f"Its working copy and staged copy differ ({copies}). "
                    f"Stage the version you mean to commit, then review it.")
        return (f"Open it, confirm it holds no real names/addresses/amounts or "
                f"corporate classification markings, then add "
                f"`{rel}  blob={next(iter(ids))}  # <who checked, what they saw>` "
                f"to .pii-allowlist - or untrack it.")
    approved = allowed[rel]
    if approved is None:
        return (f"It is in .pii-allowlist with no blob= digest, so the approval "
                f"is not tied to the bytes a human saw. Open it again and put "
                f"the blob= of the reviewed copy on its line ({copies}).")
    if ids != {approved}:
        return (f"It CHANGED since it was approved (approved blob={approved}, "
                f"{copies}). Open it again, then update blob= on its line.")
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
    # A path in a merge conflict is listed once per index stage; check it once.
    return [ROOT / n.decode() for n in dict.fromkeys(out.split(b"\0")) if n]


def main() -> int:
    try:
        files = tracked_files()
        staged, gitlinks = index_blobs()
        sizes = staged_sizes(set(staged.values()))
    except (subprocess.CalledProcessError, FileNotFoundError) as exc:
        detail = getattr(exc, "stderr", None) or b""
        print(f"no-pii: cannot read the repository: {exc} "
              f"{detail.decode(errors='replace').strip()}", file=sys.stderr)
        return 2

    allowed = load_allowlist()
    violations: list[str] = []
    scanned = 0
    opaque = 0

    reader = StagedBlobs()
    for path in files:
        rel = path.relative_to(ROOT).as_posix()
        # A submodule holds a commit id, not contents, and once initialised it
        # is a directory that cannot be read as a file. Its own repository is
        # checked there (Codex review of PR #5). A regular file that has
        # replaced one is still scanned: `git add -A` stages it as a file.
        if rel in gitlinks and not path.is_file():
            continue
        # Every copy the next commit could carry: the working copy, which
        # `git add -A` would stage, and the staged copy when it differs. Reading
        # only the working copy passed an address staged behind a clean file,
        # and skipped a staged file whose working copy was deleted (Codex
        # review of PR #1).
        #
        # A copy that cannot be read is a violation, not a skip: this guard
        # fails closed. Sharing one try/except let a failed staged read
        # discard a working copy already read, address and all (review of
        # e42f925).
        copies: list[tuple[str, bytes]] = []
        unreadable: list[str] = []
        working_read = False
        staged_size = sizes.get(staged[rel]) if rel in staged else None
        if rel in staged and staged_size is None:
            unreadable.append(f"staged copy (blob {staged[rel]})")
        try:
            working_size = path.stat().st_size if path.is_file() else None
        except OSError:
            working_size = None
        too_big = max(working_size or 0, staged_size or 0) > SCAN_LIMIT
        if too_big or path.suffix.lower() in OPAQUE_SUFFIXES:
            head = b""
            if path.exists():
                try:
                    with path.open("rb") as f:
                        head = f.read(len(SQLITE_MAGIC))
                    working_read = True
                except OSError as exc:
                    unreadable.append(f"working copy ({exc})")
            if unreadable:
                violations.append(
                    f"{rel}: cannot read its {' or '.join(unreadable)}, so it "
                    f"cannot be checked. Close whatever holds it and run again.")
            # The staged copy's first bytes are sniffed as well: a database
            # staged under an opaque name behind a clean working copy was
            # otherwise reported only as an unapproved binary, never as the
            # database it is (review of ca9b106).
            staged_ok = rel in staged and staged_size is not None
            staged_first = (staged_head(staged[rel], len(SQLITE_MAGIC))
                            if staged_ok else None)
            if staged_ok and staged_first is None:
                staged_ok = False
                violations.append(
                    f"{rel}: cannot read its staged copy (blob {staged[rel]}), "
                    f"so it cannot be checked.")
            if not working_read and not staged_ok:
                continue
            opaque += 1
            db_copy = next((label for label, data in
                            (("", head), (" (staged copy)", staged_first))
                            if data and data.startswith(SQLITE_MAGIC)), None)
            if db_copy is not None:
                violations.append(
                    f"{rel}{db_copy}: SQLite database is TRACKED IN GIT. It "
                    f"holds recipients, responses and contribution amounts. "
                    f"Untrack it (git rm --cached) and keep it gitignored.")
                continue
            suffix = path.suffix.lower()
            if suffix in OPAQUE_SUFFIXES:
                what = ("tracked binary document that this check CANNOT read. "
                        "Personal data in a screenshot is invisible to a text "
                        "scan.")
            elif suffix in ROSTER_SUFFIXES:
                what = (f"tracked {path.suffix} file. Rosters hold names, and a "
                        f"name is personal data even with no address next to "
                        f"it - which no pattern here can detect.")
            else:
                what = (f"is over the {SCAN_LIMIT // 2**20} MiB scan limit, so "
                        f"this check does not read it.")
            current = None
            if working_read:
                try:
                    current = blob_id(path)
                except subprocess.CalledProcessError as exc:
                    violations.append(
                        f"{rel}: cannot hash its working copy ({exc}), so it "
                        f"cannot be checked. Close whatever holds it and run "
                        f"again.")
            # A staged copy git cannot read is already a violation; asking a
            # human to approve its id would ask for the impossible (review of
            # ca9b106).
            staged_id = staged[rel] if staged_ok else None
            if current is not None or staged_id is not None:
                problem = approval_problem(rel, current, allowed, staged_id)
                if problem:
                    violations.append(f"{rel}: {what} {problem}")
            continue

        if path.exists():
            try:
                copies.append(("", path.read_bytes()))
                working_read = True
            except OSError as exc:
                unreadable.append(f"working copy ({exc})")
        if rel in staged and staged_size is not None:
            # Only reached for a blob under the scan limit and not opaque
            # by name; read now, when this file is scanned.
            staged_copy = reader.read(staged[rel])
            if staged_copy is None:
                unreadable.append(f"staged copy (blob {staged[rel]})")
            elif not copies or staged_copy != copies[0][1]:
                copies.append((" (staged copy)", staged_copy))
        if unreadable:
            violations.append(
                f"{rel}: cannot read its {' or '.join(unreadable)}, so it "
                f"cannot be checked. Close whatever holds it and run again.")
        if not copies:
            continue

        def approval(what: str) -> None:
            # A working copy that could not be read cannot be hashed either:
            # on Windows, git cannot open a file another program holds. The
            # uncaught error stopped the guard before the other files were
            # scanned (Codex review of PR #5). The unreadable copy is already
            # a violation; approval is then judged on the staged copy alone.
            current = None
            if working_read:
                try:
                    current = blob_id(path)
                except subprocess.CalledProcessError as exc:
                    violations.append(
                        f"{rel}: cannot hash its working copy ({exc}), so it "
                        f"cannot be checked. Close whatever holds it and run "
                        f"again.")
            if current is None and staged.get(rel) is None:
                return
            problem = approval_problem(rel, current, allowed, staged.get(rel))
            if problem:
                violations.append(f"{rel}: {what} {problem}")

        db_copy = next((label for label, data in copies
                        if data.startswith(SQLITE_MAGIC)), None)
        if db_copy is not None:
            violations.append(
                f"{rel}{db_copy}: SQLite database is TRACKED IN GIT. "
                f"It holds recipients, responses and contribution amounts. "
                f"Untrack it (git rm --cached) and keep it gitignored."
            )
            continue

        if path.suffix.lower() in ROSTER_SUFFIXES:
            before = len(violations)
            approval(
                f"tracked {path.suffix} file. Rosters hold names, and a name is "
                f"personal data even with no address next to it - which no "
                f"pattern here can detect.")
            if len(violations) > before:
                continue

        texts = [(label, scannable_text(data)) for label, data in copies]
        if any(blob is None for _, blob in texts):
            opaque += 1
            approval(
                "holds NUL bytes, so it is binary or UTF-16 without a byte-order "
                "mark, and this check CANNOT read it (re-saving it as UTF-8 also "
                "works).")
            continue
        scanned += 1
        seen: set[tuple[str, bytes]] = set()
        for label, blob in texts:
            for kind, pattern in (("real", EMAIL), ("obfuscated", OBFUSCATED)):
                for m in pattern.finditer(blob):
                    if kind == "real" and (
                            m.group(0).lower() in ALLOWED_ADDRESSES
                            or reserved_domain(m.group(1))):
                        continue
                    found = m.group(0).strip()
                    if (kind, found) in seen:
                        continue
                    seen.add((kind, found))
                    line = blob[: m.start()].count(b"\n") + 1
                    violations.append(
                        f"{rel}{label}:{line}: {kind} email address "
                        f"{found.decode(errors='replace')!r} in a tracked file"
                    )

    reader.close()

    # Printed first: returning on the drift check swallowed every violation
    # already found (review of ca9b106).
    for v in violations:
        print(f"no-pii: {v}", file=sys.stderr)
    if not scanned:
        print("no-pii: FAIL - scanned no text files; the guard has drifted",
              file=sys.stderr)
        return 1

    if violations:
        print(f"\nno-pii: FAIL - {len(violations)} violation(s)", file=sys.stderr)
        return 1
    print(f"no-pii: OK - {scanned} text file(s) scanned, "
          f"{opaque} binary file(s) human-allowlisted, no PII found")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

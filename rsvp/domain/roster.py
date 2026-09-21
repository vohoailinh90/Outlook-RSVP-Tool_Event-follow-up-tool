"""Turning a recipient list into the list of real people to track.

A row on the Recipients tab may be an Exchange distribution list rather than a
person. For vote tracking it has to be replaced by its members, or the
"responded / not responded" panel shows one ambiguous group row instead of
each colleague.

Expanding a distribution list requires Outlook, which is why this module takes
the expander as an argument instead of importing it. The merge rules - order,
de-duplication, what happens when expansion fails - are ordinary logic and are
testable here without Outlook.
"""
from __future__ import annotations


def merge_expanded_roster(recipients, expand, cache=None):
    """Expand group rows and de-duplicate, preserving first-seen order.

    `expand(email)` returns a list of (name, email) members, or None when the
    address is not a group - or when Outlook could not resolve it. Both cases
    are treated the same on purpose: an address that cannot be expanded is
    kept as the person it appears to be, rather than dropped.

    `cache` is an optional dict reused across calls so the same group is not
    queried against the address book twice in one session. It is mutated.

    Comparison is case-insensitive on the address, because Exchange returns
    casing inconsistently and the same colleague must not appear twice.
    """
    if cache is None:
        cache = {}

    roster: list[tuple[str, str]] = []
    seen: set[str] = set()

    for name, email in recipients:
        key = (email or "").lower()
        if key not in cache:
            try:
                cache[key] = expand(email)
            except Exception:
                # A COM failure means "not expandable", not "drop this person".
                cache[key] = None

        members = cache[key]
        if members is None:
            if key not in seen:
                seen.add(key)
                roster.append((name, email))
        else:
            for member_name, member_email in members:
                member_key = (member_email or "").lower()
                if member_key not in seen:
                    seen.add(member_key)
                    roster.append((member_name, member_email))

    return roster

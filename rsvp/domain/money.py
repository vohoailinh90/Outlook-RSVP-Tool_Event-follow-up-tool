"""Amounts: parsing them out of free text, and the arithmetic over them.

These numbers decide what colleagues are asked to contribute and what the
organiser is shown as still owing, so the parsing rules below are written out
explicitly rather than left to a regex that happens to work on the examples
someone tried.

A note on the type: these are floats, as they always were. JPY and VND are
effectively integer currencies and the UI renders with `,.0f`, so float is
adequate here, but it is not the right type for money in general. Converting
to Decimal would change every call site and the database column type, which is
a larger change than this one and does not belong in the same commit.
"""
from __future__ import annotations

import re

# A number, possibly with grouping separators: 3000, 3,000, 3.000, 1,234,567.89
_NUMBER = r"\d[\d.,]*\d|\d"

# Currency markers seen in this tool's event budgets: Japanese, Vietnamese,
# and the bare ISO codes people type. Used to pick the RIGHT number out of a
# sentence, not to validate or convert anything.
_CURRENCY = r"(?:JPY|VND|USD|EUR|yen|円|¥|₫|đ|\$|dong|đồng)"

# A number with a currency marker on either side, e.g. "3000 JPY", "¥3,000".
_WITH_CURRENCY = re.compile(
    rf"(?:{_CURRENCY}\s*(?P<before>{_NUMBER}))|(?:(?P<after>{_NUMBER})\s*{_CURRENCY})",
    re.IGNORECASE,
)
_ANY_NUMBER = re.compile(_NUMBER)


def _to_float(token: str) -> float:
    """Resolve grouping separators, then convert.

    The hard case is a single separator: "3.000" is three thousand in
    Vietnamese and European formatting, but Python reads it as 3.0 - a
    thousand-fold understatement of a contribution, silently. The rule used
    here is the one those formats actually follow: a separator followed by
    exactly three digits, with no other separator present, is a thousands
    separator. "12.5" keeps its decimal point because 5 is not three digits.
    """
    token = token.strip()
    has_comma, has_dot = "," in token, "." in token

    if has_comma and has_dot:
        # Both present: whichever comes last is the decimal separator.
        if token.rfind(",") > token.rfind("."):
            token = token.replace(".", "").replace(",", ".")
        else:
            token = token.replace(",", "")
    elif has_comma or has_dot:
        sep = "," if has_comma else "."
        head, _, tail = token.rpartition(sep)
        if len(tail) == 3 and token.count(sep) >= 1 and head:
            # Thousands grouping: 3,000 / 3.000 / 1.234.567
            token = token.replace(sep, "")
        else:
            # Decimal: 12.5 / 12,5
            token = token.replace(sep, ".")

    try:
        return float(token)
    except ValueError:
        return 0.0


def parse_amount_from_text(text) -> float:
    """Pull the intended amount out of a free-text budget field.

    Tab 1's "Expected gift budget" is typed by hand, so it arrives as things
    like "3,000 JPY / person" or "2026 year-end party, 3000 JPY".

    A number sitting next to a currency marker wins over one that is not.
    The previous implementation took the FIRST number in the string, which
    read the year out of "2026 year-end party, 3000 JPY" and the headcount
    out of "5 people x 3000 JPY" - both silently, and both into money totals.
    With no currency marker anywhere, the first number is still used, so a
    plain "3000" behaves exactly as before.

    Returns 0.0 when there is no number at all.
    """
    if not text:
        return 0.0
    text = str(text)

    match = _WITH_CURRENCY.search(text)
    if match:
        return _to_float(match.group("before") or match.group("after"))

    match = _ANY_NUMBER.search(text)
    return _to_float(match.group(0)) if match else 0.0


def format_amount(value: float) -> str:
    """Render an amount the way every total in the UI is rendered."""
    return f"{value:,.0f}"


def remaining_amount(total_collected: float, amount_paid: float) -> float:
    """What the organiser is still owed (negative means overpaid)."""
    return total_collected - amount_paid


def sum_contributions(roster) -> float:
    """Total of the `amount` field across a roster of per-person dicts."""
    return sum(info.get("amount", 0.0) or 0.0 for info in roster)


def count_actual_attendees(roster) -> int:
    """How many people are marked as having actually attended."""
    return sum(1 for info in roster
               if (info.get("actual_attend") or "").strip().lower() == "yes")

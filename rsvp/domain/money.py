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
# and the bare ISO codes people type.
#
# The alphabetic codes must not touch another LETTER on either side. Without
# that, "EUR" matched inside "European", "dong" inside "dongles" and "usd"
# inside "usdollars", so "Budget 3000 per person for the 2026 European event"
# returned 2026 - the exact first-number bug this function exists to prevent,
# wearing a disguise.
#
# The guard is a letter-excluding lookaround, not \b: digits are word
# characters too, so \b missed the glued forms "12.500USD" and "USD12.500"
# and those fell back to the first number in the string (Codex review of
# PR #2). A digit may touch a code; a letter may not.
#
# "đ" cannot take a leading \b, because "200000đ" has no boundary between the
# digit and the letter. It instead requires that no letter FOLLOWS, so it
# matches the suffix in "200000đ" but not the first letter of "đại hội".
#
# "US$" is listed before the bare symbols so that it is found as one marker:
# otherwise only its "$" would match, and the USD decimal rule below could not
# tell it apart from a "$" that means some other dollar or peso. It takes a
# "no letter before" lookbehind rather than \b, so the glued suffix in
# "3000US$" is found while "BUS$" is not.
_CURRENCY = (
    r"(?:(?<![^\W\d_])US\$"
    r"|(?<![^\W\d_])(?:JPY|VND|VN\u0110|USD|EUR|yen|dong|\u0111\u1ed3ng)"
    r"(?![^\W\d_])"
    r"|[\u00a5\u20ab$\u5186]"
    r"|\u0111(?![^\W\d_]))"
)
_CURRENCY_RE = re.compile(_CURRENCY, re.IGNORECASE)

# Anchored number patterns, used to look just beside a marker rather than to
# scan the whole string for a number-then-marker pair.
_NUMBER_AT_START = re.compile(rf"\s*({_NUMBER})")
_NUMBER_AT_END = re.compile(rf"({_NUMBER})\s*$")
_ANY_NUMBER = re.compile(_NUMBER)


# Markers of a currency written with a minor unit and an English-style decimal
# point: "12.500 USD" / "US$3.500" mean twelve and a half / three and a half
# dollars, not thousands.
#
# Only unambiguous USD markers belong here. A bare "$" is NOT one: it is
# shared by currencies that group with the dot (a Chilean "$3.000" is three
# thousand pesos), and a Vietnamese organiser may well type "$3.000" for three
# thousand dollars. Raised by Codex review of PR #2; the owner chose to keep
# the default thousands reading for bare "$". EUR is not here either:
# European formatting uses the dot for grouping ("3.000 EUR" is three
# thousand), so the default rule is already right for it.
#
# The decision looks at BOTH sides of the chosen number, not only at the
# marker that selected it: in "$3.000 USD" the bare "$" is found first, but
# the explicit "USD" beside the same number settles the currency (Codex
# review of PR #2). Same adjacency as everywhere else: whitespace at most,
# except that a bare "$" may sit between a leading "USD" and the number, as in
# "USD $3.000".
_USD_RIGHT_AFTER = re.compile(r"\s*(?:US\$|USD(?![^\W\d_]))",
                              re.IGNORECASE)
_USD_RIGHT_BEFORE = re.compile(
    r"(?<![^\W\d_])(?:US\$|USD(?:\s*\$)?)\s*$", re.IGNORECASE)


def _usd_marker_beside(text: str, start: int, end: int) -> bool:
    """True when an explicit USD marker sits beside text[start:end]."""
    return bool(_USD_RIGHT_AFTER.match(text, end)
                or _USD_RIGHT_BEFORE.search(text, 0, start))


def _amount_beside_a_currency_marker(text: str) -> re.Match | None:
    """Return the match of the number adjacent to the first currency marker,
    if any. Its group(1) is the number; its span locates it in `text`.

    Markers are located first, then the text immediately beside each one is
    checked for a number. The earlier form was a single alternation that, for
    every starting digit, matched a whole digit run and then backtracked
    looking for a marker - quadratic on a long marker-free number, and this
    function runs synchronously from a keystroke handler, so pasting a few
    thousand digits into the amount field stalled the UI for most of a second.
    Locating markers first makes each digit run be considered once.

    The number BEFORE a marker is found by stepping back from the marker, not
    by searching from the start of the text: a search from offset zero for
    every marker was quadratic again on a paste of many markers with no
    number beside them - 16,000 "$" took most of a second (Codex review of
    PR #1).
    """
    for marker in _CURRENCY_RE.finditer(text):
        after = _NUMBER_AT_START.match(text, marker.end())
        before = _number_just_before(text, marker.start())
        if after and before:
            return _nearer_side(marker, before, after)
        if after or before:
            return after or before
    return None


# Symbols conventionally written BEFORE the amount: "$50", "US$50", "¥3000".
# Every other marker - the ISO codes, 円, đ, yen, dong - is conventionally
# written after it: "3,000 JPY", "3000円", "200.000đ".
_PREFIX_MARKERS = {"$", "US$", "\u00a5"}


def _nearer_side(marker: re.Match, before: re.Match, after: re.Match) -> re.Match:
    """Choose between numbers on both sides of one marker.

    Always taking the number after it read a headcount as the amount:
    "3,000 JPY 5 people" -> 5, "3000円 5人" -> 5, "200.000đ 3 người" -> 3
    (Codex review of PR #1). A number written against the marker with no
    space belongs to it; failing that, the marker's conventional side wins.
    """
    glued_before = before.end(1) == marker.start()
    glued_after = after.start(1) == marker.end()
    if glued_before != glued_after:
        return before if glued_before else after
    return after if marker.group(0).upper() in _PREFIX_MARKERS else before


def _number_just_before(text: str, end: int) -> re.Match | None:
    """The match of _NUMBER_AT_END ending at `end`, found in time
    proportional to the number and the whitespace before `end`."""
    stop = end
    while stop > 0 and text[stop - 1].isspace():
        stop -= 1
    start = stop
    while start > 0 and (text[start - 1].isdecimal() or text[start - 1] in ".,"):
        start -= 1
    # _NUMBER starts with a digit, so skip separators left of the first one.
    while start < stop and not text[start].isdecimal():
        start += 1
    if start == stop:
        return None
    return _NUMBER_AT_END.match(text, start, end)


def _to_float(token: str, dot_is_decimal: bool = False) -> float:
    """Resolve grouping separators, then convert.

    The hard case is a single separator: "3.000" is three thousand in
    Vietnamese and European formatting, but Python reads it as 3.0 - a
    thousand-fold understatement of a contribution, silently. The rule used
    here is the one those formats actually follow: a separator followed by
    exactly three digits, with no other separator present, is a thousands
    separator. "12.5" keeps its decimal point because 5 is not three digits.

    `dot_is_decimal` is set when the amount is written beside an explicit USD
    marker ("USD" or "US$", not a bare "$").
    A single dot is then always a decimal point, so "12.500 USD" is 12.5.
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
        # Three trailing digits after a single separator is read as
        # thousands grouping by default. That is correct for JPY and VND -
        # the two currencies this tool is actually used for, both of which
        # have no minor unit in practice - and for a bare "3.000", which a
        # Vietnamese organiser types for three thousand.
        #
        # The one exception is a single dot beside an explicit USD marker:
        # "US$3.500" and "12.500 USD" are dollars with a minor unit, so the
        # dot is a decimal point. A bare "$" does not qualify. The repository
        # owner chose this on PR #1 over both the locale-blind rule and
        # rejecting the form. A comma stays grouping
        # for USD ("3,000 USD"), and a repeated dot ("1.234.567 USD") can
        # only be grouping, so neither is affected. Pinned by
        # TestKnownAmbiguousCases in tests/test_domain_money.py.
        #
        # A leading zero before the separator means a decimal, never grouping:
        # nobody writes "0.500" for five hundred, but "0.500" for a half is
        # ordinary. Without this, the three-digit rule turned 0.5 into 500 -
        # wrong by 1000x UPWARD, which is a worse failure than the
        # thousand-fold understatement this rule exists to fix.
        looks_grouped = len(tail) == 3 and head and head.lstrip("+-") != "0"
        if dot_is_decimal and sep == "." and token.count(".") == 1:
            looks_grouped = False
        if looks_grouped:
            # Thousands grouping: 3,000 / 3.000 / 1.234.567
            token = token.replace(sep, "")
        else:
            # Decimal: 12.5 / 12,5 / 0.500
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

    "Next to" means adjacent, separated by whitespace at most: "3000 JPY",
    "JPY 3000", "¥3,000". A marker further away does not count, and the first
    number is used instead - so "JPY quota is 10 max, paid 3000" yields 10.0,
    not 3000.0. That limit is deliberate: allowing words in between would let
    a marker in one clause capture a number from another, which is the same
    class of error in the opposite direction.

    With no adjacent currency marker, the first number is used, so a plain
    "3000" behaves exactly as before.

    Returns 0.0 when there is no number at all.
    """
    if not text:
        return 0.0
    text = str(text)

    beside = _amount_beside_a_currency_marker(text)
    if beside is not None:
        return _to_float(
            beside.group(1),
            dot_is_decimal=_usd_marker_beside(text, *beside.span(1)))

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

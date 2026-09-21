"""Money rules: parsing amounts out of free text, and the arithmetic on them.

These numbers decide what colleagues are asked to contribute and what the
organiser is shown as outstanding, so the cases below are written as the rules
themselves rather than as a handful of examples that happen to pass.

Phase 2 changed parse_amount_from_text. TestBehaviourChangedInPhase2 lists
every input whose result moved, with the old value recorded, so the change is
reviewable rather than buried in a diff.
"""
from __future__ import annotations

import pytest

from rsvp.domain import (
    count_actual_attendees,
    format_amount,
    parse_amount_from_text,
    remaining_amount,
    sum_contributions,
)


class TestUnchangedByPhase2:
    """Inputs that behaved correctly before and must still behave that way."""

    @pytest.mark.parametrize("text,expected", [
        ("3,000 JPY / person", 3000.0),
        ("3000", 3000.0),
        ("", 0.0),
        ("   ", 0.0),
        ("no digits here", 0.0),
        (None, 0.0),
        ("¥3,000", 3000.0),
        ("1,234,567 JPY", 1234567.0),
        ("3,000.50", 3000.5),
        ("Budget 5000 VND", 5000.0),
        ("khoảng 200000đ/người", 200000.0),
        ("予算 3,000円", 3000.0),
        ("12.5", 12.5),
        ("abc 12.5 xyz", 12.5),
        ("0", 0.0),
        # No currency marker anywhere: the first number is still used, so a
        # bare year in a field with no amount reads as it always did.
        ("Event 2026", 2026.0),
        ("2026 party", 2026.0),
    ])
    def test_result_is_unchanged(self, text, expected):
        assert parse_amount_from_text(text) == expected


class TestBehaviourChangedInPhase2:
    """Defects fixed. Each case records what it used to return.

    Every one of these silently produced a wrong money figure: the UI showed
    a confident number and nothing indicated it came from the wrong part of
    the string.
    """

    @pytest.mark.parametrize("text,was,now,why", [
        ("2026 year-end party, 3000 JPY", 2026.0, 3000.0,
         "took the first number, which was the year"),
        ("5 people x 3000 JPY", 5.0, 3000.0,
         "took the first number, which was the headcount"),
        ("3.000", 3.0, 3000.0,
         "read a Vietnamese/European thousands separator as a decimal point - "
         "a thousand-fold understatement of a contribution"),
        ("3.000 VND", 3.0, 3000.0, "same separator bug, with a currency marker"),
        ("1.234.567", 0.0, 1234567.0,
         "float('1.234.567') raised, and the exception handler returned 0.0 - "
         "a real amount silently became nothing"),
        ("approx 3000-4000 JPY", 3000.0, 4000.0,
         "an ambiguous range: it used to take the first number, it now takes "
         "the one the currency marker is attached to. Neither reading is "
         "obviously right; this one at least follows the stated rule"),
    ])
    def test_defect_is_fixed(self, text, was, now, why):
        assert parse_amount_from_text(text) == now, why
        assert was != now, "this case belongs in TestUnchangedByPhase2"


class TestStillWrongOnPurpose:
    """Known remaining limits, characterized so they are not mistaken for fixed."""

    def test_a_leading_zero_is_a_decimal_not_grouping(self):
        """Found by an exhaustive old-vs-new diff over 1,998 inputs: the
        three-digit rule turned 0.5 into 500, wrong by 1000x UPWARD, which is
        a worse failure than the understatement the rule exists to fix.
        Nobody writes "0.500" to mean five hundred."""
        assert parse_amount_from_text("0.500") == 0.5
        assert parse_amount_from_text("0,500") == 0.5

    def test_negative_amounts_lose_their_sign(self):
        """Out of scope for phase 2: allowing negatives changes what the
        totals and the `,.0f` displays can show, which is a wider change than
        correcting which number gets picked."""
        assert parse_amount_from_text("-500") == 500.0

    def test_amounts_are_floats_not_decimals(self):
        """Adequate for JPY and VND, which are integer currencies rendered
        with `,.0f`. Not the right type for money in general; converting
        would change every call site and the database column."""
        assert isinstance(parse_amount_from_text("3000"), float)


class TestKnownAmbiguousCases:
    """Inputs where no reading is objectively right, pinned so the choice is
    reviewable rather than accidental.

    Raised by review of 727a0fd: the thousands rule is locale-blind, so a USD
    amount written with a trailing zero reads as thousands. These assertions
    exist to make that a decision on the record - if someone later decides the
    other reading is correct, they change a test that says why, instead of
    discovering the behaviour from a wrong invoice.
    """

    @pytest.mark.parametrize("text,parsed,note", [
        ("$3.500 per head", 3500.0,
         "right for JPY/VND, wrong if the author meant three dollars fifty"),
        ("3.500 USD", 3500.0, "same ambiguity with an explicit ISO code"),
        ("3.500 JPY", 3500.0, "unambiguous: JPY has no minor unit in practice"),
        ("3.500 VND", 3500.0, "unambiguous: VND likewise"),
        ("1.234", 1234.0, "no currency at all; thousands is the likelier intent"),
    ])
    def test_three_trailing_digits_reads_as_thousands(self, text, parsed, note):
        assert parse_amount_from_text(text) == parsed, note

    def test_two_trailing_digits_stays_a_decimal(self):
        """The standard way to write a minor unit is unaffected, which is why
        the ambiguity above is narrow rather than general."""
        assert parse_amount_from_text("$3.50") == 3.5
        assert parse_amount_from_text("3.50 USD") == 3.5

    def test_a_distant_currency_marker_does_not_win(self):
        """Adjacency is required. Widening it would let a marker in one clause
        capture a number from another - the same error in reverse."""
        assert parse_amount_from_text("JPY quota is 10 max, paid 3000") == 10.0
        assert parse_amount_from_text("cost is 3000 (JPY)") == 3000.0


class TestArithmetic:
    def test_remaining_is_collected_minus_paid(self):
        assert remaining_amount(36000.0, 30000.0) == 6000.0

    def test_remaining_goes_negative_when_overpaid(self):
        """The organiser paid more than was collected - they are owed money.
        Clamping this to zero would hide that."""
        assert remaining_amount(30000.0, 36000.0) == -6000.0

    def test_format_matches_the_ui(self):
        assert format_amount(1234567.0) == "1,234,567"
        assert format_amount(0.0) == "0"
        assert format_amount(-6000.0) == "-6,000"

    def test_sum_contributions_tolerates_missing_and_none(self):
        roster = [{"amount": 3000.0}, {}, {"amount": None}, {"amount": 2000.0}]
        assert sum_contributions(roster) == 5000.0

    @pytest.mark.parametrize("value,counted", [
        ("Yes", True), ("yes", True), (" YES ", True),
        ("No", False), ("", False), (None, False), ("maybe", False),
    ])
    def test_attendance_counting_is_case_and_space_insensitive(self, value, counted):
        assert count_actual_attendees([{"actual_attend": value}]) == (1 if counted else 0)

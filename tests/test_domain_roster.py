"""Turning a recipient list into the people to actually track.

A row may be an Exchange distribution list. Expanding one needs Outlook; the
merge rules do not, which is why they live in rsvp/domain/roster.py and can be
tested here with a fake expander.
"""
from __future__ import annotations

from rsvp.domain import merge_expanded_roster


def expander(mapping):
    """A fake address book. Returns members for groups, None for people."""
    return lambda email: mapping.get((email or "").lower())


def test_a_plain_recipient_is_kept_as_is():
    people = [("Alice", "alice@example.com")]
    assert merge_expanded_roster(people, expander({})) == people


def test_a_group_is_replaced_by_its_members():
    result = merge_expanded_roster(
        [("Team", "team@example.com")],
        expander({"team@example.com": [("Alice", "alice@example.com"),
                                       ("Bob", "bob@example.com")]}))
    assert result == [("Alice", "alice@example.com"), ("Bob", "bob@example.com")]


def test_someone_in_two_groups_appears_once():
    result = merge_expanded_roster(
        [("A", "a@example.com"), ("B", "b@example.com")],
        expander({"a@example.com": [("Alice", "alice@example.com")],
                  "b@example.com": [("Alice", "alice@example.com"),
                                    ("Bob", "bob@example.com")]}))
    assert result == [("Alice", "alice@example.com"), ("Bob", "bob@example.com")]


def test_duplicates_differing_only_in_case_are_one_person():
    """Exchange returns casing inconsistently; the same colleague must not be
    invited twice or counted twice in the response totals."""
    result = merge_expanded_roster(
        [("Alice", "Alice@Example.com"), ("alice", "alice@example.com")],
        expander({}))
    assert len(result) == 1


def test_order_is_first_seen():
    result = merge_expanded_roster(
        [("Z", "z@example.com"), ("Team", "team@example.com")],
        expander({"team@example.com": [("A", "a@example.com")]}))
    assert result == [("Z", "z@example.com"), ("A", "a@example.com")]


def test_an_expansion_failure_keeps_the_person_rather_than_dropping_them():
    """A COM error means "could not expand", not "this person is not invited".

    Dropping them would silently remove someone from the tracking list, and
    the organiser would never know the invitation went unaccounted for.
    """
    def boom(email):
        raise RuntimeError("COM call failed")

    result = merge_expanded_roster([("Alice", "alice@example.com")], boom)
    assert result == [("Alice", "alice@example.com")]


def test_each_address_is_expanded_only_once_across_calls():
    calls = []

    def counting(email):
        calls.append(email)
        return None

    cache: dict = {}
    people = [("Alice", "alice@example.com")]
    merge_expanded_roster(people, counting, cache=cache)
    merge_expanded_roster(people, counting, cache=cache)
    assert calls == ["alice@example.com"], "the address book was queried twice"


def test_an_empty_recipient_list_yields_an_empty_roster():
    assert merge_expanded_roster([], expander({})) == []

"""Event writes in rsvp/storage/db.py that the UI relies on to keep data.

update_event is the primitive for automatic writes (scan counts, reminder and
calendar times, Amount paid): it must never create a row. rename_event backs
an Event ID edit on Tab 7: it must carry every column and every per-event
table across, because the old rename (insert the visible columns under the
new ID, delete the old row) silently dropped the rest.
"""
from __future__ import annotations

import pytest

from rsvp.storage import db


@pytest.fixture
def path(tmp_path):
    return str(tmp_path / "rsvp_data.db")


def _seed(path, event_id="EV1"):
    db.save_event_record({"EventID": event_id, "EventName": "Party",
                          "SentDate": "2026-09-01 10:00", "Balance": "legacy 123",
                          "AmountPaid": "5000"}, path)
    db.save_recipients(event_id, [("Example Person", "person@example.com")], path)
    db.save_gift_roster(event_id, {"person@example.com": {
        "name": "Example Person", "checked": True, "amount": 3000.0}}, path)
    db.save_attendance_roster(event_id, {"person@example.com": {
        "name": "Example Person", "vote": "Yes", "actual_attend": "Yes",
        "free": False, "amount": 3000.0}}, path)
    db.save_responses(event_id, {"person@example.com": {
        "name": "Example Person", "vote": "Yes", "received": None}}, None, path)


def _row(path, event_id):
    return next((r for r in db.load_history(path) if r["EventID"] == event_id), None)


class TestUpdateEvent:
    def test_updates_only_the_given_columns(self, path):
        _seed(path)
        assert db.update_event("EV1", {"Yes": 4, "NoResponse": 1}, path) is True
        row = _row(path, "EV1")
        assert (row["Yes"], row["NoResponse"]) == ("4", "1")
        assert row["SentDate"] == "2026-09-01 10:00"
        assert row["Balance"] == "legacy 123"

    def test_never_creates_a_row(self, path):
        assert db.update_event("NOPE", {"Yes": 4}, path) is False
        assert db.load_history(path) == []

    def test_save_event_record_still_merges(self, path):
        _seed(path)
        db.save_event_record({"EventID": "EV1", "Location": "Room 1"}, path)
        row = _row(path, "EV1")
        assert row["Location"] == "Room 1" and row["EventName"] == "Party"


class TestRenameEvent:
    def test_carries_every_column_and_every_table(self, path):
        _seed(path)
        assert db.rename_event("EV1", "EV2", path) is True
        assert _row(path, "EV1") is None
        row = _row(path, "EV2")
        # Not shown on Tab 7 - the old insert-then-delete rename lost these.
        assert row["Balance"] == "legacy 123"
        assert row["AmountPaid"] == "5000"
        assert db.load_recipients("EV2", path) == [("Example Person", "person@example.com")]
        assert db.load_gift_roster("EV2", path)["person@example.com"]["checked"] is True
        assert db.load_attendance_roster("EV2", path)["person@example.com"]["amount"] == 3000.0
        assert db.load_responses("EV2", path)[0]["person@example.com"]["vote"] == "Yes"
        for loader in (db.load_recipients, db.load_gift_roster, db.load_attendance_roster):
            assert not loader("EV1", path)
        assert db.load_responses("EV1", path)[0] == {}

    def test_refuses_a_target_that_holds_data(self, path):
        _seed(path, "EV1")
        db.save_recipients("EV2", [("Other Person", "other@example.com")], path)
        with pytest.raises(ValueError):
            db.rename_event("EV1", "EV2", path)
        # Nothing moved: the transaction rolled back whole.
        assert _row(path, "EV1") is not None
        assert db.load_recipients("EV1", path) == [("Example Person", "person@example.com")]

    def test_refuses_an_empty_target(self, path):
        _seed(path)
        with pytest.raises(ValueError):
            db.rename_event("EV1", "  ", path)

    def test_unknown_source_changes_nothing(self, path):
        _seed(path)
        assert db.rename_event("NOPE", "EV9", path) is False
        assert _row(path, "EV9") is None

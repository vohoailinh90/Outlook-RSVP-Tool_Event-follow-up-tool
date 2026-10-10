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


# The schema of the separately developed copy of this app (the
# "Event-Invitation" lineage) whose database the user runs, as its db.py
# creates it. A database it wrote must open here, and stay readable by it.
LEGACY_SCHEMA = """
CREATE TABLE events (
    EventID TEXT PRIMARY KEY,
    EventName TEXT, EventDate TEXT, Deadline TEXT, Location TEXT, Budget TEXT,
    EmailLanguage TEXT, OrganizerNote TEXT, RecipientFile TEXT, SentDate TEXT, UpdateInviteDate TEXT,
    TotalInvited TEXT, Yes TEXT, No TEXT, Maybe TEXT, NoResponse TEXT,
    ReportFile TEXT, CalendarSent TEXT,
    ActualAttendees TEXT, CostPerPerson TEXT, TotalIncome TEXT, TotalExpense TEXT, Balance TEXT,
    ReminderSent TEXT, LastReminderSentDate TEXT,
    EventMode TEXT, Organizer TEXT, GuestOfHonor TEXT, GiftBudget TEXT, GiftDeadline TEXT,
    StartTime TEXT, EndTime TEXT,
    AmountPaid TEXT,
    Round1Label TEXT,
    GiftItemName TEXT, GiftItemLink TEXT, GiftItemPrice TEXT,
    GiftLinkEvent TEXT,
    DeptFundRemaining TEXT,
    LastScanTime TEXT,
    RowOrder INTEGER
);
CREATE TABLE recipients (EventID TEXT NOT NULL, Email TEXT NOT NULL, Name TEXT,
    PRIMARY KEY (EventID, Email));
CREATE TABLE gift_contributions (EventID TEXT NOT NULL, Email TEXT NOT NULL, Name TEXT,
    Checked INTEGER DEFAULT 0, Amount REAL DEFAULT 0, SendEmail INTEGER DEFAULT 0,
    ManualAmount INTEGER DEFAULT 0, PRIMARY KEY (EventID, Email));
CREATE TABLE attendance (EventID TEXT NOT NULL, Email TEXT NOT NULL, Name TEXT, Vote TEXT,
    ActualAttend TEXT, Free INTEGER DEFAULT 0, Amount REAL DEFAULT 0, PRIMARY KEY (EventID, Email));
CREATE TABLE attendance_rounds (EventID TEXT NOT NULL, RoundKey TEXT NOT NULL, RoundLabel TEXT,
    AmountPaid TEXT, SortOrder INTEGER DEFAULT 0, PRIMARY KEY (EventID, RoundKey));
CREATE TABLE attendance_extra_amounts (EventID TEXT NOT NULL, Email TEXT NOT NULL,
    RoundKey TEXT NOT NULL, Amount REAL DEFAULT 0, Attend TEXT,
    PRIMARY KEY (EventID, Email, RoundKey));
CREATE TABLE responses (EventID TEXT NOT NULL, Email TEXT NOT NULL, Name TEXT, Vote TEXT,
    ReceivedAt TEXT, Manual INTEGER DEFAULT 0, PRIMARY KEY (EventID, Email));
CREATE TABLE app_settings (Key TEXT PRIMARY KEY, Value TEXT);
"""


@pytest.fixture
def legacy_path(tmp_path):
    """A database as the other lineage leaves it: two payment rounds, a gift
    item, a saved column order. Example data only."""
    import sqlite3
    path = str(tmp_path / "legacy.db")
    conn = sqlite3.connect(path)
    conn.executescript(LEGACY_SCHEMA)
    conn.execute(
        "INSERT INTO events (EventID, EventName, AmountPaid, Round1Label, GiftItemName, "
        "GiftItemLink, GiftItemPrice, GiftLinkEvent, Balance, DeptFundRemaining, RowOrder) "
        "VALUES ('EV1', 'Party', '70000', 'Dinner', 'Speaker', 'https://shop.example.com/item', "
        "'3,570 JPY', 'Yes', '8,430', '8,430', 1)")
    conn.executemany("INSERT INTO attendance VALUES ('EV1', ?, ?, 'Yes', ?, 0, ?)", [
        ("a@example.com", "Person A", "Yes", 6000.0), ("b@example.com", "Person B", "Yes", 6000.0)])
    conn.execute("INSERT INTO attendance_rounds VALUES ('EV1', 'round_2', 'Karaoke', '3570', 0)")
    conn.executemany("INSERT INTO attendance_extra_amounts VALUES ('EV1', ?, 'round_2', ?, ?)", [
        ("a@example.com", 2000.0, "Yes"), ("b@example.com", 0.0, "No")])
    conn.execute("INSERT INTO gift_contributions VALUES ('EV1', 'a@example.com', 'Person A', 1, 1500, 0, 1)")
    conn.execute("INSERT INTO app_settings VALUES ('history_column_order', 'EventName,EventID')")
    conn.commit()
    conn.close()
    return path


def _legacy_read(path):
    """What the other lineage's loaders read, with its own SQL."""
    import sqlite3
    conn = sqlite3.connect(path)
    try:
        q = lambda sql: conn.execute(sql).fetchall()
        return {
            "event": q("SELECT AmountPaid, Round1Label, GiftItemName, GiftItemLink, GiftItemPrice, "
                       "GiftLinkEvent, Balance FROM events WHERE EventID = 'EV1'"),
            "attendance": q("SELECT Email, ActualAttend, Free, Amount FROM attendance "
                            "WHERE EventID = 'EV1' ORDER BY rowid"),
            "rounds": q("SELECT RoundKey, RoundLabel, AmountPaid FROM attendance_rounds "
                        "WHERE EventID = 'EV1' ORDER BY SortOrder"),
            "extra": sorted(q("SELECT Email, RoundKey, Amount, Attend FROM attendance_extra_amounts "
                              "WHERE EventID = 'EV1'")),
            "gift": q("SELECT Email, Checked, Amount, SendEmail, ManualAmount FROM gift_contributions "
                      "WHERE EventID = 'EV1'"),
            "settings": q("SELECT Key, Value FROM app_settings"),
        }
    finally:
        conn.close()


class TestLegacyDatabase:
    def test_reads_rounds_gift_item_and_settings(self, legacy_path):
        row = _row(legacy_path, "EV1")
        assert (row["Round1Label"], row["GiftItemPrice"], row["GiftLinkEvent"]) == (
            "Dinner", "3,570 JPY", "Yes")
        assert db.load_attendance_rounds("EV1", legacy_path) == [
            {"key": "round_2", "label": "Karaoke", "amount_paid": "3570"}]
        roster = db.load_attendance_roster("EV1", legacy_path)
        assert roster["a@example.com"]["extra_amounts"] == {"round_2": 2000.0}
        assert roster["b@example.com"]["extra_attends"] == {"round_2": "No"}
        assert db.load_gift_roster("EV1", legacy_path)["a@example.com"]["manual_amount"] is True
        assert db.get_setting("history_column_order", None, legacy_path) == "EventName,EventID"

    def test_a_round_trip_leaves_it_readable_by_the_other_lineage(self, legacy_path):
        before = _legacy_read(legacy_path)
        db.save_attendance("EV1", db.load_attendance_roster("EV1", legacy_path),
                           db.load_attendance_rounds("EV1", legacy_path), legacy_path)
        db.save_gift_roster("EV1", db.load_gift_roster("EV1", legacy_path), legacy_path)
        db.update_event("EV1", {"EventName": "Party 2"}, legacy_path)
        assert _legacy_read(legacy_path) == before

    def test_a_database_from_before_the_merge_gains_the_new_tables(self, tmp_path):
        import sqlite3
        path = str(tmp_path / "old.db")
        conn = sqlite3.connect(path)
        conn.executescript(LEGACY_SCHEMA.split("CREATE TABLE attendance_rounds")[0].replace(
            "    Round1Label TEXT,\n    GiftItemName TEXT, GiftItemLink TEXT, GiftItemPrice TEXT,\n"
            "    GiftLinkEvent TEXT,\n    DeptFundRemaining TEXT,\n", "").replace(
            ",\n    ManualAmount INTEGER DEFAULT 0", ""))
        conn.execute("INSERT INTO events (EventID, EventName, RowOrder) VALUES ('EV1', 'Party', 1)")
        conn.commit()
        conn.close()
        assert db.update_event("EV1", {"GiftItemName": "Speaker", "Round1Label": "Dinner"}, path)
        assert _row(path, "EV1")["GiftItemName"] == "Speaker"
        assert db.load_attendance_rounds("EV1", path) == []
        assert db.next_round_index("EV1", path) == 2


class TestPaymentRounds:
    def _roster(self, extra=None):
        return {"a@example.com": {"name": "Person A", "vote": "Yes", "actual_attend": "Yes",
                                  "free": False, "amount": 6000.0,
                                  "extra_amounts": dict(extra or {}),
                                  "extra_attends": {k: "Yes" for k in (extra or {})}}}

    def test_rounds_and_amounts_are_saved_together(self, path):
        rounds = [{"key": "round_2", "label": "Karaoke", "amount_paid": "1000"}]
        db.save_attendance("EV1", self._roster({"round_2": 2000.0}), rounds, path)
        assert db.load_attendance_rounds("EV1", path) == rounds
        assert db.load_attendance_roster("EV1", path)["a@example.com"]["extra_amounts"] == {
            "round_2": 2000.0}

    def test_removing_a_round_removes_its_amounts(self, path):
        db.save_attendance("EV1", self._roster({"round_2": 2000.0}),
                           [{"key": "round_2", "label": "Karaoke"}], path)
        db.save_attendance("EV1", self._roster(), [], path)
        import sqlite3
        conn = sqlite3.connect(path)
        assert conn.execute("SELECT COUNT(*) FROM attendance_extra_amounts").fetchone()[0] == 0
        conn.close()

    def test_amounts_for_an_unknown_round_are_neither_saved_nor_loaded(self, path):
        import sqlite3
        db.save_attendance("EV1", self._roster({"round_7": 900.0}), [], path)
        conn = sqlite3.connect(path)
        assert conn.execute("SELECT COUNT(*) FROM attendance_extra_amounts").fetchone()[0] == 0
        # Left behind by a failed delete in the other lineage:
        conn.execute("INSERT INTO attendance_extra_amounts VALUES ('EV1', 'a@example.com', 'round_7', 900, 'Yes')")
        conn.commit()
        conn.close()
        assert db.load_attendance_roster("EV1", path)["a@example.com"]["extra_amounts"] == {}

    def test_a_removed_round_key_is_never_handed_out_again(self, path):
        import sqlite3
        db.save_attendance("EV1", self._roster(), [{"key": "round_3", "label": "Late"}], path)
        assert db.next_round_index("EV1", path) == 4
        db.save_attendance("EV1", self._roster(), [], path)
        conn = sqlite3.connect(path)
        conn.execute("INSERT INTO attendance_extra_amounts VALUES ('EV1', 'a@example.com', 'round_5', 1, 'Yes')")
        conn.commit()
        conn.close()
        # round_5 survives only as an orphan amount; 5 is still taken.
        assert db.next_round_index("EV1", path) == 6
        assert db.next_round_index("EV2", path) == 2

    def test_saving_only_the_rounds_keeps_the_people(self, path):
        db.save_attendance("EV1", self._roster({"round_2": 2000.0}),
                           [{"key": "round_2", "label": "Karaoke"}], path)
        db.save_attendance("EV1", None, [{"key": "round_2", "label": "Bar", "amount_paid": "500"}], path)
        roster = db.load_attendance_roster("EV1", path)
        assert roster["a@example.com"]["extra_amounts"] == {"round_2": 2000.0}
        assert db.load_attendance_rounds("EV1", path)[0]["label"] == "Bar"

    def test_rename_moves_the_rounds(self, path):
        _seed(path)
        db.save_attendance("EV1", self._roster({"round_2": 2000.0}),
                           [{"key": "round_2", "label": "Karaoke"}], path)
        db.rename_event("EV1", "EV2", path)
        assert db.load_attendance_rounds("EV2", path)[0]["label"] == "Karaoke"
        assert db.load_attendance_roster("EV2", path)["a@example.com"]["extra_amounts"] == {
            "round_2": 2000.0}
        assert db.load_attendance_rounds("EV1", path) == []


def test_a_typed_gift_amount_keeps_its_flag(path):
    db.save_gift_roster("EV1", {"a@example.com": {"name": "Person A", "checked": True,
                                                  "amount": 1500.0, "manual_amount": True}}, path)
    assert db.load_gift_roster("EV1", path)["a@example.com"]["manual_amount"] is True


def test_settings_round_trip(path):
    assert db.get_setting("history_column_order", "default", path) == "default"
    db.set_setting("history_column_order", "EventID,EventName", path)
    assert db.get_setting("history_column_order", None, path) == "EventID,EventName"

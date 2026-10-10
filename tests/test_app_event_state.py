"""RSVPApp keeps each event's data with that event.

These drive the real Tk application against tests/fake_outlook.py, in a
temporary directory (so its rsvp_data.db is a throwaway), with message boxes
and file dialogs replaced and background workers run inline. They need
tkinter, pywin32's import and a display: they run on the Windows CI and skip
where any of those is missing, so they never report a pass they did not run.

Each test pins a way one event's data used to end up in another, or be lost:
state carried over by '⬅ Load setup', auto-saves keyed on whatever Event ID
Tab 1 happened to show, a reminder or calendar invite built from another
event's vote table, Tab 7 saves reverting values written meanwhile, and an
Event ID rename that dropped the event's data.
"""
from __future__ import annotations

import types
from datetime import datetime, timezone

import pytest

from rsvp.storage import db
from tests.fake_outlook import FakeOutlook


class _InlineThread:
    """after() from another thread needs a running mainloop; the tests pump
    update() instead, so workers run on the calling thread."""

    def __init__(self, target=None, daemon=None, **_):
        self.target = target

    def start(self):
        self.target()


class ScriptedOutlook(FakeOutlook):
    def __init__(self):
        super().__init__()
        self.scan_result = {}

    def scan_voting_responses(self, event_id, folder_paths=None, scan_all=False):
        self._record("scan_voting_responses", event_id=event_id,
                     folder_paths=folder_paths, scan_all=scan_all)
        return {k: dict(v) for k, v in self.scan_result.items()}, 0

    def sends(self):
        return [name for name, _ in self.calls if name.startswith("send_")]


@pytest.fixture
def app(monolith, monkeypatch, tmp_path):
    import tkinter

    monkeypatch.chdir(tmp_path)
    dialogs = []

    def recorder(kind, answer=None):
        def record(title, message=None, **_):
            dialogs.append((kind, title, message))
            return answer
        return record

    for kind in ("showinfo", "showwarning", "showerror"):
        monkeypatch.setattr(monolith.messagebox, kind, recorder(kind))
    monkeypatch.setattr(monolith.messagebox, "askyesno", recorder("askyesno", True))
    monkeypatch.setattr(monolith, "threading", types.SimpleNamespace(Thread=_InlineThread))

    outlook = ScriptedOutlook()
    try:
        instance = monolith.RSVPApp(outlook=outlook)
    except tkinter.TclError as exc:
        pytest.skip(f"no display: {exc}")
    # Tk reports an exception raised in a callback and carries on, which
    # would let a test pass over a crash; collect them and fail instead.
    callback_errors = []
    instance.report_callback_exception = lambda *exc_info: callback_errors.append(exc_info[1])
    instance.dialogs = dialogs
    instance.fake = outlook
    instance.db_path = instance.history_path.get()
    instance.update()
    yield instance
    instance.update()
    instance.destroy()
    assert not callback_errors, f"a Tk callback raised: {callback_errors!r}"


def _row(app, event_id):
    return next((r for r in db.load_history(app.db_path) if r["EventID"] == event_id), None)


def _select(app, tab):
    app.nb.select(tab)
    app.update()


def _start_event(app, event_id="EV1"):
    """EV1 in History with three recipients, scanned: Alice Yes, Bob No."""
    app.var_event_id.set(event_id)
    app.var_event_name.set("Party")
    app.var_budget.set("3,000 JPY")
    app.recipients = [("Alice Example", "alice@example.com"),
                      ("Bob Example", "bob@example.com"),
                      ("Carol Example", "carol@example.com")]
    app._refresh_recipient_tree()
    app._update_history_from_tab1()
    when = datetime(2026, 10, 1, 9, 0, tzinfo=timezone.utc)  # pywin32 returns aware times
    app.fake.scan_result = {
        "alice@example.com": {"name": "Alice Example", "vote": "Yes", "received": when},
        "bob@example.com": {"name": "Bob Example", "vote": "No", "received": when},
    }
    app._collect_responses()
    app.update()


def _load(app, event_id):
    labels = list(app._refresh_history_combo_values())
    app.combo_load_history.current(
        next(i for i, label in enumerate(labels) if label.startswith(f"{event_id} — ")))
    app._load_from_history()
    app.update()


def test_a_scan_saves_votes_and_counts_for_the_scanned_event(app):
    _start_event(app)
    row = _row(app, "EV1")
    assert (row["Yes"], row["No"], row["NoResponse"], row["TotalInvited"]) == ("1", "1", "1", "3")
    assert db.load_responses("EV1", app.db_path)[0]["alice@example.com"]["vote"] == "Yes"


def test_a_manual_vote_survives_a_rescan_with_an_older_email(app):
    _start_event(app)
    carol = next(i for i in app.tree_responses.get_children()
                 if app.tree_responses.set(i, "email") == "carol@example.com")
    app._commit_response_vote_edit(carol, "vote", "Maybe")
    app.fake.scan_result["carol@example.com"] = {
        "name": "Carol Example", "vote": "No",
        "received": datetime(2020, 1, 1, tzinfo=timezone.utc)}
    app._collect_responses()   # aware vs naive times: used to raise TypeError
    app.update()
    assert app.responses["carol@example.com"]["vote"] == "Maybe"
    assert _row(app, "EV1")["Maybe"] == "1"


def test_an_empty_scan_does_not_wipe_saved_votes(app):
    _start_event(app)
    app.fake.scan_result = {}
    app._collect_responses()
    app.update()
    assert db.load_responses("EV1", app.db_path)[0]["alice@example.com"]["vote"] == "Yes"
    assert _row(app, "EV1")["Yes"] == "1"


def _restart(app):
    """What the app holds after a restart: nothing in memory, the database
    as it was. The user types the Event ID back and has the recipients."""
    app.var_event_id.set("")
    app._clear_event_state()
    app.var_event_id.set("EV1")
    app.recipients = [("Alice Example", "alice@example.com"),
                      ("Bob Example", "bob@example.com"),
                      ("Carol Example", "carol@example.com")]
    app._refresh_recipient_tree()


def test_an_empty_first_scan_after_a_restart_keeps_saved_votes(app):
    _start_event(app)
    _restart(app)
    app.fake.scan_result = {}
    app._collect_responses()
    app.update()
    assert db.load_responses("EV1", app.db_path)[0]["alice@example.com"]["vote"] == "Yes"
    row = _row(app, "EV1")
    assert (row["Yes"], row["No"]) == ("1", "1")


def test_the_first_scan_after_a_restart_keeps_saved_manual_votes(app):
    _start_event(app)
    carol = next(i for i in app.tree_responses.get_children()
                 if app.tree_responses.set(i, "email") == "carol@example.com")
    app._commit_response_vote_edit(carol, "vote", "Maybe")
    _restart(app)
    app._collect_responses()   # finds Alice and Bob again, nothing from Carol
    app.update()
    assert app.responses["carol@example.com"]["vote"] == "Maybe"
    assert db.load_responses("EV1", app.db_path)[0]["carol@example.com"]["manual"] is True


def test_loading_an_event_carries_nothing_over_and_keeps_amount_paid(app):
    _start_event(app)
    _select(app, app.tab_calendar)
    app.var_amount_paid.set("5000")
    app.update()
    assert _row(app, "EV1")["AmountPaid"] == "5000"
    _select(app, app.tab_gift)
    app._commit_gift_amount_edit("alice@example.com", "amount", "1,000")
    db.save_event_record({"EventID": "EV2", "EventName": "Other"}, app.db_path)

    _load(app, "EV2")

    assert app.var_event_id.get() == "EV2"
    assert app.recipients == [] and app.responses == {} and app._pending_recipients == []
    _select(app, app.tab_calendar)
    assert app._attendance_roster == {} and app.var_amount_paid.get() == "0"
    _select(app, app.tab_gift)
    assert app._gift_roster == {}
    # Resetting Amount paid on load must not be saved over EV1's figure.
    assert _row(app, "EV1")["AmountPaid"] == "5000"
    assert db.load_gift_roster("EV1", app.db_path)["alice@example.com"]["amount"] == 1000.0


def test_a_new_event_id_on_tab1_does_not_inherit_attendance_or_gift_ticks(app):
    _start_event(app)
    _select(app, app.tab_calendar)
    assert "alice@example.com" in app._attendance_roster
    _select(app, app.tab_gift)
    app._commit_gift_amount_edit("alice@example.com", "amount", "1,000")

    app.var_event_id.set("EV2")   # the "use as a template" flow
    _select(app, app.tab_calendar)
    assert app._attendance_roster == {}
    _select(app, app.tab_gift)
    assert not any(info["checked"] for info in app._gift_roster.values())
    assert db.load_gift_roster("EV2", app.db_path) == {}


def test_reminder_and_calendar_refuse_another_events_vote_table(app):
    _start_event(app)
    app.var_event_id.set("EV2")
    app._send_reminder()
    app._send_calendar()
    assert app.fake.sends() == []
    assert sum(1 for _, title, _ in app.dialogs if title == "Scan this event first") == 2


def test_tab7_saves_only_edited_cells(app):
    _start_event(app)
    app._commit_history_edit("EV1", "Location", "Hall B")
    # Written meanwhile without redrawing Tab 7, so the table holds a stale
    # Yes count; saving the whole row would put it back.
    db.update_event("EV1", {"Yes": "7"}, app.db_path)
    app._save_history_edits()
    row = _row(app, "EV1")
    assert row["Location"] == "Hall B"
    assert row["Yes"] == "7"


def test_tab7_keeps_unsaved_edits_across_a_redraw(app):
    _start_event(app)
    app._commit_history_edit("EV1", "Organizer", "Example Organizer")
    app._refresh_history_tree()   # e.g. after a scan's automatic write
    assert app.tree_history.set("EV1", "Organizer") == "Example Organizer"


def test_tab7_rename_moves_the_whole_event(app):
    _start_event(app)
    db.update_event("EV1", {"Balance": "legacy", "AmountPaid": "5000"}, app.db_path)
    app._refresh_history_tree()
    app._commit_history_edit("EV1", "EventID", "EV1B")
    app._save_history_edits()
    assert _row(app, "EV1") is None
    row = _row(app, "EV1B")
    assert (row["Balance"], row["AmountPaid"]) == ("legacy", "5000")
    assert len(db.load_recipients("EV1B", app.db_path)) == 3
    assert app.var_event_id.get() == "EV1B" and app._last_scanned_event_id == "EV1B"


def test_tab7_keeps_other_edits_when_a_rename_succeeds_but_they_fail(app, monkeypatch):
    _start_event(app)
    app._commit_history_edit("EV1", "EventID", "EV1B")
    app._commit_history_edit("EV1", "Location", "Hall B")

    def locked(*_args, **_kwargs):
        raise OSError("database is locked")
    monkeypatch.setattr(db, "update_event", locked)
    app._save_history_edits()

    assert _row(app, "EV1B") is not None                       # the rename went through
    assert app.tree_history.set("EV1B", "Location") == "Hall B"  # the edit is still on screen
    assert app._history_edits == {"EV1B": {"Location": "Hall B"}}

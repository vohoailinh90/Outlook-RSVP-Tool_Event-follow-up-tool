"""Features merged from the separately developed copy of this app that the
user runs (the "Event-Invitation" lineage): payment rounds on Tab 5, the
gift item and its money on Tab 6, History's money columns, search and
column order, and HTML thank-you / report emails.

Driven like tests/test_app_event_state.py - the real Tk app against the fake
Outlook, in a throwaway database - and run where that file runs. Several
tests pin a way the merged code could lose or misplace money: rounds of one
event shown and saved under another, figures typed in the other copy
overwritten, a typed gift amount replaced by the budget.
"""
from __future__ import annotations

import sqlite3

import pytest

from rsvp.storage import db
from tests.test_app_event_state import _load, _row, _select, _start_event, app  # noqa: F401


def _attendance(app, extra_scan=True):
    """EV1 scanned with Alice Yes and Carol Maybe; Tab 5 opened."""
    _start_event(app)
    if extra_scan:
        when = app.fake.scan_result["alice@example.com"]["received"]
        app.fake.scan_result["carol@example.com"] = {
            "name": "Carol Example", "vote": "Maybe", "received": when}
        app._collect_responses()
        app.update()
    _select(app, app.tab_calendar)


def _add_round(app, monolith, monkeypatch, label):
    monkeypatch.setattr(monolith.simpledialog, "askstring", lambda *a, **kw: label)
    app._add_amount_round()
    app.update()
    return app._extra_rounds[-1]["key"]


# ── Tab 5: payment rounds ──────────────────────────────────────────────

def test_a_round_is_saved_with_its_amounts(app, monolith, monkeypatch):
    _attendance(app)
    key = _add_round(app, monolith, monkeypatch, "Karaoke")
    app._commit_attendance_edit("alice@example.com", f"attend_{key}", "Yes")
    app._round_vars[key]["paid"].set("1,500")
    app.update()

    assert db.load_attendance_rounds("EV1", app.db_path) == [
        {"key": key, "label": "Karaoke", "amount_paid": "1,500"}]
    alice = db.load_attendance_roster("EV1", app.db_path)["alice@example.com"]
    assert alice["extra_amounts"][key] == 3000.0    # the expected event budget
    assert app.var_total_collected_amount.get() == "9,000"  # 3,000 x 2 in round 1 + 3,000


def test_free_exempts_someone_from_every_round(app, monolith, monkeypatch):
    _attendance(app)
    key = _add_round(app, monolith, monkeypatch, "Karaoke")
    app._commit_attendance_edit("alice@example.com", f"attend_{key}", "Yes")
    app._commit_attendance_edit("alice@example.com", "free", "Yes")
    alice = app._attendance_roster["alice@example.com"]
    assert (alice["amount"], alice["extra_amounts"][key]) == (0.0, 0.0)


def test_removing_a_round_removes_its_amounts(app, monolith, monkeypatch):
    _attendance(app)
    key = _add_round(app, monolith, monkeypatch, "Karaoke")
    app._commit_attendance_edit("alice@example.com", f"attend_{key}", "Yes")
    app._remove_amount_round(key)    # askyesno answers Yes in this fixture
    app.update()
    assert db.load_attendance_rounds("EV1", app.db_path) == []
    conn = sqlite3.connect(app.db_path)
    assert conn.execute("SELECT COUNT(*) FROM attendance_extra_amounts").fetchone()[0] == 0
    conn.close()
    assert key not in app.tree_attendance["columns"][-1]


def test_another_event_id_on_tab1_does_not_inherit_the_rounds(app, monolith, monkeypatch):
    """Critic H2: only the roster used to swap owners when Tab 1's Event ID
    changed; rounds left on screen would be saved under the new event by
    its next keystroke."""
    _attendance(app)
    _add_round(app, monolith, monkeypatch, "Karaoke")

    _select(app, app.tab_config)
    app.var_event_id.set("EV2")
    _select(app, app.tab_calendar)
    assert app._extra_rounds == [] and app._round1_label == ""
    app.var_amount_paid.set("100")   # a keystroke that saves
    app.update()
    assert db.load_attendance_rounds("EV2", app.db_path) == []
    # EV1 still has its round, and gets it back when it is loaded.
    db.save_event_record({"EventID": "EV2", "EventName": "Other"}, app.db_path)
    _load(app, "EV1")
    _select(app, app.tab_calendar)
    assert [r["label"] for r in app._extra_rounds] == ["Karaoke"]


def test_the_first_rounds_name_waits_for_a_history_row(app, monolith, monkeypatch):
    """Critic H3: Round1Label lives on the History row; with none yet it is
    reported, then written by Tab 1's save."""
    _attendance(app)
    _select(app, app.tab_config)
    app.var_event_id.set("NEW")
    _select(app, app.tab_calendar)
    monkeypatch.setattr(monolith.simpledialog, "askstring", lambda *a, **kw: "Dinner")
    app._rename_round(None)
    assert any(d[1] == "Not saved" for d in app.dialogs)
    app._update_history_from_tab1()
    assert _row(app, "NEW")["Round1Label"] == "Dinner"


def test_the_thank_you_email_has_the_rounds_table_and_goes_as_html(app, monolith, monkeypatch):
    _attendance(app)
    key = _add_round(app, monolith, monkeypatch, "Karaoke")
    app._commit_attendance_edit("alice@example.com", f"attend_{key}", "Yes")
    app._apply_thankyou_body_lang()
    body = app.txt_thankyou_body.get("1.0", "end")
    assert "Karaoke" in body and "│" in body

    app._send_thank_you_email()
    app.update()
    name, call = next(c for c in app.fake.calls if c[0] == "send_thankyou_email")
    assert "<table" in call["html_body"] and "Karaoke" in call["html_body"]


def test_the_excel_report_has_a_column_pair_per_round(app, monolith, monkeypatch):
    _attendance(app)
    _add_round(app, monolith, monkeypatch, "Karaoke")
    ws = app._build_attendance_workbook().active
    header = [c.value for c in ws[1]]
    assert header[-2:] == ["Karaoke Attend", "Karaoke"]


# ── History's money columns ────────────────────────────────────────────

def test_history_money_follows_tab_5_and_tab_6(app, monolith, monkeypatch):
    _attendance(app)
    app.var_amount_paid.set("4000")
    app.update()
    _select(app, app.tab_gift)
    app.var_gift_item_price.set("1,000")
    app.update()
    row = _row(app, "EV1")
    # Round 1: Alice and Carol attend at 3,000; 4,000 paid out; gift 1,000.
    assert (row["ActualAttendees"], row["CostPerPerson"]) == ("2", "3,000")
    assert (row["TotalIncome"], row["TotalExpense"], row["Balance"]) == ("6,000", "5,000", "1,000")


def test_figures_typed_in_the_other_copy_survive_untouched(app):
    """Critic H4: an event with no attendance or gift data here keeps the
    money columns the other copy of the app wrote; reading History never
    writes, so Dept. Fund Left is not stored either."""
    db.save_event_record({"EventID": "OLD1", "EventName": "Old", "Balance": "500",
                          "TotalIncome": "9,999"}, app.db_path)
    db.save_event_record({"EventID": "OLD2", "EventName": "Older", "Balance": "-200"}, app.db_path)
    _start_event(app)
    _select(app, app.tab_history)
    app._refresh_history_tree()
    old1 = _row(app, "OLD1")
    assert (old1["Balance"], old1["TotalIncome"], old1["DeptFundRemaining"]) == ("500", "9,999", "")
    shown = dict(zip(app.tree_history["columns"], app.tree_history.item("OLD2", "values")))
    assert shown["DeptFundRemaining"] == "300"


def test_an_unreadable_balance_is_reported_not_guessed(app):
    db.save_event_record({"EventID": "OLD1", "EventName": "Old", "Balance": "about 500"}, app.db_path)
    _select(app, app.tab_history)
    app._refresh_history_tree()
    shown = dict(zip(app.tree_history["columns"], app.tree_history.item("OLD1", "values")))
    assert shown["DeptFundRemaining"] == "?"
    assert "OLD1" in app.lbl_history_note.cget("text")


def test_calculated_columns_cannot_be_edited_on_tab_7(app):
    _start_event(app)
    _select(app, app.tab_history)
    app._commit_history_edit("EV1", "Balance", "123")
    app._commit_history_edit("EV1", "Location", "Room 2")
    assert app._history_edits == {"EV1": {"Location": "Room 2"}}


def test_a_column_order_saved_by_the_other_copy_is_healed(app):
    """Critic M3: unknown names (ReminderSent, which this copy does not
    show) are dropped and missing columns appended, or Tk refuses the
    order and the table draws nothing."""
    db.set_setting("history_column_order", "Balance,ReminderSent,EventName,EventID", app.db_path)
    _select(app, app.tab_history)
    app._apply_history_column_order()
    shown = app._tree_display_columns(app.tree_history)
    assert shown[:3] == ("Balance", "EventName", "EventID")
    assert "ReminderSent" not in shown and len(shown) == len(app.tree_history["columns"])


def test_a_dragged_order_is_remembered(app):
    _select(app, app.tab_history)
    order = list(app.tree_history["columns"])
    order.insert(0, order.pop(order.index("Balance")))
    app._on_history_columns_reordered(order)
    assert db.get_setting("history_column_order", None, app.db_path).startswith("Balance,EventID")


# ── Tab 6: the gift item and typed amounts ─────────────────────────────

def test_the_gift_item_is_saved_and_reloaded(app):
    _start_event(app)
    _select(app, app.tab_gift)
    app.var_gift_item_name.set("Speaker")
    app.var_gift_item_link.set("https://shop.example.com/item")
    app.var_gift_link_event.set(True)
    app._on_gift_link_event_toggled()
    assert (_row(app, "EV1")["GiftItemName"], _row(app, "EV1")["GiftLinkEvent"]) == ("Speaker", "Yes")
    db.save_event_record({"EventID": "EV2", "EventName": "Other"}, app.db_path)
    _load(app, "EV2")
    assert app.var_gift_item_name.get() == ""
    _load(app, "EV1")
    assert app.var_gift_item_name.get() == "Speaker" and app.var_gift_link_event.get() is True


def test_the_gift_report_has_the_gift_and_goes_as_html(app):
    _start_event(app)
    _select(app, app.tab_gift)
    app.var_gift_item_name.set("Speaker")
    app.var_gift_item_price.set("3,570")
    app._gift_roster["alice@example.com"]["send_email"] = True
    app._generate_gift_report_draft()
    assert "Speaker" in app.txt_gift_report_body.get("1.0", "end")
    app._send_gift_report_email()
    app.update()
    _name, call = next(c for c in app.fake.calls if c[0] == "send_gift_report_email")
    assert "Speaker" in call["html_body"] and call["html_body"].startswith("<div")


def test_a_typed_gift_amount_survives_ticking_contributed_again(app):
    _start_event(app)
    app.var_gift_budget.set("1,000")
    _select(app, app.tab_gift)
    app._commit_gift_amount_edit("alice@example.com", "amount", "2,500")
    app._toggle_all_gift_column("check")   # tick everyone shown
    assert app._gift_roster["alice@example.com"]["amount"] == 2500.0
    assert app._gift_roster["bob@example.com"]["amount"] == 1000.0
    assert db.load_gift_roster("EV1", app.db_path)["alice@example.com"]["manual_amount"] is True


@pytest.mark.parametrize("typed", ["abc", "-500", "1 000", "USD -500"])
def test_a_gift_amount_that_is_not_a_positive_number_is_refused(app, typed):
    _start_event(app)
    _select(app, app.tab_gift)
    app._commit_gift_amount_edit("alice@example.com", "amount", "2,500")
    app._commit_gift_amount_edit("alice@example.com", "amount", typed)
    assert app._gift_roster["alice@example.com"]["amount"] == 2500.0
    assert app.dialogs[-1][0] == "showwarning"


# ── Tab 1: search ──────────────────────────────────────────────────────

def test_search_filters_the_list_and_loads_the_event_shown(app):
    for i, place in enumerate(["Hall A", "Hall B", "Rooftop"], start=1):
        db.save_event_record({"EventID": f"EV{i}", "EventName": f"Party {i}", "Location": place},
                             app.db_path)
    app._refresh_history_combo_values()
    app.var_history_search_field.set("Location")
    app.var_history_search.set("roof")
    app.update()
    assert list(app.combo_load_history["values"]) == ["EV3 — Party 3   [Location: Rooftop]"]
    app._load_from_history()
    app.update()
    assert app.var_event_id.get() == "EV3"


# ── Tab 2: group expansion reports what it could not list ──────────────

def test_expanding_keeps_a_group_not_every_member_of_which_could_be_listed(app):
    """Codex review of PR #12: saving the members it could list dropped the
    rest for good - the group's address, the only way back to them, was
    gone. The group is kept as one row and the dialog says why."""
    app.recipients = [("Team", "team@example.com")]
    app.fake.groups["team@example.com"] = (
        [("Person A", "a@example.com")], ["Nested group"], ["Team: 2 direct (1 sub-groups, 1 people)"])
    app._expand_group_recipients()
    app.update()
    title, message = app.dialogs[-1][1:]
    assert app.recipients == [("Team", "team@example.com")]
    assert "Nested group" in message and "kept as one row, not expanded" in message


def test_working_on_an_event_with_no_tracked_money_keeps_its_typed_figures(app):
    """Critic H4: an event whose money was typed in the other copy, with no
    attendance or gift data here, is not recomputed to zero when it is
    loaded and touched - opened on every page, saved on Tab 1, its gift
    item named."""
    db.save_event_record({"EventID": "OLD1", "EventName": "Old", "TotalIncome": "9,999",
                          "Balance": "500"}, app.db_path)
    _load(app, "OLD1")
    for tab in (app.tab_calendar, app.tab_gift, app.tab_history):
        _select(app, tab)
    app.var_gift_item_name.set("Speaker")
    app._save_event_from_tab1()
    app.update()
    row = _row(app, "OLD1")
    assert (row["GiftItemName"], row["TotalIncome"], row["Balance"]) == ("Speaker", "9,999", "500")


def test_what_was_paid_out_reaches_history_with_nobody_in_the_table(app):
    """Codex review of PR #12: a cancelled event's costs typed as round 1's
    Amount paid never reached History, since nothing else was tracked."""
    db.save_event_record({"EventID": "OLD1", "EventName": "Old"}, app.db_path)
    _load(app, "OLD1")
    _select(app, app.tab_calendar)
    app.var_amount_paid.set("4,000")
    app.update()
    row = _row(app, "OLD1")
    assert (row["AmountPaid"], row["TotalExpense"], row["Balance"]) == ("4,000", "4,000", "-4,000")


@pytest.mark.parametrize("typed", ["abc", "-500", "1 000", "500 + 300", "¥-500"])
def test_an_amount_cell_refuses_what_is_not_a_paid_amount(app, typed):
    """Review finding: parse_amount_from_text read "-500" as 500 and "abc"
    as 0, silently changing what someone paid."""
    _attendance(app)
    app._commit_attendance_edit("alice@example.com", "amount", "2,000")
    app._commit_attendance_edit("alice@example.com", "amount", typed)
    assert app._attendance_roster["alice@example.com"]["amount"] == 2000.0
    assert app.dialogs[-1][1] == "Amount not changed"


def test_pasting_a_blank_attend_leaves_it_untouched(app):
    _attendance(app)
    app._commit_attendance_edit("alice@example.com", "actual_attend", "")
    assert app._attendance_roster["alice@example.com"]["actual_attend"] == ""
    app._commit_attendance_edit("alice@example.com", "actual_attend", "Maybe")
    assert app._attendance_roster["alice@example.com"]["actual_attend"] == "No"


def test_a_group_none_of_whose_members_could_be_listed_keeps_its_row(app):
    """Second review: replacing such a group by nobody saved a list that had
    silently lost everyone in it. The row stays and the dialog says why."""
    app.recipients = [("Team", "team@example.com"), ("Dee Example", "dee@example.com")]
    app.fake.groups["team@example.com"] = ([], ["Team"], ["Team: 0 members listed"])
    app._expand_group_recipients()
    app.update()
    assert app.recipients == [("Team", "team@example.com"), ("Dee Example", "dee@example.com")]
    assert "Team" in app.dialogs[-1][2]


def test_a_group_that_lists_nobody_with_an_address_keeps_its_row(app):
    """Third review: members without an SMTP address are skipped without
    counting as a failure, so such a group came back as an empty list."""
    app.recipients = [("Team", "team@example.com")]
    app.fake.groups["team@example.com"] = ([], [], [])
    app._expand_group_recipients()
    app.update()
    assert app.recipients == [("Team", "team@example.com")]
    assert "listed nobody" in app.dialogs[-1][2]


@pytest.mark.parametrize("unlisted", [([], ["Team"], []), ([], [], [])])
def test_response_tracking_keeps_a_group_whose_members_could_not_be_listed(app, unlisted):
    """Tab 4 follows the same rule: a group it cannot list stays one row to
    chase, rather than vanishing from "not responded"."""
    app.fake.groups["team@example.com"] = unlisted
    app.fake.groups["club@example.com"] = ([("Person A", "a@example.com")], [], [])
    roster = app._build_effective_roster([("Team", "team@example.com"), ("Club", "club@example.com")])
    assert roster == [("Team", "team@example.com"), ("Person A", "a@example.com")]


def test_response_tracking_keeps_a_partly_listed_groups_row_for_the_rest(app):
    """Codex review of PR #12: a group with one sub-group that could not be
    listed was replaced by the people it could list, so that sub-group's
    people vanished from "not responded" and the reminders."""
    app.fake.groups["team@example.com"] = ([("Person A", "a@example.com")], ["Sub team"], [])
    roster = app._build_effective_roster([("Team", "team@example.com")])
    assert roster == [("Person A", "a@example.com"), ("Team", "team@example.com")]


def test_a_group_not_listed_yet_is_asked_again_on_the_next_scan(app):
    """Third review: the failure was cached for the session, so downloading
    the address book changed nothing until the app restarted. People and
    listed groups are still asked once."""
    people = [("Team", "team@example.com"), ("Dee Example", "dee@example.com"),
              ("Club", "club@example.com")]
    app.fake.groups["team@example.com"] = ([], ["Team"], [])
    app.fake.groups["club@example.com"] = ([("Person A", "a@example.com")], [], [])
    app._build_effective_roster(people)
    app.fake.groups["team@example.com"] = ([("Person B", "b@example.com")], [], [])
    roster = app._build_effective_roster(people)
    assert roster == [("Person B", "b@example.com"), ("Dee Example", "dee@example.com"),
                      ("Person A", "a@example.com")]
    asked = [kw["email_or_name"] for name, kw in app.fake.calls
             if name == "expand_group_members_detailed"]
    assert sorted(asked) == ["club@example.com", "dee@example.com",
                             "team@example.com", "team@example.com"]


def _press(widget, sequence):
    widget.focus_force()
    widget.update()
    widget.event_generate(sequence)
    widget.update()


def test_a_paste_reports_the_cells_it_refused_once(app):
    """One dialog for the whole paste, not one per refused cell - and every
    refused amount is left as it was."""
    _attendance(app)
    tree = app.tree_attendance
    rows = list(tree.get_children())
    for iid in rows:
        app._commit_attendance_edit(iid, "amount", "2,000")
    app._render_attendance_tree()
    tree.selection_set(rows[0])
    shown = app._tree_display_columns(tree)
    amount_at = shown.index("amount")
    lines = []
    for iid in rows:
        cells = [tree.set(iid, c) for c in shown]
        cells[amount_at] = "abc"
        lines.append("\t".join(cells))
    tree.clipboard_clear()
    tree.clipboard_append("\n".join(lines))
    before = len(app.dialogs)

    _press(tree, "<Control-v>")

    assert [d[1] for d in app.dialogs[before:]] == ["Some cells not changed"]
    assert f"{len(rows)} pasted value(s)" in app.dialogs[-1][2]
    assert all(app._attendance_roster[iid]["amount"] == 2000.0 for iid in rows)
    app._commit_attendance_edit(rows[0], "amount", "abc")   # typed again: its own dialog
    assert app.dialogs[-1][1] == "Amount not changed"


def test_a_rounds_attend_header_ticks_everyone_and_fills_their_amounts(app, monolith, monkeypatch):
    _attendance(app)
    key = _add_round(app, monolith, monkeypatch, "Karaoke")
    header = app.tree_attendance.heading(f"attend_{key}", "command")
    shown = [iid for iid in app.tree_attendance.get_children() if iid in app._attendance_roster]
    assert shown

    app.tk.call(header)
    app.update()
    saved = db.load_attendance_roster("EV1", app.db_path)
    assert {e: (saved[e]["extra_attends"][key], saved[e]["extra_amounts"][key]) for e in shown} == {
        e: ("Yes", 3000.0) for e in shown}

    app.tk.call(header)
    app.update()
    saved = db.load_attendance_roster("EV1", app.db_path)
    assert {saved[e]["extra_attends"][key] for e in shown} == {"No"}


def test_copy_follows_the_column_order_shown_on_tab_7(app):
    _start_event(app)
    _select(app, app.tab_history)
    order = list(app.tree_history["columns"])
    order.insert(0, order.pop(order.index("EventName")))
    app._on_history_columns_reordered(order)
    app._apply_history_column_order()
    tree = app.tree_history
    row = next(iid for iid in tree.get_children() if tree.set(iid, "EventID") == "EV1")
    tree.selection_set(row)

    _press(tree, "<Control-c>")

    assert tree.clipboard_get().split("\t")[:2] == ["Party", "EV1"]


def test_a_typed_amount_stays_until_attend_or_free_really_changes(app, monolith, monkeypatch):
    """Pasting back a copied row re-commits its unchanged ✅/⬜ cells; that
    used to reset every typed amount to the budget."""
    _attendance(app)
    key = _add_round(app, monolith, monkeypatch, "Karaoke")
    app._commit_attendance_edit("alice@example.com", f"attend_{key}", "Yes")
    app._commit_attendance_edit("alice@example.com", "amount", "2,000")
    app._commit_attendance_edit("alice@example.com", f"extra_{key}", "1,000")
    alice = app._attendance_roster["alice@example.com"]
    attend = alice["actual_attend"]

    app._commit_attendance_edit("alice@example.com", "actual_attend", "✅" if attend == "Yes" else "⬜")
    app._commit_attendance_edit("alice@example.com", "free", "⬜")
    app._commit_attendance_edit("alice@example.com", f"attend_{key}", "✅")
    app.tk.call(app.tree_attendance.heading(f"attend_{key}", "command"))   # ticks the others
    assert (alice["amount"], alice["extra_amounts"][key]) == (2000.0, 1000.0)

    app._commit_attendance_edit("alice@example.com", "free", "✅")
    assert (alice["amount"], alice["extra_amounts"][key]) == (0.0, 0.0)


def test_a_paid_box_not_read_as_it_looks_is_pointed_out(app, monolith, monkeypatch):
    """Third review: "1 000" counted as 1 in every total, the History, the
    email and Excel, and nothing on screen said so."""
    _attendance(app)
    key = _add_round(app, monolith, monkeypatch, "Karaoke")
    app._round_vars[key]["paid"].set("1 000")
    app.update()
    warning = app.lbl_rounds_warning.cget("text")
    assert app.lbl_rounds_warning.winfo_manager() == "pack"
    assert "Karaoke" in warning and "1 000" in warning and "Round 1" not in warning
    app._round_vars[key]["paid"].set("1,000")
    app.update()
    assert app.lbl_rounds_warning.winfo_manager() == ""


def test_a_gift_price_not_read_as_it_looks_is_pointed_out(app):
    _start_event(app)
    _select(app, app.tab_gift)
    app.var_gift_item_price.set("3 570")
    app.update()
    assert "3 570" in app.lbl_gift_price_warning.cget("text")
    app.var_gift_item_price.set("3,570 JPY")
    app.update()
    assert app.lbl_gift_price_warning.winfo_manager() == ""


def test_party_figures_that_cannot_be_read_are_not_reported_as_zero(app, monkeypatch):
    """Third review: a failed read of the party's figures gave 0 / 0 / 0, so
    a linked report showed the party as costing nothing."""
    _start_event(app)
    _select(app, app.tab_gift)
    app._gift_roster["alice@example.com"]["send_email"] = True

    def locked(*_a, **_kw):
        raise sqlite3.OperationalError("database is locked")

    monkeypatch.setattr(db, "load_attendance_roster", locked)
    app.var_gift_link_event.set(True)
    app._on_gift_link_event_toggled()
    app.update()
    assert app.var_gift_grand_collected.get() == "?"
    assert "could not be read" in app.lbl_gift_event_figures.cget("text")

    app._send_gift_report_email()
    app.update()
    assert app.dialogs[-1][1] == "Party figures unknown"
    assert not [c for c in app.fake.calls if c[0] == "send_gift_report_email"]


def test_a_table_emptied_by_new_votes_is_saved_empty(app):
    """Codex review of PR #12: a rescan turning every Yes/Maybe into No
    left the old table - and History's money - in the database, back on the
    next load, while the screen showed nobody."""
    _attendance(app)
    assert set(db.load_attendance_roster("EV1", app.db_path)) == {"alice@example.com",
                                                                  "carol@example.com"}
    assert _row(app, "EV1")["ActualAttendees"] == "2"
    when = app.fake.scan_result["alice@example.com"]["received"]
    for email in ("alice@example.com", "carol@example.com"):
        app.fake.scan_result[email] = {"name": email, "vote": "No", "received": when}
    app._collect_responses()
    _select(app, app.tab_config)
    _select(app, app.tab_calendar)

    assert db.load_attendance_roster("EV1", app.db_path) == {}
    row = _row(app, "EV1")
    assert (row["ActualAttendees"], row["TotalIncome"]) == ("0", "0")
    _load(app, "EV1")
    _select(app, app.tab_calendar)
    assert app._attendance_roster == {}


# ── Codex review of PR #12: gift totals, imported amounts, lost saves ──

def test_only_people_marked_as_contributing_count_towards_the_gift(app):
    """An import can leave an amount on someone not marked as having
    contributed; the contributor report leaves them out, so the totals, the
    email and History must too."""
    _start_event(app)
    _select(app, app.tab_gift)
    app._gift_roster["alice@example.com"].update(checked=True, amount=2000.0)
    app._gift_roster["bob@example.com"].update(checked=False, amount=5000.0)
    app._update_gift_contributed_count()
    app._save_gift_roster_to_db(silent=True)
    assert app.var_gift_total_amount.get() == "2,000"
    assert app._gift_money()["gift_collected"] == 2000.0
    assert _row(app, "EV1")["TotalIncome"] == "2,000"


def test_ticking_everyone_keeps_amounts_loaded_from_a_file(app, monolith, monkeypatch):
    """A partly ticked imported list: the header tick used to replace every
    imported amount with the expected budget."""
    _start_event(app)
    app.var_gift_budget.set("1,000")
    _select(app, app.tab_gift)
    monkeypatch.setattr(monolith.filedialog, "askopenfilename", lambda **_: "gifts.xlsx")
    monkeypatch.setattr(monolith, "read_gift_contribution_rows", lambda _p: [
        {"name": "Alice Example", "email": "alice@example.com", "checked": True, "amount": 2500.0},
        {"name": "Carol Example", "email": "carol@example.com", "checked": False, "amount": 1800.0},
    ])
    app._load_gift_list_from_file()
    app._toggle_all_gift_column("check")
    roster = app._gift_roster
    assert (roster["alice@example.com"]["amount"], roster["carol@example.com"]["amount"],
            roster["bob@example.com"]["amount"]) == (2500.0, 1800.0, 1000.0)


@pytest.mark.parametrize("which", ["round", "gift"])
def test_a_money_save_that_fails_is_reported_once(app, monolith, monkeypatch, which):
    """A later round's Paid box, like every Tab 5 / Tab 6 edit, saved with
    its failure swallowed: the screen kept the amount, the database lost it."""
    _attendance(app)
    key = _add_round(app, monolith, monkeypatch, "Karaoke")
    if which == "gift":
        _select(app, app.tab_gift)

    def locked(*_a, **_kw):
        raise sqlite3.OperationalError("database is locked")

    save = "save_attendance" if which == "round" else "save_gift_roster"
    monkeypatch.setattr(db, save, locked)
    before = len(app.dialogs)
    for typed in ("1,000", "1,500"):
        if which == "round":
            app._round_vars[key]["paid"].set(typed)
        else:
            app._commit_gift_amount_edit("alice@example.com", "amount", typed)
        app.update()
    reported = app.dialogs[before:]
    assert [d[1] for d in reported] == ["Not saved"]
    assert "NOT saved" in reported[0][2] and "database is locked" in reported[0][2]


def test_ticking_everyone_leaves_rows_already_ticked_alone(app):
    """Like Tab 5's headers: only rows the click changes get the budget."""
    _start_event(app)
    app.var_gift_budget.set("1,000")
    _select(app, app.tab_gift)
    app._gift_roster["alice@example.com"].update(checked=True, amount=1000.0, manual_amount=False)
    app.var_gift_budget.set("1,500")
    app._toggle_all_gift_column("check")
    assert (app._gift_roster["alice@example.com"]["amount"],
            app._gift_roster["bob@example.com"]["amount"]) == (1000.0, 1500.0)


def test_choosing_whom_to_send_the_gift_report_to_keeps_typed_figures(app):
    """Codex review of PR #12: ticking only "Send email" saved gift rows with
    nobody contributing, which counted as tracked money and recomputed the
    figures typed in the other copy to zero."""
    db.save_event_record({"EventID": "OLD1", "EventName": "Old", "TotalIncome": "9,999",
                          "Balance": "500"}, app.db_path)
    db.save_recipients("OLD1", [("Alice Example", "alice@example.com")], app.db_path)
    _load(app, "OLD1")
    _select(app, app.tab_gift)
    app._gift_roster["alice@example.com"]["send_email"] = True
    app._save_gift_roster_to_db(silent=True)
    assert (_row(app, "OLD1")["TotalIncome"], _row(app, "OLD1")["Balance"]) == ("9,999", "500")

    app._gift_roster["alice@example.com"].update(checked=True, amount=1200.0)
    app._save_gift_roster_to_db(silent=True)
    assert _row(app, "OLD1")["TotalIncome"] == "1,200"
    # Unticking the last contributor recomputes too, rather than keeping 1,200.
    app._gift_roster["alice@example.com"].update(checked=False, amount=0.0)
    app._save_gift_roster_to_db(silent=True)
    assert _row(app, "OLD1")["TotalIncome"] == "0"


def test_a_later_rounds_amount_cell_refuses_two_numbers(app, monolith, monkeypatch):
    """Codex review of PR #12: "1 000" in a later round's cell was saved as 1."""
    _attendance(app)
    key = _add_round(app, monolith, monkeypatch, "Karaoke")
    app._commit_attendance_edit("alice@example.com", f"extra_{key}", "2,000")
    app._commit_attendance_edit("alice@example.com", f"extra_{key}", "1 000")
    assert app._attendance_roster["alice@example.com"]["extra_amounts"][key] == 2000.0
    assert app.dialogs[-1][1] == "Amount not changed"


def test_clearing_the_only_tracked_money_recomputes_history(app):
    """Codex review of PR #12: clearing the gift price - the event's only
    tracked money - left its expense and balance in History, because the
    event no longer counted as tracked."""
    db.save_event_record({"EventID": "OLD1", "EventName": "Old"}, app.db_path)
    _load(app, "OLD1")
    _select(app, app.tab_gift)
    app.var_gift_item_price.set("3,000")
    app.update()
    assert (_row(app, "OLD1")["TotalExpense"], _row(app, "OLD1")["Balance"]) == ("3,000", "-3,000")
    app.var_gift_item_price.set("")
    app.update()
    assert (_row(app, "OLD1")["TotalExpense"], _row(app, "OLD1")["Balance"]) == ("0", "0")


def test_a_gift_item_that_could_not_be_read_is_never_saved_over(app, monolith, monkeypatch):
    """Codex review of PR #12: a failed read showed empty gift boxes still
    owned by the event, and the next edit saved all four blanks over the
    stored item."""
    db.save_event_record({"EventID": "EV9", "EventName": "Nine", "GiftItemName": "Speaker",
                          "GiftItemPrice": "3,570"}, app.db_path)
    real = app._history_record

    def unreadable(event_id):
        raise sqlite3.OperationalError("database is locked")

    monkeypatch.setattr(app, "_history_record", unreadable)
    app._adopt_gift_item("EV9")
    assert app.dialogs[-1][1] == "Gift item not loaded"
    monkeypatch.setattr(app, "_history_record", real)
    app.var_gift_item_link.set("https://shop.example.com/item")
    app.update()
    row = _row(app, "EV9")
    assert (row["GiftItemName"], row["GiftItemPrice"], row["GiftItemLink"]) == ("Speaker", "3,570", "")

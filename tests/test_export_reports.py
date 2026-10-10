"""The Excel reports in rsvp/export/reports.py.

They are attached to emails sent to colleagues, so a wrong figure is a wrong
figure in someone's inbox. The first test pins the bug that moved them here:
the attendance report re-read "Remaining amount" from the label on screen
with parse_amount_from_text, which drops the sign, so an overpaid event was
reported as money left over.
"""
from __future__ import annotations

from rsvp.domain import payment_rounds
from rsvp.export import reports


def _rows(wb):
    return list(wb.active.iter_rows(values_only=True))


def _summary(wb, label):
    return next(r for r in _rows(wb) if r[1] == label)


ATTENDANCE = {
    "alice@example.com": {"name": "Alice Example", "vote": "Yes", "actual_attend": "Yes",
                          "free": False, "amount": 3000.0},
    "bob@example.com": {"name": "Bob Example", "vote": "Maybe", "actual_attend": "No",
                        "free": False, "amount": 0.0},
    "carol@example.com": {"name": "Carol Example", "vote": "Yes", "actual_attend": "Yes",
                          "free": True, "amount": 0.0},
}


def _figures(roster, paid, rounds=(), label="Round 1"):
    return payment_rounds(roster.values(), label, paid, list(rounds))


class TestAttendanceWorkbook:
    def test_an_overpaid_event_keeps_its_negative_remaining(self):
        wb = reports.attendance_workbook(ATTENDANCE, _figures(ATTENDANCE, "5000"))
        assert _summary(wb, "Round 1 — Collect amount:")[6] == 3000.0
        assert _summary(wb, "Round 1 — Amount paid:")[6] == 5000.0
        assert _summary(wb, "Round 1 — Remaining amount:")[6] == -2000.0
        assert _summary(wb, "Total remaining amount (all rounds):")[6] == -2000.0

    def test_totals(self):
        wb = reports.attendance_workbook(ATTENDANCE, _figures(ATTENDANCE, "1000"))
        assert _summary(wb, "Total actual attend:")[4] == 2
        assert _summary(wb, "Round 1 — Remaining amount:")[6] == 2000.0
        header, *people = _rows(wb)[:4]
        assert header == ("No.", "Name", "Email", "Vote", "Round 1 Attend", "Free", "Round 1")
        assert people[2][5] == "Yes"          # Carol is Free
        assert people[1][6] is None           # a zero amount is left blank

    def test_every_round_has_its_columns_and_its_breakdown(self):
        roster = {email: dict(info) for email, info in ATTENDANCE.items()}
        roster["alice@example.com"].update(extra_attends={"round_2": "Yes"},
                                           extra_amounts={"round_2": 2000.0})
        rounds = [{"key": "round_2", "label": "Karaoke", "amount_paid": "2,500"}]
        wb = reports.attendance_workbook(roster, _figures(roster, "1000", rounds, "Dinner"))
        header, alice = _rows(wb)[:2]
        assert header[4:] == ("Dinner Attend", "Free", "Dinner", "Karaoke Attend", "Karaoke")
        assert alice[7:] == ("Yes", 2000.0)
        assert _summary(wb, "Karaoke — Attendees:")[4] == 1
        assert _summary(wb, "Karaoke — Remaining amount:")[6] == -500.0
        assert _summary(wb, "Total collected amount (all rounds):")[6] == 5000.0
        assert _summary(wb, "Total remaining amount (all rounds):")[6] == 1500.0


GIFT = {
    "alice@example.com": {"name": "Alice Example", "checked": True, "amount": 3000.0,
                          "send_email": True},
    "bob@example.com": {"name": "Bob Example", "checked": False, "amount": 0.0,
                        "send_email": True},
    "carol@example.com": {"name": "Carol Example", "checked": True, "amount": 2500.0,
                          "send_email": False},
}


class TestGiftWorkbooks:
    def test_the_full_list_reads_back_without_its_total_row(self, tmp_path):
        path = tmp_path / "Gift_Contribution_List_EV1.xlsx"
        wb = reports.gift_contribution_workbook(GIFT)
        assert _summary(wb, "TOTAL COLLECTED:")[3] == 5500.0
        wb.save(path)
        rows = reports.read_gift_contribution_rows(path)
        assert [(r["email"], r["checked"], r["amount"]) for r in rows] == [
            ("alice@example.com", True, 3000.0),
            ("bob@example.com", False, 0.0),
            ("carol@example.com", True, 2500.0),
        ]

    def test_the_report_lists_only_contributors_renumbered(self):
        wb = reports.gift_report_workbook(GIFT)
        header, alice, carol = _rows(wb)[:3]
        assert header == ("No.", "Name", "Email", "Amount")
        assert (alice[0], alice[2], carol[0], carol[2]) == (1, "alice@example.com", 2, "carol@example.com")
        assert "bob@example.com" not in [r[2] for r in _rows(wb)]
        assert _summary(wb, "TOTAL COLLECTED:")[3] == 5500.0

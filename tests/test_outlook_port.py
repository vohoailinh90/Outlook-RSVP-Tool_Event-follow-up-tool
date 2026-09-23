"""Phase 3: the Outlook seam.

Two things are proven here. First, that the port, the real COM module and
the test fake agree on every signature, so a test against the fake says
something about the real call. Second, that the one send path which can
call Outlook's Send() - the Tab 3 invite, whose auto-send comes from a
checkbox - hands Outlook exactly what the user chose.
"""
from __future__ import annotations

import ast
from datetime import datetime
from pathlib import Path

import pytest

from rsvp.services.invite import InviteRequest, send_invite
from tests.fake_outlook import FakeOutlook

ROOT = Path(__file__).resolve().parents[1]
NOON = datetime(2026, 9, 24, 12, 0)


def _signatures(path: Path, cls: str | None) -> dict[str, str]:
    """Public function signatures, parsed rather than imported: importing
    outlook_com needs pywin32, and this check must run anywhere."""
    tree = ast.parse(path.read_text(encoding="utf-8"))
    if cls is not None:
        body = next(n for n in tree.body
                    if isinstance(n, ast.ClassDef) and n.name == cls).body
    else:
        body = tree.body
    sigs = {}
    for node in body:
        if isinstance(node, ast.FunctionDef) and not node.name.startswith("_"):
            args = node.args
            if cls is not None:           # drop self
                args = ast.arguments(
                    posonlyargs=args.posonlyargs, args=args.args[1:],
                    vararg=args.vararg, kwonlyargs=args.kwonlyargs,
                    kw_defaults=args.kw_defaults, kwarg=args.kwarg,
                    defaults=args.defaults)
            sigs[node.name] = ast.unparse(args)
    return sigs


class TestSignaturesAgree:
    def test_port_matches_outlook_com(self):
        port = _signatures(ROOT / "rsvp" / "ports" / "outlook.py", "OutlookPort")
        real = _signatures(ROOT / "outlook_com.py", None)
        assert port == real

    def test_fake_matches_port(self):
        port = _signatures(ROOT / "rsvp" / "ports" / "outlook.py", "OutlookPort")
        fake = _signatures(ROOT / "tests" / "fake_outlook.py", "FakeOutlook")
        assert fake == port


def _request(**overrides) -> InviteRequest:
    base = dict(
        recipients=[("Example Person", "person@example.com"),
                    ("Other Person", "other@example.com")],
        subject="[Confirm-EV1] Party", body="Body text",
        voting_options="Yes;No;Maybe", auto_send=False,
        send_to_override=None, mode="invite",
        record={"EventID": "EV1", "EventName": "Party"})
    base.update(overrides)
    return InviteRequest(**base)


def _send(request, outlook=None, save=None):
    outlook = outlook or FakeOutlook()
    saved: list[dict] = []
    result = send_invite(outlook, save or saved.append, request,
                         now=lambda: NOON)
    return outlook, saved, result


class TestSendInvite:
    @pytest.mark.parametrize("auto_send", [False, True])
    def test_outlook_gets_exactly_the_auto_send_the_user_chose(self, auto_send):
        """auto_send=True is Send(): irreversible. It must never be set
        unless the user ticked it, and never dropped when they did."""
        outlook, _, result = _send(_request(auto_send=auto_send))
        [(name, call)] = outlook.calls
        assert name == "send_voting_invite"
        assert call["auto_send"] is auto_send
        assert ("SENT" if auto_send else "OPENED for review") in result.message

    def test_invite_goes_to_each_recipient_with_voting_buttons(self):
        outlook, _, result = _send(_request())
        call = outlook.calls[0][1]
        assert call["recipients"] == _request().recipients
        assert call["send_to_override"] is None
        assert call["use_voting_buttons"] is True
        assert "individually for 2 people" in result.message

    def test_group_override_is_passed_through(self):
        outlook, _, result = _send(_request(send_to_override="team@example.com"))
        assert outlook.calls[0][1]["send_to_override"] == "team@example.com"
        assert "to group address: team@example.com" in result.message

    def test_gift_notice_has_no_voting_buttons(self):
        outlook, _, _ = _send(_request(mode="gift"))
        assert outlook.calls[0][1]["use_voting_buttons"] is False

    @pytest.mark.parametrize("mode, column", [
        ("invite", "SentDate"), ("update", "UpdateInviteDate"), ("gift", None)])
    def test_history_records_the_send_time_in_one_column_by_mode(
            self, mode, column):
        _, saved, result = _send(_request(mode=mode))
        [record] = saved
        assert record["EventID"] == "EV1"
        written = {"SentDate", "UpdateInviteDate"} & record.keys()
        assert written == ({column} if column else set())
        if column:
            assert record[column] == "2026-09-24 12:00"
        assert result.send_time == NOON and result.history_written

    def test_request_record_is_not_mutated(self):
        request = _request()
        _send(request)
        assert "SentDate" not in request.record

    def test_outlook_failure_records_nothing(self):
        saved: list[dict] = []
        with pytest.raises(RuntimeError):
            send_invite(FakeOutlook(fail=RuntimeError("Outlook closed")),
                        saved.append, _request(), now=lambda: NOON)
        assert saved == []

    def test_history_failure_still_reports_the_send(self):
        def broken(record):
            raise OSError("database is locked")
        outlook, _, result = _send(_request(), save=broken)
        assert len(outlook.calls) == 1
        assert result.history_written is False
        assert "OPENED for review" in result.message

    def test_unknown_mode_is_refused_before_outlook_is_called(self):
        outlook = FakeOutlook()
        with pytest.raises(ValueError):
            send_invite(outlook, lambda r: None, _request(mode="invte"))
        assert outlook.calls == []

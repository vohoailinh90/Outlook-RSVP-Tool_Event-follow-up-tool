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
        real = _signatures(ROOT / "rsvp" / "adapters" / "outlook_com.py", None)
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


def test_the_vote_illustration_resolves_to_a_real_file():
    """outlook_com attaches how_to_vote.png only `if os.path.exists(...)`, so a
    wrong APP_ROOT drops the image from every invite without an error. The
    module needs pywin32, so evaluate its APP_ROOT expression from source."""
    import os
    adapter = ROOT / "rsvp" / "adapters" / "outlook_com.py"
    tree = ast.parse(adapter.read_text(encoding="utf-8"))
    expr = next(node.value for node in tree.body if isinstance(node, ast.Assign)
                and any(getattr(t, "id", None) == "APP_ROOT" for t in node.targets))
    app_root = eval(compile(ast.Expression(expr), str(adapter), "eval"),
                    {"os": os, "__file__": str(adapter)})
    assert os.path.isfile(os.path.join(app_root, "how_to_vote.png"))


# ── Distribution-list expansion, driven with stand-in AddressEntry objects ──
# The adapter imports pywin32 at module level, so this runs where the app
# runs (the Windows CI) and skips elsewhere. Nothing here touches Outlook:
# the entries below only answer the properties the expansion reads.

class _Members:
    def __init__(self, items):
        self._items, self.Count = items, len(items)

    def Item(self, i):
        return self._items[i - 1]


class _Entry:
    """A person (an SMTP address) or a personal contact group (members)."""

    def __init__(self, name, *, smtp="", members=None, entry_id=""):
        self.Name, self.Address, self.ID = name, smtp, entry_id
        is_group = members is not None
        self.DisplayType = 5 if is_group else 0              # olPrivateDistList
        self.AddressEntryUserType = 11 if is_group else 0    # olOutlookDistributionList...
        self.Members = _Members(members) if is_group else None

    def GetExchangeDistributionList(self):
        return None

    def GetExchangeUser(self):
        return None


class _Namespace:
    def GetAddressEntryFromID(self, entry_id):
        raise LookupError("not in this address book")

    def CreateRecipient(self, key):
        raise LookupError("not in this address book")


def test_two_groups_with_the_same_name_are_both_expanded():
    """Codex review of PR #12: two personal groups without an SMTP address
    but with the same display name shared one "already expanded" key, so
    the second group's people were dropped without being reported."""
    outlook_com = pytest.importorskip("rsvp.adapters.outlook_com")
    team = _Entry("Team", members=[
        _Entry("Friends", entry_id="ID-1", members=[_Entry("Person A", smtp="a@example.com")]),
        _Entry("Friends", entry_id="ID-2", members=[_Entry("Person B", smtp="b@example.com")]),
    ], entry_id="ID-0")
    failed, diag = [], []
    people = outlook_com._expand_dl_addr_entry(_Namespace(), team, set(), set(), 10, failed, diag)
    assert sorted(email for _name, email in people) == ["a@example.com", "b@example.com"]
    assert failed == []


def test_a_group_inside_itself_is_expanded_once():
    """The key still stops a loop: a group listed inside itself."""
    outlook_com = pytest.importorskip("rsvp.adapters.outlook_com")
    loop = _Entry("Loop", entry_id="ID-L", members=[_Entry("Person A", smtp="a@example.com")])
    loop.Members = _Members([_Entry("Person A", smtp="a@example.com"), loop])
    people = outlook_com._expand_dl_addr_entry(_Namespace(), loop, set(), set(), 10, [], [])
    assert people == [("Person A", "a@example.com")]


def test_a_member_without_an_email_address_is_reported():
    """Codex review of PR #12: a member Outlook gives no SMTP address for was
    skipped without a word, so the group was replaced by the others and that
    person dropped from every later invitation."""
    outlook_com = pytest.importorskip("rsvp.adapters.outlook_com")
    team = _Entry("Team", entry_id="ID-T", members=[
        _Entry("Person A", smtp="a@example.com"), _Entry("Person C")])
    failed = []
    people = outlook_com._expand_dl_addr_entry(_Namespace(), team, set(), set(), 10, failed, [])
    assert people == [("Person A", "a@example.com")]
    assert failed == ["Team: Person C (no email address)"]

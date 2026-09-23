"""Sending the Tab 3 email: the RSVP invite, an update invite, or a gift
contribution notice.

This is the one path in the app whose auto-send choice comes straight from a
checkbox, so it is the one path that can call Outlook's Send() and put a
message beyond recall. It was the first extracted from the UI for that
reason: a test can now check exactly what reaches Outlook.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Callable

from rsvp.ports import OutlookPort

MODES = ("invite", "update", "gift")


@dataclass(frozen=True)
class InviteRequest:
    recipients: list                 # [(name, email)] - the roster
    subject: str
    body: str
    voting_options: str
    auto_send: bool                  # True sends at once; False opens a draft
    send_to_override: str | None     # a group address to send to instead
    mode: str                        # one of MODES
    record: dict = field(default_factory=dict)   # Tab 1 fields for History


@dataclass(frozen=True)
class InviteResult:
    send_time: datetime
    history_written: bool
    message: str                     # the confirmation shown to the user


def send_invite(outlook: OutlookPort,
                save_record: Callable[[dict], None],
                request: InviteRequest,
                now: Callable[[], datetime] = datetime.now) -> InviteResult:
    """Hand the email to Outlook, then record the send in History.

    An Outlook failure propagates: nothing was sent and nothing is recorded.
    A History failure does not: the email has already been handed to
    Outlook, so the send is reported and the History write is best-effort.
    """
    if request.mode not in MODES:
        raise ValueError(f"unknown mode {request.mode!r}")
    is_gift = request.mode == "gift"
    outlook.send_voting_invite(
        request.recipients, request.subject, request.body,
        voting_options=request.voting_options,
        auto_send=request.auto_send,
        send_to_override=request.send_to_override,
        # A gift contribution notice is an announcement, not a vote: no
        # voting buttons and no voting illustration.
        use_voting_buttons=not is_gift,
    )
    send_time = now()

    # Only one of the two send-time columns is written, by mode; the other is
    # left out so an UPDATE of an existing row keeps its old value. A gift
    # notice is not an RSVP send and touches neither.
    record = dict(request.record)
    stamp = send_time.strftime("%Y-%m-%d %H:%M")
    if request.mode == "update":
        record["UpdateInviteDate"] = stamp
    elif request.mode == "invite":
        record["SentDate"] = stamp
    try:
        save_record(record)
        history_written = True
    except Exception:
        history_written = False

    return InviteResult(send_time, history_written, _confirmation(request))


def _confirmation(request: InviteRequest) -> str:
    mode_label = {"gift": "Gift Contribution Notice",
                  "update": "UPDATE invite",
                  "invite": "invite"}[request.mode]
    history_note = {
        "gift": "🗂 History updated with the latest Tab 1 details for this EventID.",
        "update": "🗂 History updated with the update-invite time for this EventID.",
        "invite": "🗂 History updated with the send time for this EventID.",
    }[request.mode]
    return (
        f"{mode_label} email {'SENT' if request.auto_send else 'OPENED for review'} "
        + (f"to group address: {request.send_to_override}"
           if request.send_to_override
           else f"individually for {len(request.recipients)} people")
        + ".\n\n"
        + ("" if request.auto_send else "Check it in Outlook, then click Send.")
        + f"\n\n{history_note}"
    )

"""The operations the application needs from Outlook.

`outlook_com.py` is the only real implementation, and a module satisfies a
Protocol structurally, so it is passed in as-is; there is no adapter class
that would only restate every signature a third time. Tests pass a fake.

Every signature here mirrors `outlook_com.py` exactly, keyword defaults
included. tests/test_outlook_port.py compares the three (this Protocol,
outlook_com.py and the test fake) so none of them can drift.

Threading: outlook_com calls pythoncom.CoInitialize()/CoUninitialize() inside
each function, on whichever thread calls it. Calling through this Protocol
does not change the calling thread, so the COM apartment rules are the same
as calling outlook_com directly.

Sending is irreversible: with auto_send=False a message is only opened for
review (Display()); with auto_send=True it is sent (Send()) the moment the
call runs.
"""
from __future__ import annotations

from typing import Any, Protocol


class OutlookPort(Protocol):
    def send_voting_invite(self, recipients, subject, body,
                           voting_options="Yes;No;Maybe", auto_send=False,
                           send_to_override=None,
                           use_voting_buttons=True) -> Any: ...

    def list_folder_paths(self, max_depth=3) -> list[str]: ...

    def scan_voting_responses(self, event_id, folder_paths=None,
                              scan_all=False) -> Any: ...

    def send_calendar_invite(self, attendees, subject, location, start_dt,
                             end_dt, body="", attach_event_id=None,
                             attach_hint_datetime=None,
                             include_update=True) -> Any: ...

    def send_reminder_email(self, pending, subject, body,
                            voting_options="Yes;No;Maybe", auto_send=False,
                            attach_event_id=None,
                            attach_hint_datetime=None) -> Any: ...

    def send_gift_reminder_email(self, pending, subject, body,
                                 auto_send=False,
                                 attach_event_id=None) -> Any: ...

    def send_gift_report_email(self, recipients, subject, body,
                               excel_path=None, auto_send=False) -> Any: ...

    def send_thankyou_email(self, recipients, subject, body, excel_path=None,
                            event_name=None, auto_send=False) -> Any: ...

    def expand_group_members(self, email_or_name, max_depth=6) -> Any: ...

"""An OutlookPort that records every call and never touches COM.

Its signatures must match rsvp/ports/outlook.py and outlook_com.py exactly;
tests/test_outlook_port.py fails if any of the three drift apart.
"""
from __future__ import annotations


class FakeOutlook:
    def __init__(self, fail: Exception | None = None):
        self.calls: list[tuple[str, dict]] = []
        self.fail = fail

    def _record(self, name: str, **kwargs):
        self.calls.append((name, kwargs))
        if self.fail is not None:
            raise self.fail

    def send_voting_invite(self, recipients, subject, body,
                           voting_options="Yes;No;Maybe", auto_send=False,
                           send_to_override=None, use_voting_buttons=True):
        self._record("send_voting_invite", recipients=recipients,
                     subject=subject, body=body,
                     voting_options=voting_options, auto_send=auto_send,
                     send_to_override=send_to_override,
                     use_voting_buttons=use_voting_buttons)
        return object()

    def list_folder_paths(self, max_depth=3):
        self._record("list_folder_paths", max_depth=max_depth)
        return []

    def scan_voting_responses(self, event_id, folder_paths=None,
                              scan_all=False):
        self._record("scan_voting_responses", event_id=event_id,
                     folder_paths=folder_paths, scan_all=scan_all)
        return {}, 0

    def send_calendar_invite(self, attendees, subject, location, start_dt,
                             end_dt, body="", attach_event_id=None,
                             attach_hint_datetime=None, include_update=True):
        self._record("send_calendar_invite", attendees=attendees,
                     subject=subject, location=location, start_dt=start_dt,
                     end_dt=end_dt, body=body,
                     attach_event_id=attach_event_id,
                     attach_hint_datetime=attach_hint_datetime,
                     include_update=include_update)
        return object(), 0

    def send_reminder_email(self, pending, subject, body,
                            voting_options="Yes;No;Maybe", auto_send=False,
                            attach_event_id=None, attach_hint_datetime=None):
        self._record("send_reminder_email", pending=pending, subject=subject,
                     body=body, voting_options=voting_options,
                     auto_send=auto_send, attach_event_id=attach_event_id,
                     attach_hint_datetime=attach_hint_datetime)
        return object(), False

    def send_gift_reminder_email(self, pending, subject, body,
                                 auto_send=False, attach_event_id=None):
        self._record("send_gift_reminder_email", pending=pending,
                     subject=subject, body=body, auto_send=auto_send,
                     attach_event_id=attach_event_id)
        return object(), False

    def send_gift_report_email(self, recipients, subject, body,
                               excel_path=None, auto_send=False):
        self._record("send_gift_report_email", recipients=recipients,
                     subject=subject, body=body, excel_path=excel_path,
                     auto_send=auto_send)
        return object(), False

    def send_thankyou_email(self, recipients, subject, body, excel_path=None,
                            event_name=None, auto_send=False):
        self._record("send_thankyou_email", recipients=recipients,
                     subject=subject, body=body, excel_path=excel_path,
                     event_name=event_name, auto_send=auto_send)
        return object(), False

    def expand_group_members(self, email_or_name, max_depth=6):
        self._record("expand_group_members", email_or_name=email_or_name,
                     max_depth=max_depth)
        return None

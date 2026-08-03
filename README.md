# Outlook RSVP Tool

A desktop application (Python/Tkinter) that automates the entire internal event
management workflow: sending Yes/No/Maybe voting-button emails via Outlook, automatically
collecting responses, sending Calendar Invites, tracking attendance & payment, managing
gift contributions, automated reminders, and keeping a full event history — with support
for 3 languages (English / Japanese / Vietnamese / bilingual).

Runs on **Windows with Outlook Desktop** (using COM automation via `pywin32` — no API
key needed, no Azure app registration required).

---

## ⚠️ Requirements

- **Windows** (won't run on Mac/Linux — requires Outlook Desktop COM)
- **Outlook Desktop** installed and **signed in** (not Outlook Web/OWA)
- Python 3.8+

```bash
pip install pywin32 openpyxl --break-system-packages
```

---

## 1. Architecture & core files

All data lives in a **single SQLite file**: `rsvp_data.db` (auto-created next to the app
on first run). Excel is now only used for **manual export/import** when you need to
share a report — it's no longer where "live" data is stored, unlike the older version.

**3 files that must sit in the same folder:**

| File | Role |
|---|---|
| `rsvp_app.py` | Main UI (7 tabs) + all logic |
| `db.py` | SQLite storage layer (`rsvp_data.db`) — replaces the old `history.py`/Excel |
| `outlook_com.py` | Sending/reading emails, Calendar Invites via Outlook COM |

**Other supporting files:**

| File | Role |
|---|---|
| `how_to_vote.png` | Instructional image showing how to click the Vote buttons — auto-attached to every invite/reminder email |
| `send_scheduled_reminders.py` + `run_scheduled_reminders.bat` | Background script for Windows Task Scheduler to send reminders on a schedule |
| `rsvp_data.db` | The main database — **back this up regularly** |

> 💡 Old Excel files (`RSVP_History.xlsx`, etc.), if left over from a previous version,
> are **automatically migrated once** into `rsvp_data.db` the first time the app starts,
> so no old event data is lost.

---

## 2. Running the app

```bash
python rsvp_app.py
```

The window opens with 7 tabs. For each new event, work through them left to right.

---

## 3. Workflow (tab by tab)

### Tab 1 — Event Setup
Enter event details: Event ID (unique identifier used to match emails), event name,
date/time (Start/End Time), location, deadline, budget. Choose an **Event Mode**:
- `Event` — a normal event
- `Gift` — gift-contribution collection only (adds Organizer / Guest of Honor /
  Expected Gift Budget / Gift Contribution Deadline)
- `Event + Gift` — both

### Tab 2 — Recipients
Load the recipient list (Name + Email) from an Excel file, or enter manually.

### Tab 3 — Compose & Send
Compose and send email using one of 3 send modes:
- **Voting invite** — Yes/No/Maybe voting-button email
- **Send Gift Contribution Notice** — call for gift contributions
- Choose language: English / Japanese / Vietnamese / Bilingual (JP+EN)

Content can be hand-edited; supports bilingual translation via a copy-paste bridge with
Microsoft Copilot (paste a ready-made prompt, paste the translated result back into the
app).

### Tab 4 — Collect Responses
**Scan Inbox** to automatically read Yes/No/Maybe responses from recipients' reply
emails. Individual votes can be manually corrected (tick "Manual edit" to open a
dropdown) — manually edited votes are preserved across future Scan Inbox runs, unless a
newer reply email for that same person is found.

### Tab 5 — Attendance & Payment
Send Calendar Invites to Yes/Maybe recipients. Track actual attendance (Actual Attend)
and how much each person contributed. Enter "Amount paid" (the actual amount spent) to
auto-calculate "Remaining amount". After the event, send a **Thank You** email with the
Attendance & Payment Excel file and the Calendar Invite attached.

### Tab 6 — Gift Contribution
Track who has contributed gift money (✅/⬜ checkboxes), with name/email search. Choose
who receives the report email (the "Send email" column, independent of "Contributed").
Send a summary report email (no per-person list included) with a separate Excel file
containing only the people who contributed.

### Tab 7 — Event History
Review the full history of all created events, Yes/No/Maybe counts, and send status.

---

## 4. Automated reminders (Task Scheduler)

`send_scheduled_reminders.py` runs independently, reading `rsvp_data.db` to find people
who "haven't responded" and sending them individual reminder emails. Set it up to run
daily via `run_scheduled_reminders.bat` + Windows Task Scheduler.

---

## 5. How recipients respond

Recipients simply **click a Vote button** (Yes/No/Maybe) right inside the email — no
typing required, no need to keep the Subject line intact. The attached `how_to_vote.png`
illustrates the 3 steps.

---

## 6. Common issues

| Issue | Cause / Fix |
|---|---|
| `ModuleNotFoundError: win32com` | `pywin32` not installed — run `pip install pywin32` again |
| App can't open Outlook / sending fails | Outlook Desktop isn't open/signed in — open Outlook first, then retry |
| Scan Inbox doesn't pick up new responses | Check that the Event ID in Tab 1 matches the one used when the invite was sent; make sure reply emails are in the Inbox (not moved to another folder by a rule) |
| Calendar Invite not attaching to the Thank You email | The Calendar Invite's Subject in Outlook must match the Event Name exactly for the app to find it (best-effort search) |
| App errors related to `rsvp_data.db` on startup | Back up the old `.db` file, check it isn't locked/open by another program (SQLite lock) |

---

## 7. Current limitations

- Only reads email from the **Inbox** — if you have a rule that auto-moves email to
  another folder, you'll need to update the folder lookup in `outlook_com.py`
- Can't distinguish between two people sharing the same email address
- The old "Actual cost tracking" feature (formerly in Tab 4) has been fully replaced by
  the more detailed Tab 5 (Attendance & Payment) — the old DB columns are kept as-is
  (not deleted) so historical event data isn't lost, but nothing reads/writes them
  anymore

---

## 8. Possible future extensions

- An "Export to Excel" button on each tab for sharing reports outside the app
- A multi-event dashboard summary in Tab 7
- Support for multiple organizers sharing the same `rsvp_data.db`

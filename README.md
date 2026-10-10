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

**What must sit together in one folder:**

| File | Role |
|---|---|
| `rsvp_app.py` | Main UI (7 tabs) — the file you run |
| `rsvp/` | The app's package: `storage/db.py` (SQLite, `rsvp_data.db`), `storage/settings.py` (saved wording and Copilot prompt), `adapters/outlook_com.py` (Outlook COM: emails, Calendar Invites), `export/` (Excel reports and the one-time import of old Excel files), `domain/`, `i18n/`, `ports/`, `services/` |

**Other supporting files:**

| File | Role |
|---|---|
| `how_to_vote.png` | Instructional image showing how to click the Vote buttons — auto-attached to every invite/reminder email |
| `rsvp_data.db` | The main database — **back this up regularly**. Not in git: it holds real names and addresses, so it is gitignored and `scripts/check_no_pii.py` fails the build if it is committed |
| `scripts/check_*.py` | Guards that run in CI — see `CLAUDE.md` |
| `.claude/agents/` | Agent roster for Claude Code — see `CLAUDE.md` |
| `docs/agentic/ARCHITECTURE.md` | Target structure and the staged plan to reach it |

> 💡 Old Excel files (`RSVP_History.xlsx`, etc.), if left over from a previous version,
> are **automatically migrated once** into `rsvp_data.db` the first time the app starts,
> so no old event data is lost.

---

## 2. Running the app

```bash
python rsvp_app.py
```

The window opens with a sidebar listing the 7 steps of the workflow (the "tabs" below);
for each new event, work through them top to bottom. The current Event ID is shown at the
top right of every page. The look follows the Automation UI Kit (white cards, black primary
buttons, colour only for status); its tokens live in `rsvp/ui/`.

---

## 3. Workflow (tab by tab)

### Tab 1 — Event Setup
Enter event details: Event ID (unique identifier used to match emails), event name,
date/time (Start/End Time), location, deadline, budget. **💾 Save event details** writes them
to History (no email is sent); **🆕 Save & start a new event** saves and then clears every
page for the next event. Choose an **Event Mode**:
- `Event` — a normal event
- `Gift` — gift-contribution collection only (adds Organizer / Guest of Honor /
  Expected Gift Budget / Gift Contribution Deadline)
- `Event + Gift` — both

### Tab 2 — Recipients
**Import from Excel…** (Name in column A, Email in column B, on any row — title and header
rows are skipped, repeated addresses are dropped), or add people by hand. The list is saved
to the database after every change; **Export to Excel** writes a file in the same layout.

### Tab 3 — Compose & Send
Compose and send email using one of 3 send modes:
- **Voting invite** — Yes/No/Maybe voting-button email
- **Send Gift Contribution Notice** — call for gift contributions
- Choose language: English / Japanese / Vietnamese / Bilingual (JP+EN)

The preview is rebuilt from Tab 1 each time you open the tab, unless you edited it by hand.
Content can be hand-edited; supports bilingual translation via a copy-paste bridge with
Microsoft Copilot (paste a ready-made prompt, paste the translated result back into the
app).

**Bilingual (JP+EN):** the note box holds only your note, in any language; the fixed-part
box holds the event details and voting instructions in Japanese and English, either side of
a divider line. **📋 Copy note + prompt** sends just the note to Copilot, which answers with
a `[JA]` part and an `[EN]` part in one go. Paste that answer back and Save: the note box then
shows both languages, and each half of the email gets its own greeting and fixed part,
Japanese first and English second. An empty note needs no translation at all.

### Tab 4 — Collect Responses
**Scan Inbox** reads the Yes/No/Maybe replies from every folder of your mailbox. Individual
votes can be manually corrected (tick "Manual edit" to open a dropdown) — manually edited
votes are preserved across future Scan Inbox runs, unless a newer reply email for that same
person is found. Results and the Yes/No/Maybe counts in History are saved automatically.
The reminder (and Tab 5's Calendar Invite) is only sent once the event on Tab 1 has been
scanned, so it can never go to another event's list.

### Tab 5 — Attendance & Payment
Send Calendar Invites to Yes/Maybe recipients. After the event, tick who actually came
(click an Attend or Free cell; click a column header to tick everyone) and correct amounts
by double-clicking them. Ticking Attend fills that round's amount from the expected event
budget; Free exempts a person from every round.

**Payment rounds:** round 1 is the main event. **➕ Add round** adds a follow-up (a second
venue, another evening) with its own Attend and amount columns; rename a round with
**✎ Rename** (or by double-clicking its amount header), remove it with **✖ Remove**. Type
what was actually paid out in each round; the rounds card shows each round's attendees,
collected, paid out and remaining, and the tiles at the top the totals over all rounds.

The **Thank You** email carries a table of every round's money (sent as HTML so the
columns line up in Outlook), with the Attendance & Payment Excel file and the Calendar
Invite attached. **🔄 Update from table** rebuilds the text after you change the table;
text you did not edit is rebuilt automatically before sending.

### Tab 6 — Gift Contribution
Track who has contributed gift money (✅/⬜ checkboxes), with name/email search.
Double-click an Amount to type what someone actually gave; a typed amount is kept when
Contributed is ticked again. Choose who receives the report email (the "Send email"
column, independent of "Contributed").

**Gift item & money:** the gift's name, order link (**🌐 Open link**) and price give the
gift's remaining money. Tick **🔗 Add the party's money** to add Tab 5's totals for one
Event + Gift summary. The report email (HTML, so its table lines up and the link opens)
quotes these figures, with a separate Excel file listing only the contributors.

### Tab 7 — Event History
Review the full history of all created events, Yes/No/Maybe counts, and send status.
Double-click a cell to edit it; **Save changes** writes only the cells you edited.
Renaming an Event ID moves the whole event (recipients, votes, attendance, payment
rounds, gift list). Drag a column heading sideways to reorder the columns (remembered in
the database file); **↔ Fit columns** and **↺ Reset order** tidy up.

The money columns are calculated, not typed: **Actual Att. (main)** and **Cost/Person**
describe round 1; **Income**, **Expense** and **Balance** cover every round and the gift;
**Dept. Fund Left** is the running total of Balance from the first event down to that row.
An event with nothing tracked on Tab 5/6 keeps the figures it already has.

### Loading a past event
Type in the search box above the list to find an event by any field, or pick one field
(Location, Guest of Honor, Gift Item...) on its left.
**⬅ Load setup from selected event** (Tab 1) replaces everything the app holds for the
current event with what is saved for the selected one — nothing carries over. It asks first
when that would discard Copilot translations or email text you edited by hand. To use an
event as a template, load it and change the Event ID: the recipient list follows the new ID,
while votes, attendance, gift ticks and Amount paid stay with the original event.

---

## Moving over from the "Event-Invitation" copy

That copy of the app and this one use the same database layout. Copy its `rsvp_data.db`
next to this `rsvp_app.py` (keep the original as a backup); every event, payment round,
gift item and your History column order comes across. The other copy can still open the
file afterwards.

---

## 4. Automated reminders (Task Scheduler) — NOT IMPLEMENTED

> ⚠️ Earlier versions of this README described a `send_scheduled_reminders.py` script and
> a `run_scheduled_reminders.bat` launcher for Windows Task Scheduler. **Neither file
> exists in this repository.** The description is kept here only to say so, because the
> previous wording read as though the feature shipped.
>
> Reminders are sent manually today, from the "Collect Responses" and "Gift Contribution"
> tabs. Scheduled sending is an unbuilt feature, not a broken one.

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
| Scan Inbox doesn't pick up new responses | Check that the Event ID in Tab 1 matches the one used when the invite was sent. Outlook's search index can lag a minute behind new mail; if a scan finds nothing where votes were found before, the previous results are kept — scan again shortly |
| Calendar Invite not attaching to the Thank You email | The Calendar Invite's Subject in Outlook must match the Event Name exactly for the app to find it (best-effort search) |
| App errors related to `rsvp_data.db` on startup | Back up the old `.db` file, check it isn't locked/open by another program (SQLite lock) |

---

## 7. Current limitations

- Can't distinguish between two people sharing the same email address
- ReminderSent and ReportFile are hidden from Tab 7 (nothing here writes them) but kept in
  the database, and History's **Export to Excel** still writes them
- Amount paid, the first round's name and the gift item are stored on the event's History
  row: for an event not in History yet they are kept on screen and written when you save
  the event on Tab 1

---

## 8. Possible future extensions

- An "Export to Excel" button on each tab for sharing reports outside the app
- A multi-event dashboard summary in Tab 7
- Support for multiple organizers sharing the same `rsvp_data.db`

"""
rsvp_app.py — MAIN GUI for the Outlook RSVP Tool
=============================================================================
Runs on WINDOWS, requires Outlook desktop installed & signed in.

Install:
    pip install pywin32 openpyxl tkcalendar

Run:
    python rsvp_app.py

All on-screen text (labels, buttons, dialogs) is in ENGLISH, regardless of
which language the OUTGOING EMAIL is composed in — the email language is a
separate setting on the "Compose & Send" tab (English / Japanese /
Vietnamese / Bilingual EN+JP), since recipients may be Japanese and Indian
colleagues.

7 tabs:
    1. Event Setup        — event details incl. expected budget; can reload
                             a past event as a template
    2. Recipients          — load / edit the invite list directly
    3. Compose & Send       — Voting Buttons email (Yes/No/Maybe), choose
                             email language, translate the free-text note
                             via a Copilot copy/paste bridge
    4. Collect Responses    — scan Inbox, read vote results, export report
    5. Attendance & Payment — send a Calendar Invite to everyone who voted
                             Yes or Maybe, then track who actually showed
                             up and how much they paid
    6. Gift Contribution    — track who has contributed to a gift, with
                             per-person amount and running total
    7. Event History        — browse / reuse everything saved so far
"""
import os
import re
import tempfile
import threading
from datetime import datetime, timedelta

import tkinter as tk
import webbrowser
from tkinter import ttk, filedialog, messagebox, simpledialog
from tkinter import font as tkfont

try:
    from tkcalendar import DateEntry
    HAS_TKCALENDAR = True
except ImportError:
    HAS_TKCALENDAR = False

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from rsvp.adapters import outlook_com
from rsvp.export import legacy_excel, reports
from rsvp.export.reports import read_gift_contribution_rows
from rsvp.storage import db, settings
from rsvp.ports import OutlookPort
from rsvp.services.invite import InviteRequest, send_invite
from rsvp.ui import theme
from rsvp.ui.tokens import COLORS
from rsvp.ui.widgets import Banner, Card, PageHeader, Sidebar, WrapLabel, autohide, bordered, kpi_row

APP_TITLE = "Outlook RSVP Tool"

# The sidebar and page header: (attribute of the page, sidebar label, page title, description).
PAGES = [
    ("tab_config", "Event Setup", "Event setup",
     "The details every email is built from. Save the event here, or reload a past one."),
    ("tab_recipients", "Recipients", "Recipients",
     "Who is invited. The list is saved for the current Event ID after every change."),
    ("tab_compose", "Compose & Send", "Compose & send",
     "Write the invitation, translate it if needed, and open it in Outlook."),
    ("tab_collect", "Collect Responses", "Collect responses",
     "Read the Yes / No / Maybe votes from your mailbox, and remind whoever has not answered."),
    ("tab_calendar", "Attendance & Payment", "Attendance & payment",
     "Calendar invite for Yes / Maybe, who actually came, what they paid, and the thank-you email."),
    ("tab_gift", "Gift Contribution", "Gift contribution",
     "Who has contributed to the gift, reminders, and a report for the people you choose."),
    ("tab_history", "Event History", "Event history",
     "Every saved event. Double-click a cell to edit it, then save."),
]

# Tab 7 (Event History): (column, header, width) - the default column order;
# the user can drag headings to reorder them (saved in the database, see
# _load_history_column_order()). Same columns and headers as the other
# lineage of this app. Not shown: ReminderSent (only ever set for a
# scheduled-reminder script that does not exist here), ReportFile (only set
# by a report export that had no button) and AmountPaid (edited on Tab 5).
# Their stored values are kept, and "Export to Excel" still writes them.
HISTORY_TABLE_COLUMNS = [
    ("EventID", "Event ID", 110), ("EventName", "Event Name", 170),
    ("EventDate", "Date", 96), ("Deadline", "Deadline", 96),
    ("Location", "Location", 110), ("Budget", "Budget", 90),
    ("EmailLanguage", "Language", 140), ("OrganizerNote", "Organizer Note", 220),
    ("RecipientFile", "Recipient File", 220), ("SentDate", "Sent Date", 120),
    ("UpdateInviteDate", "Update Invite Date", 130),
    ("TotalInvited", "Invited", 65), ("Yes", "Yes", 45),
    ("ActualAttendees", "Actual Att. (main)", 120),
    ("No", "No", 45), ("Maybe", "Maybe", 55), ("NoResponse", "No Resp.", 70),
    ("CalendarSent", "Calendar Invite", 140),
    ("LastReminderSentDate", "Last Reminder Sent", 140),
    ("EventMode", "Event Mode", 90), ("Organizer", "Organizer", 140),
    ("GuestOfHonor", "Guest of Honor", 140), ("GiftBudget", "Gift Budget", 100),
    ("GiftDeadline", "Gift Deadline", 90), ("StartTime", "Start", 65), ("EndTime", "End", 65),
    ("GiftItemName", "Gift Item", 140), ("GiftItemLink", "Gift Order Link", 180),
    ("GiftItemPrice", "Gift Price", 90),
    # The money columns, all on the right so they read together.
    ("CostPerPerson", "Cost/Person", 90), ("TotalIncome", "Income (all)", 100),
    ("TotalExpense", "Expense (all)", 100), ("Balance", "Balance", 90),
    ("DeptFundRemaining", "Dept. Fund Left", 110),
]
HISTORY_TABLE_KEYS = tuple(c for c, _h, _w in HISTORY_TABLE_COLUMNS)
HISTORY_TABLE_HEADERS = dict((c, h) for c, h, _w in HISTORY_TABLE_COLUMNS)
# Written by the app, never typed on Tab 7: the money columns follow Tab 5
# and Tab 6 (_sync_event_money()), Dept. Fund Left is a running total of
# Balance, and the gift item is edited on Tab 6. A Tab 7 edit would be
# overwritten by the next change there, or overwrite it.
HISTORY_READ_ONLY = frozenset({
    "ActualAttendees", "CostPerPerson", "TotalIncome", "TotalExpense", "Balance",
    "DeptFundRemaining", "GiftItemName", "GiftItemLink", "GiftItemPrice",
})
HISTORY_COLUMN_ORDER_KEY = "history_column_order"  # app_settings key, shared with the other lineage
HISTORY_SEARCH_ANY = "(Any field)"

# The first payment round's name on Tab 5 when the user has not named it.
ROUND1_DEFAULT_LABEL = "Round 1"
TRUTHY = ("yes", "true", "1", "✅", "x")


# ══════════════════════════════════════════════════════════════════════════
# Language, message and paste-cleanup helpers
#
# Moved to rsvp/i18n/ (phase 1, docs/agentic/ARCHITECTURE.md). They are
# re-exported here under their original names so every call site in this
# file keeps working unchanged - the extraction moved the code, it did not
# rewrite the 108 methods that call it.
#
# EVERY public name is re-exported, including the label tables nothing in
# this file currently reads. Re-exporting only what is referenced today would
# quietly shrink this module's namespace - a behavior change at the module
# API level, and one a behavior snapshot catches even though the application
# still runs. The extraction is supposed to move code, not narrow an API.
#
# New code should import from rsvp.i18n directly rather than relying on
# these names; that package imports without tkinter or Outlook, which is
# what makes it testable anywhere.
# ══════════════════════════════════════════════════════════════════════════
from rsvp.domain import (  # noqa: F401
    amount_for,
    contributed_total,
    count_actual_attendees,
    gift_figures,
    history_figures,
    is_yes,
    merge_expanded_roster,
    format_amount,
    parse_amount_from_text,
    parse_typed_amount,
    payment_rounds,
    remaining_amount,
    round_totals,
    running_fund,
    sum_contributions,
    unclear_typed_amount,
)
from rsvp.i18n import (  # noqa: F401
    BILINGUAL_SEPARATOR,
    CALENDAR_LABELS,
    DEFAULT_PROMPT_BILINGUAL,
    DEFAULT_PROMPT_SINGLE,
    EMOJI_PATTERN,
    GIFT_LABELS,
    GIFT_REPORT_LABELS,
    GREETING,
    LANG_LABELS,
    LANG_LABEL_TO_CODE,
    NOT_TRANSLATED_FLAG,
    REMINDER_LABELS,
    THANKYOU_LABELS,
    TRANSLATE_TARGETS,
    UPDATE_NOTICE,
    build_calendar_body,
    build_editable_block,
    build_fixed_block,
    build_gift_fixed_block,
    build_gift_reminder_body,
    build_gift_reminder_subject,
    build_gift_report_body,
    build_gift_report_subject,
    build_gift_subject,
    build_greeting,
    build_reminder_body,
    build_reminder_subject,
    build_subject,
    build_thankyou_body,
    build_thankyou_subject,
    build_update_notice,
    build_bilingual_body,
    cleanup_pasted_translation,
    dedupe_pasted_translation,
    detect_possible_duplicate_paste,
    join_bilingual,
    parse_bilingual_reply,
    split_bilingual,
    text_body_to_html,
)

# Bilingual is not a target to pick: with Email language = Bilingual the note,
# in any language, always goes to Copilot for Japanese and English at once.
BILINGUAL_TARGET = "Bilingual (Japanese + English)"
SINGLE_TARGETS = [t for t in TRANSLATE_TARGETS if t != BILINGUAL_TARGET]


# ══════════════════════════════════════════════════════════════════════════
# Date-picker helper widget
# ══════════════════════════════════════════════════════════════════════════
# The drop-down calendar in the kit's colours instead of tkcalendar's blue.
CALENDAR_COLORS = dict(
    background=COLORS["primary"], foreground=COLORS["primary_foreground"],
    headersbackground=COLORS["muted"], headersforeground=COLORS["muted_foreground"],
    normalbackground=COLORS["card"], normalforeground=COLORS["foreground"],
    weekendbackground=COLORS["card"], weekendforeground=COLORS["foreground"],
    othermonthbackground=COLORS["muted"], othermonthforeground=COLORS["muted_foreground"],
    othermonthwebackground=COLORS["muted"], othermonthweforeground=COLORS["muted_foreground"],
    selectbackground=COLORS["primary"], selectforeground=COLORS["primary_foreground"],
    bordercolor=COLORS["border"], borderwidth=0,
)


def make_date_picker(parent, initial=None):
    if HAS_TKCALENDAR:
        w = DateEntry(parent, date_pattern="dd/mm/yyyy", width=14, **CALENDAR_COLORS)
        if initial:
            try:
                w.set_date(initial)
            except Exception:
                pass
        return w
    else:
        w = ttk.Entry(parent, width=16)
        w.insert(0, initial.strftime("%d/%m/%Y") if hasattr(initial, "strftime") else (initial or ""))
        return w


def get_date_str(widget):
    if HAS_TKCALENDAR and isinstance(widget, DateEntry):
        return widget.get_date().strftime("%d/%m/%Y")
    return widget.get().strip()


def get_date_obj(widget, default_hour=18, default_minute=0):
    if HAS_TKCALENDAR and isinstance(widget, DateEntry):
        d = widget.get_date()
        return datetime(d.year, d.month, d.day, default_hour, default_minute)
    try:
        d, m, y = widget.get().strip().split("/")
        return datetime(int(y), int(m), int(d), default_hour, default_minute)
    except Exception:
        return datetime.now()


def set_date_str(widget, date_str):
    """Nạp lại giá trị ngày (chuỗi 'dd/mm/yyyy', đọc từ History) vào 1 ô
    date-picker (DateEntry của tkcalendar, hoặc Entry thường nếu không có
    tkcalendar). Bỏ qua an toàn (không làm gì) nếu date_str rỗng hoặc không
    parse được — để không làm hỏng ô đang có sẵn."""
    if not date_str:
        return
    try:
        d, m, y = str(date_str).strip().split("/")
        d, m, y = int(d), int(m), int(y)
        if HAS_TKCALENDAR and isinstance(widget, DateEntry):
            widget.set_date(datetime(y, m, d))
        else:
            widget.delete(0, "end")
            widget.insert(0, f"{d:02d}/{m:02d}/{y}")
    except Exception:
        pass  # không parse được (định dạng lạ) — giữ nguyên giá trị đang có trên ô


# ══════════════════════════════════════════════════════════════════════════
# Scrollable page container — every page's real content goes inside `.body`,
# so long pages get a vertical scrollbar instead of being cut off on smaller
# screens / windows.
# ══════════════════════════════════════════════════════════════════════════
class ScrollableFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        # tk.Canvas does not follow the ttk theme, so it gets the page colour explicitly.
        canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0, background=COLORS["background"])
        vscroll = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        # Page margins of the kit's main area (p-6), minus what the header has.
        self.body = ttk.Frame(canvas, padding=(28, 20, 28, 28))

        self.body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas_window = canvas.create_window((0, 0), window=self.body, anchor="nw")

        # Each time the canvas is resized, the body is forced to exactly its
        # width, so cards fill the page and wrapping labels (WrapLabel) reflow.
        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
        canvas.bind("<Configure>", _on_canvas_configure)

        canvas.configure(yscrollcommand=vscroll.set)

        # The body is always exactly as wide as the canvas, so there is never
        # anything to scroll sideways.
        canvas.grid(row=0, column=0, sticky="nsew")
        vscroll.grid(row=0, column=1, sticky="ns")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _bind_wheel(_e):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _unbind_wheel(_e):
            canvas.unbind_all("<MouseWheel>")

        canvas.bind("<Enter>", _bind_wheel)
        canvas.bind("<Leave>", _unbind_wheel)
        self.canvas = canvas


def make_scrollable_treeview(parent, columns, height=10, show="headings", horizontal=False):
    """Creates a Treeview WITH its scrollbar(s), inside a bordered container.
    `horizontal` adds a bottom scrollbar, for tables wider than the page.
    Returns (container, tree). Grid/pack the returned CONTAINER — never the tree itself —
    since the tree's actual parent is the container, not the outer `parent` passed in."""
    container = bordered(parent)
    tree = ttk.Treeview(container, columns=columns, show=show, height=height)
    vscroll = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=autohide(vscroll))
    tree.grid(row=0, column=0, sticky="nsew")
    vscroll.grid(row=0, column=1, sticky="ns")
    if horizontal:
        hscroll = ttk.Scrollbar(container, orient="horizontal", command=tree.xview)
        tree.configure(xscrollcommand=autohide(hscroll))
        hscroll.grid(row=1, column=0, sticky="ew")
    container.rowconfigure(0, weight=1)
    container.columnconfigure(0, weight=1)
    return container, tree


def make_scrollable_text(parent, **text_kwargs):
    """Creates a Text widget WITH its scrollbar, both inside a bordered container
    whose border turns to the focus colour while the text has focus.
    Returns (container, text_widget). Grid/pack the returned CONTAINER — never the text
    widget itself — since the text widget's actual parent is the container."""
    background = text_kwargs.get("bg", COLORS["card"])
    container = bordered(parent, bg=background)
    text_widget = tk.Text(container, **text_kwargs)
    vscroll = ttk.Scrollbar(container, orient="vertical", command=text_widget.yview)
    text_widget.configure(yscrollcommand=vscroll.set, wrap="word")
    text_widget.pack(side="left", fill="both", expand=True)
    vscroll.pack(side="right", fill="y")
    text_widget.bind("<FocusIn>", lambda e: container.configure(highlightbackground=COLORS["ring"],
                                                                highlightcolor=COLORS["ring"]), add="+")
    text_widget.bind("<FocusOut>", lambda e: container.configure(highlightbackground=COLORS["border"],
                                                                 highlightcolor=COLORS["border"]), add="+")
    return container, text_widget


def button_row(parent, *buttons, pady=(12, 0)):
    """A row of buttons, left-aligned: (text, command[, style])."""
    row = ttk.Frame(parent)
    row.pack(fill="x", pady=pady)
    made = []
    for i, spec in enumerate(buttons):
        text, command = spec[0], spec[1]
        style = spec[2] if len(spec) > 2 else "TButton"
        b = ttk.Button(row, text=text, command=command, style=style)
        b.pack(side="left", padx=(0 if i == 0 else 8, 0))
        made.append(b)
    return row, made


def field(parent, row, column, label, widget_factory, columnspan=1, hint=None):
    """A form field the kit's way: a small label above its input. Returns the input."""
    cell = ttk.Frame(parent)
    cell.grid(row=row, column=column, columnspan=columnspan, sticky="nwe",
              padx=(0 if column == 0 else 8, 0), pady=(0, 12))
    ttk.Label(cell, text=label, style="Field.TLabel").pack(anchor="w", pady=(0, 4))
    widget = widget_factory(cell)
    widget.pack(anchor="w", fill="x")
    if hint:
        WrapLabel(cell, text=hint, style="Hint.TLabel").pack(anchor="w", fill="x", pady=(3, 0))
    return widget


# ══════════════════════════════════════════════════════════════════════════
class RSVPApp(tk.Tk):
    def __init__(self, outlook: OutlookPort | None = None):
        super().__init__()
        # Every Outlook operation goes through self.outlook. The real one is
        # the outlook_com module, which satisfies OutlookPort as it stands;
        # tests pass a fake. This is the only place rsvp_app.py names
        # outlook_com - scripts/check_layering.py fails any other use.
        self.outlook: OutlookPort = outlook if outlook is not None else outlook_com
        self.title(APP_TITLE)
        self.geometry("1240x800")
        self.minsize(980, 640)
        # Before any widget exists: every style, font and default comes from here.
        self.fonts = theme.apply(self)

        # ── shared state across tabs ──
        self.recipients = []
        self.recipient_file = tk.StringVar(value="")  # the Excel file the list was last imported from
        self.responses = {}
        self._pending_recipients = []  # cập nhật mỗi lần Scan Inbox (Tab 4) — người trong Tab 2 chưa vote
        # Event ID của lần Scan Inbox gần nhất: the event Tab 4's table — and
        # everything derived from it — belongs to, which can differ from Tab 1's.
        self._last_scanned_event_id = None
        self._last_scan_time = None
        self._last_responses_roster = None
        self._vote_counts = {}
        # The event each auto-saved roster belongs to. Auto-saves write there,
        # never to whatever Event ID Tab 1 happens to show, and a roster owned
        # by another event is replaced when its tab opens.
        self._attendance_event = None
        self._attendance_roster = {}
        self._gift_event = None
        self._gift_roster = {}
        self._amount_paid_event = None
        self._amount_paid_quiet = False
        self._amount_paid_save_failed = False
        # Tab 5's payment rounds after the first, the first round's name and
        # the next "round_N": owned by _attendance_event together with the
        # table, and loaded or cleared only with it (_adopt_attendance_owner()).
        self._extra_rounds = []
        self._round1_label = ""
        self._next_round_index = 2
        self._round_vars = {}
        # Tab 6's gift item belongs to _gift_event; setting it from the
        # database must not save it straight back.
        self._gift_item_quiet = False
        # The event whose gift item is on screen: None until it was read, so
        # blank boxes shown after a failed read are never saved over it.
        self._gift_item_event = None
        # History's money columns follow every saved change on Tab 5 / Tab 6
        # (_sync_event_money()); muted while an event is loaded or cleared.
        self._suspend_money_sync = False
        # Saves that failed and were reported once already, by what they save.
        self._save_failures_reported = set()
        # While a paste is applied, cells it could not take are collected
        # here and reported once, not as one dialog per cell.
        self._paste_refusals = None
        self._yes_emails = []
        self._group_expansion_cache = {}
        # MỚI: đổi kiến trúc lưu trữ — history_path giờ trỏ tới 1 file SQLite
        # DUY NHẤT (rsvp_data.db, xem db.py) thay vì RSVP_History.xlsx. Toàn
        # bộ Events/Recipients/Gift Contribution/Attendance & Payment/
        # Responded result giờ sống trong DB này; Excel CHỈ còn được tạo ra
        # khi bấm nút "Export to Excel" ở từng tab, không tự động ghi liên
        # tục nữa (giữ tên biến "history_path" để không phải sửa lại quá
        # nhiều nơi — chỉ đổi Ý NGHĨA những gì nó trỏ tới).
        self.history_path = tk.StringVar(value=db.DB_FILE_DEFAULT)
        self.full_translations = {"en": "", "ja": "", "vi": "", "bilingual": ""}  # Copilot-translated full email bodies
        # MỚI: bản dịch Copilot RIÊNG cho mode "Send Gift Contribution Notice"
        # — tách khỏi self.full_translations (dùng cho Invite/Update invite)
        # để tránh 2 loại nội dung email hoàn toàn khác nhau ghi đè lẫn nhau
        # khi chỉ đổi qua lại Send mode trên CÙNG 1 ngôn ngữ.
        self.gift_full_translations = {"en": "", "ja": "", "vi": "", "bilingual": ""}
        # The text each draft generator last put in its box, so a box that
        # differs from it holds hand edits (see _hand_edited_drafts).
        self._generated_drafts = {}
        self.fixed_overrides = settings.load_fixed_overrides()  # user-customized default FIXED wording, persisted to disk
        self.prompt_overrides = settings.load_prompt_overrides()  # user-customized Copilot prompt templates, persisted to disk

        # MỚI: nhập 1 LẦN DUY NHẤT dữ liệu từ bộ file Excel CŨ (nếu có, từ
        # trước khi chuyển sang kiến trúc SQLite này) vào rsvp_data.db — an
        # toàn khi gọi lại mỗi lần mở app (tự bỏ qua nếu DB đã có dữ liệu
        # rồi, xem docstring migrate_from_excel_if_needed()). Thông báo kết
        # quả cho user biết SAU KHI cửa sổ chính đã hiện ra (self.after),
        # tránh popup chặn trước khi app kịp vẽ xong.
        try:
            migrated, event_count, migrate_notes = legacy_excel.migrate_from_excel_if_needed(
                db_path=self.history_path.get())
        except Exception:
            migrated, event_count, migrate_notes = False, 0, []
        if migrated and event_count:
            note_text = ("\n\n" + "\n".join(migrate_notes[:5])) if migrate_notes else ""
            self.after(600, lambda: messagebox.showinfo(
                "Old data imported",
                f"Found your previous RSVP_History.xlsx and imported {event_count} event(s) "
                f"(plus any matching Recipients/Gift Contribution/Attendance & Payment files) "
                f"into the new database (rsvp_data.db). Your old Excel files were NOT modified "
                f"or deleted — they're just no longer the primary source of data." + note_text
            ))


        if not HAS_TKCALENDAR:
            messagebox.showwarning(
                "tkcalendar not installed",
                "tkcalendar is not installed, so date fields will use plain text entry "
                "(type manually as dd/mm/yyyy).\n\n"
                "Install it to get a clickable calendar:\n"
                "pip install tkcalendar"
            )

        # ── shell: the kit's sidebar + header + page area ──
        shell = ttk.Frame(self)
        shell.pack(fill="both", expand=True)
        self.var_header_event = tk.StringVar(value="No event selected")
        self.var_sidebar_footer = tk.StringVar()
        self.sidebar = Sidebar(
            shell, self.fonts, APP_TITLE, "Event follow-up",
            [(i, i + 1, label) for i, (_attr, label, _t, _d) in enumerate(PAGES)],
            self._show_page, footer_var=self.var_sidebar_footer)
        self.sidebar.pack(side="left", fill="y")
        main = ttk.Frame(shell)
        main.pack(side="left", fill="both", expand=True)
        self.page_header = PageHeader(main, self.fonts, badge_var=self.var_header_event)
        self.page_header.pack(fill="x")
        tk.Frame(main, bg=COLORS["border"], height=1).pack(fill="x")

        # A notebook still holds the pages - selecting one fires
        # <<NotebookTabChanged>> - but its tab strip is hidden; the sidebar
        # is the navigation.
        nb = ttk.Notebook(main, style="Pages.TNotebook")
        nb.pack(fill="both", expand=True)
        self.nb = nb

        self.tab_config = ScrollableFrame(nb)
        self.tab_recipients = ScrollableFrame(nb)
        self.tab_compose = ScrollableFrame(nb)
        self.tab_collect = ScrollableFrame(nb)
        self.tab_calendar = ScrollableFrame(nb)
        self.tab_gift = ScrollableFrame(nb)  # MỚI — theo dõi quyên góp quà tặng
        self.tab_history = ScrollableFrame(nb)
        self._pages = [getattr(self, attr) for attr, _l, _t, _d in PAGES]
        for page in self._pages:
            nb.add(page)

        self._build_tab_config()
        self._build_tab_recipients()
        self._build_tab_compose()
        self._build_tab_collect()
        self._build_tab_calendar()
        self._build_tab_gift()
        self._build_tab_history()

        # Tab 3, 5 and 6 show content derived from Tab 1/2/4 (the email
        # preview, the attendance table, the gift list), so each is brought
        # up to date when the user SWITCHES TO it.
        nb.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        self.var_event_id.trace_add("write", lambda *a: self._refresh_header_event())
        self.var_event_name.trace_add("write", lambda *a: self._refresh_header_event())
        self.history_path.trace_add("write", lambda *a: self._refresh_sidebar_footer())
        self._refresh_header_event()
        self._refresh_sidebar_footer()
        self._show_page(0)
        self._mark_page(self.tab_config)

    def _show_page(self, index):
        self.nb.select(self._pages[index])

    def _mark_page(self, page):
        """Sidebar highlight and page header for the page now shown."""
        index = self._pages.index(page)
        _attr, _label, title, description = PAGES[index]
        self.sidebar.set_active(index)
        self.page_header.set(title, description)

    def _refresh_header_event(self):
        event_id = self.var_event_id.get().strip()
        name = self.var_event_name.get().strip()
        self.var_header_event.set(f"● {event_id}" + (f" · {name}" if name else "") if event_id
                                  else "No event selected")

    def _refresh_sidebar_footer(self):
        self.var_sidebar_footer.set(f"Database: {os.path.basename(self.history_path.get()) or '—'}")

    def _on_tab_changed(self, event):
        try:
            selected = event.widget.nametowidget(event.widget.select())
        except Exception:
            return
        if selected in self._pages:
            self._mark_page(selected)
        if selected is self.tab_compose:
            self._refresh_compose_preview_unless_edited()
        elif selected is self.tab_calendar:
            self._refresh_calendar_datetime_display()
            self._refresh_attendance_list()
            self._refresh_amount_paid_for_current_event()
            self._refresh_thankyou_body_display()
        elif selected is self.tab_gift:
            self._refresh_gift_contribution_list()

    # ── generic double-click-to-edit-a-cell helper, shared by the ──
    # ── Tab 5 "Attendance & Payment" table and the Tab 7 History table ──
    def _begin_cell_edit(self, tree, row_id, col_name, on_commit):
        """Opens a small Entry box directly on top of the given Treeview
        cell so the user can type a new value in place. Calls
        on_commit(row_id, col_name, new_value) when the edit is confirmed
        (Enter, or clicking away / losing focus). Escape cancels without
        saving. This does NOT write anything to disk by itself — each
        caller's on_commit decides what to do with the new value (e.g.
        update an in-memory roster dict, or just update the Treeview row)."""
        try:
            x, y, width, height = tree.bbox(row_id, col_name)
        except Exception:
            return
        if not width or not height:
            return  # cell is scrolled out of view — bbox() returns empty
        value = tree.set(row_id, col_name)
        entry = tk.Entry(tree)
        entry.insert(0, value)
        entry.select_range(0, "end")
        entry.focus()
        entry.place(x=x, y=y, width=width, height=height)

        done = {"value": False}

        def commit(event=None):
            if done["value"] or not entry.winfo_exists():
                return
            done["value"] = True
            new_value = entry.get()
            entry.destroy()
            on_commit(row_id, col_name, new_value)

        def cancel(event=None):
            done["value"] = True
            entry.destroy()

        entry.bind("<Return>", commit)
        entry.bind("<FocusOut>", commit)
        entry.bind("<Escape>", cancel)

    def _begin_cell_edit_combobox(self, tree, row_id, col_name, values, on_commit):
        """Same idea as _begin_cell_edit(), but opens a READONLY
        ttk.Combobox restricted to `values` instead of a free-text Entry —
        used for the "Actual Attend" column on Tab 5's Attendance &
        Payment table, so it's a proper Yes/No dropdown rather than free
        text. Commits as soon as a value is picked (<<ComboboxSelected>>),
        or on Enter/losing focus; Escape cancels without saving."""
        try:
            x, y, width, height = tree.bbox(row_id, col_name)
        except Exception:
            return
        if not width or not height:
            return  # cell is scrolled out of view — bbox() returns empty
        value = tree.set(row_id, col_name)
        combo = ttk.Combobox(tree, values=list(values), state="readonly")
        if value in values:
            combo.set(value)
        combo.place(x=x, y=y, width=width, height=height)
        combo.focus()

        done = {"value": False}

        def commit(event=None):
            if done["value"] or not combo.winfo_exists():
                return
            done["value"] = True
            new_value = combo.get()
            combo.destroy()
            on_commit(row_id, col_name, new_value)

        def cancel(event=None):
            done["value"] = True
            combo.destroy()

        combo.bind("<<ComboboxSelected>>", commit)
        combo.bind("<Return>", commit)
        combo.bind("<FocusOut>", commit)
        combo.bind("<Escape>", cancel)

    def _on_editable_tree_double_click(self, tree, event, editable_cols, on_commit):
        """Double-click handler: figures out which cell was double-clicked,
        checks it's one of editable_cols (a set of column names — pass None
        to allow editing ANY column), then opens an inline edit box via
        _begin_cell_edit(). Bind with:
            tree.bind("<Double-1>", lambda e: self._on_editable_tree_double_click(
                tree, e, {"name", "amount"}, self._my_commit_handler))"""
        region = tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = tree.identify_row(event.y)
        col_name = self._tree_column_name_at(tree, tree.identify_column(event.x))
        if not row_id or col_name is None:
            return
        if editable_cols is not None and col_name not in editable_cols:
            return
        self._begin_cell_edit(tree, row_id, col_name, on_commit)

    @staticmethod
    def _tree_display_columns(tree):
        """The columns as shown, left to right - which differs from
        tree["columns"] once the user has dragged headings around."""
        shown = tree["displaycolumns"]
        if not shown or shown in ("#all", ("#all",)) or tuple(shown) == ("#all",):
            return tuple(tree["columns"])
        return tuple(shown)

    @classmethod
    def _tree_column_name_at(cls, tree, column_ref):
        """The column name behind identify_column()'s "#N" (N counts the
        columns as shown), or None."""
        try:
            index = int(str(column_ref).replace("#", "")) - 1
        except ValueError:
            return None
        shown = cls._tree_display_columns(tree)
        return shown[index] if 0 <= index < len(shown) else None

    def _enable_treeview_copy_paste(self, tree, on_commit=None, editable_cols=None):
        """Adds Ctrl+C / Ctrl+V clipboard support to a Treeview:
        - Ctrl+C copies the selected rows (or all rows, if none selected)
          as tab-separated text, one row per line, in column order —
          pastes cleanly straight into Excel.
        - Ctrl+V pastes tab-separated (or comma-separated) clipboard text
          — e.g. copied from Excel — back INTO the table, starting at the
          currently selected row and filling downward one Treeview row per
          pasted line. It never creates new rows, only fills existing
          ones; pasted columns beyond the table's column count are
          ignored, and if editable_cols is given, any column not in it is
          skipped so paste can't corrupt read-only columns like "No.".
          Each changed cell goes through on_commit(row_id, col_name,
          new_value) if given — so pasting triggers the exact same side
          effects as manually double-click-editing that cell (e.g.
          auto-filling Amount when "Actual Attend" is pasted as "Yes") —
          otherwise the Treeview cell is just updated directly."""
        def copy_selection(event=None):
            rows = tree.selection() or tree.get_children()
            shown = self._tree_display_columns(tree)  # as the user sees them
            lines = []
            for row_id in rows:
                cells = [tree.set(row_id, c) for c in shown]
                lines.append("\t".join("" if v is None else str(v) for v in cells))
            tree.clipboard_clear()
            tree.clipboard_append("\n".join(lines))
            return "break"

        def paste_clipboard(event=None):
            try:
                text = tree.clipboard_get()
            except Exception:
                return "break"
            all_rows = list(tree.get_children())
            if not all_rows:
                return "break"
            selected = tree.selection()
            start_row = selected[0] if selected else all_rows[0]
            start_index = all_rows.index(start_row) if start_row in all_rows else 0
            columns = self._tree_display_columns(tree)
            pasted_lines = [ln for ln in text.replace("\r\n", "\n").replace("\r", "\n").split("\n") if ln != ""]
            self._paste_refusals = []
            try:
                for offset, line in enumerate(pasted_lines):
                    target_index = start_index + offset
                    if target_index >= len(all_rows):
                        break  # only fills existing rows, never creates new ones
                    row_id = all_rows[target_index]
                    cells = line.split("\t") if "\t" in line else line.split(",")
                    for col_index, new_value in enumerate(cells):
                        if col_index >= len(columns):
                            break
                        col_name = columns[col_index]
                        if editable_cols is not None and col_name not in editable_cols:
                            continue
                        if on_commit is not None:
                            on_commit(row_id, col_name, new_value.strip())
                        else:
                            tree.set(row_id, col_name, new_value.strip())
            finally:
                # Back to one dialog per refused cell for typed edits, even
                # if a commit above raised.
                refused, self._paste_refusals = self._paste_refusals, None
            if refused:
                messagebox.showwarning(
                    "Some cells not changed",
                    f"{len(refused)} pasted value(s) are not amounts someone paid and were left as they "
                    f"were: {', '.join(refused[:5])}{' ...' if len(refused) > 5 else ''}")
            return "break"

        tree.bind("<Control-c>", copy_selection)
        tree.bind("<Control-C>", copy_selection)
        tree.bind("<Control-v>", paste_clipboard)
        tree.bind("<Control-V>", paste_clipboard)

    # ══════════════════════════════════════════════════════════════════
    # TAB 1 — EVENT SETUP
    # ══════════════════════════════════════════════════════════════════
    def _build_tab_config(self):
        f = self.tab_config.body
        gap = {"fill": "x", "pady": (0, 16)}

        # "Event mode" chỉ là NHÃN LƯU Ý (Event / Gift / Event + Gift) — không
        # khoá/ẩn tính năng nào; nó giúp nhớ lại mục đích sự kiện ở History.
        self.var_event_mode = tk.StringVar(value="Event")
        self.var_event_name = tk.StringVar(value="Team Building Q3 2026")
        self.var_event_id = tk.StringVar(value="TB2026-Q3")
        # Mỗi khi Event ID đổi (gõ tay hoặc do Load setup/Register event...),
        # tự cập nhật banner trạng thái ở đầu Tab 4 — báo ngay nếu bảng kết
        # quả đang hiện là của 1 Event ID KHÁC (còn sót từ lần Scan trước).
        self.var_event_id.trace_add(
            "write", lambda *a: self._update_scan_status_banner() if hasattr(self, "banner_scan") else None)
        self.var_location = tk.StringVar(value="")
        self.var_budget = tk.StringVar(value="")
        self.var_gift_budget = tk.StringVar(value="")
        self.var_organizer = tk.StringVar(value="")
        self.var_guest_of_honor = tk.StringVar(value="")
        self.var_start_time = tk.StringVar(value="18:00")
        self.var_end_time = tk.StringVar(value="21:00")

        details = Card(f, "Event details",
                       "Name, place and budget fill the subject and body of every email. The Event ID is "
                       "how replies are matched to this event: unique, no spaces or accents.")
        details.pack(**gap)
        g = details.body
        g.columnconfigure((0, 1), weight=1, uniform="form")
        field(g, 0, 0, "Event name", lambda p: ttk.Entry(p, textvariable=self.var_event_name))
        field(g, 0, 1, "Event ID", lambda p: ttk.Entry(p, textvariable=self.var_event_id))
        field(g, 1, 0, "Location", lambda p: ttk.Entry(p, textvariable=self.var_location))
        self.combo_event_mode = field(
            g, 1, 1, "Event mode",
            lambda p: ttk.Combobox(p, state="readonly", textvariable=self.var_event_mode,
                                   values=["Event", "Gift", "Event + Gift"]),
            hint="A reminder of what the event is for; every send mode stays available.")
        field(g, 2, 0, "Expected event budget", lambda p: ttk.Entry(p, textvariable=self.var_budget),
              hint='e.g. "5,000 JPY / person"')

        schedule = Card(f, "Date & deadlines",
                        "Start and end time also set the Calendar Invite on Attendance & payment.")
        schedule.pack(**gap)
        g = schedule.body
        g.columnconfigure((0, 1, 2), weight=1, uniform="form")
        self.date_event = field(g, 0, 0, "Event date",
                                lambda p: make_date_picker(p, datetime.now() + timedelta(days=21)))
        field(g, 0, 1, "Start time (HH:MM)", lambda p: ttk.Entry(p, textvariable=self.var_start_time))
        field(g, 0, 2, "End time (HH:MM)", lambda p: ttk.Entry(p, textvariable=self.var_end_time))
        self.date_deadline = field(g, 1, 0, "Response deadline",
                                   lambda p: make_date_picker(p, datetime.now() + timedelta(days=10)))

        # Organizer = người nhận đóng góp, Guest of Honor = người được tặng
        # quà — dùng cho email "Send Gift Contribution Notice" ở Tab 3.
        gift = Card(f, "Gift contribution",
                    "Only for events that collect money for a gift (Event mode Gift, or Event + Gift). "
                    "The contribution deadline can be later than the response deadline.")
        gift.pack(**gap)
        g = gift.body
        g.columnconfigure((0, 1), weight=1, uniform="form")
        field(g, 0, 0, "Guest of honor (who the gift is for)",
              lambda p: ttk.Entry(p, textvariable=self.var_guest_of_honor))
        field(g, 0, 1, "Organizer (collects the contributions)",
              lambda p: ttk.Entry(p, textvariable=self.var_organizer))
        field(g, 1, 0, "Expected gift budget", lambda p: ttk.Entry(p, textvariable=self.var_gift_budget),
              hint='e.g. "3,000 JPY / person"')
        self.date_gift_deadline = field(g, 1, 1, "Gift contribution deadline",
                                        lambda p: make_date_picker(p, datetime.now() + timedelta(days=10)))

        note = Card(f, "Organizer note",
                    "Free text in any language, shown right after the greeting. Translate it per "
                    "language on Compose & send.")
        note.pack(**gap)
        note_box, self.entry_note = make_scrollable_text(note.body, height=4)
        self.entry_note.insert("1.0", "Please respond before the deadline so we can prepare accurate numbers.")
        note_box.pack(fill="x")

        bottom = ttk.Frame(f)
        bottom.pack(fill="x")
        bottom.columnconfigure((0, 1), weight=1, uniform="half")
        save = Card(bottom, "Save this event",
                    "Writes the details above to History. No email is sent: to tell people about a "
                    "change, use 'Send update invite' on Compose & send.")
        save.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        ttk.Button(save.body, text="💾 Save event details", style="Primary.TButton",
                   command=self._update_history_from_tab1).pack(anchor="w")
        ttk.Button(save.body, text="🆕 Save & start a new event", command=self._register_event)\
            .pack(anchor="w", pady=(8, 0))
        ttk.Label(save.body, text="Starting a new event clears every page; this one stays in History.",
                  style="Hint.TLabel").pack(anchor="w", pady=(6, 0))

        reload = Card(bottom, "Reload a past event",
                      "Brings back its details, recipients, votes, attendance and gift list. "
                      "Outlook is not scanned.")
        reload.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        # Search: pick a History field (or any) and type words; the list below
        # keeps only the events containing all of them.
        self.var_history_search_field = tk.StringVar(value=HISTORY_SEARCH_ANY)
        self.var_history_search = tk.StringVar()
        search = ttk.Frame(reload.body, style="Card.TFrame")
        search.pack(fill="x", pady=(0, 6))
        self.combo_history_search_field = ttk.Combobox(
            search, width=16, state="readonly", textvariable=self.var_history_search_field,
            values=[HISTORY_SEARCH_ANY] + [h for _c, h, _w in HISTORY_TABLE_COLUMNS])
        self.combo_history_search_field.pack(side="left")
        entry_search = ttk.Entry(search, textvariable=self.var_history_search)
        entry_search.pack(side="left", fill="x", expand=True, padx=(6, 0))
        ttk.Button(search, text="Clear", style="Small.TButton",
                   command=self._clear_history_search).pack(side="left", padx=(6, 0))
        self.var_history_search.trace_add("write", lambda *a: self._apply_history_search())
        self.combo_history_search_field.bind("<<ComboboxSelected>>", lambda e: self._apply_history_search())
        self.lbl_history_search_result = ttk.Label(reload.body, text="", style="Hint.TLabel")
        self.lbl_history_search_result.pack(anchor="w", pady=(0, 6))
        # postcommand re-reads the database each time the list opens, so an
        # event saved from any tab shows up without a manual refresh.
        self.combo_load_history = ttk.Combobox(reload.body, state="readonly",
                                               postcommand=self._refresh_history_combo_values)
        self.combo_load_history.pack(fill="x")
        ttk.Button(reload.body, text="⬅ Load setup from selected event", command=self._load_from_history)\
            .pack(anchor="w", pady=(8, 0))

        self._refresh_history_combo()

    def _event_details_record(self):
        """Tab 1's event details as a History record — the part every save
        of the event writes."""
        return {
            "EventID": self.var_event_id.get().strip(),
            "EventName": self.var_event_name.get(),
            "EventDate": get_date_str(self.date_event),
            "Deadline": get_date_str(self.date_deadline),
            "Location": self.var_location.get(),
            "Budget": self.var_budget.get(),
            "OrganizerNote": self._source_note_text(),
            "EventMode": self.var_event_mode.get(),
            "Organizer": self.var_organizer.get(),
            "GuestOfHonor": self.var_guest_of_honor.get(),
            "GiftBudget": self.var_gift_budget.get(),
            "GiftDeadline": get_date_str(self.date_gift_deadline),
            "StartTime": self.var_start_time.get(),
            "EndTime": self.var_end_time.get(),
        }

    def _save_event_from_tab1(self):
        """Writes Tab 1 (plus RecipientFile, and Tab 4's Yes/No/Maybe counts
        when the table was scanned for this event) to History. Returns the
        Event ID, or None after telling the user why nothing was saved."""
        record = self._event_details_record()
        event_id = record["EventID"]
        if not event_id:
            messagebox.showwarning("Missing Event ID", "Enter an Event ID first.")
            return None
        if self.recipients:
            self._save_recipients_to_db(silent=True)
        if self.recipient_file.get():
            record["RecipientFile"] = self.recipient_file.get()
        if self._last_scanned_event_id == event_id and self._vote_counts:
            record.update(self._vote_counts_record())
        # What belongs on the History row but could not be saved while there
        # was none (Amount paid, the first round's name, the gift item):
        # written now, from the pages that hold this event.
        if self._amount_paid_event == event_id:
            record["AmountPaid"] = self.var_amount_paid.get()
        if self._attendance_event == event_id and self._round1_label:
            record["Round1Label"] = self._round1_label
        if self._gift_item_event == event_id:
            record.update(self._gift_item_fields())
        was_tracked = self._was_money_tracked(event_id)
        if "AmountPaid" in record and not was_tracked:
            # A payout typed before the event had a History row (or changed
            # since) reaches History only now: it is money worked on here, as
            # when it is typed (_on_amount_paid_changed()). The page's 0 over
            # nothing saved is not a payout.
            try:
                stored = (self._history_record(event_id) or {}).get("AmountPaid")
                was_tracked = parse_typed_amount(record["AmountPaid"]) != parse_typed_amount(stored)
            except Exception:
                pass
        try:
            db.save_event_record(record, self.history_path.get())
        except Exception as e:
            messagebox.showerror("Error", f"Couldn't save the event:\n{e}")
            return None
        self._save_failures_reported = {k for k in self._save_failures_reported if not k.startswith("fields:")}
        self._amount_paid_save_failed = False
        self._sync_event_money(event_id, was_tracked=was_tracked)
        self._refresh_history_tree()
        self._refresh_history_combo()
        return event_id

    def _register_event(self):
        """Lưu sự kiện ĐANG NHẬP (Tab 1) vào History — rồi RESET TOÀN BỘ form
        (Tab 1-6) về mặc định để bắt đầu nhập 1 sự kiện MỚI KHÁC. Sự kiện vừa
        lưu vẫn còn nguyên trong History — nạp lại bất cứ lúc nào qua
        '⬅ Load setup from selected event'."""
        event_id = self.var_event_id.get().strip()
        if not event_id:
            messagebox.showwarning("Missing Event ID", "Enter an Event ID before saving the event.")
            return

        if not messagebox.askyesno(
            "Save & start a new event",
            f"Save event '{event_id}' to History and start a new one?\n\n"
            "After saving, the WHOLE form (Tab 1 → 6: event info, recipient list, "
            "composed content and translations, scanned results...) will be CLEARED so you can "
            "start entering a NEW event. The event you just saved is still safely stored in "
            "History, and can be reloaded any time via '⬅ Load setup from selected event'."
        ):
            return

        if not self._save_event_from_tab1():
            return
        self._reset_for_new_event()
        messagebox.showinfo(
            "Saved",
            f"Event '{event_id}' saved to History.\n\n"
            "The form has been reset to defaults — ready to enter a new event."
        )

    def _clear_event_state(self):
        """Drops everything the app holds in memory for the current event,
        Tab 2 to Tab 6, so the next event starts from its own saved data and
        never inherits another event's recipients, translations, votes,
        attendance, gift ticks or Amount paid. Tab 1's fields are left to
        the caller. Long-lived preferences (saved FIXED wording, the Copilot
        prompt) are not touched.

        Nothing here writes to the database: Amount paid's trace is muted
        while it is reset, and every roster loses its owner first."""
        # Tab 2
        self.recipients = []
        self.recipient_file.set("")
        self._refresh_recipient_tree()

        # Tab 3 — translations are content of the old event's email
        self.full_translations = {"en": "", "ja": "", "vi": "", "bilingual": ""}
        self.gift_full_translations = {"en": "", "ja": "", "vi": "", "bilingual": ""}
        self.combo_send_mode.current(0)  # về lại "Send first Invite"
        self._refresh_send_button_label()
        self.txt_translation_paste.delete("1.0", "end")
        self.var_send_to_override.set("")
        self.var_auto_send.set(False)

        # Tab 4
        self.responses = {}
        self._pending_recipients = []
        self._last_scanned_event_id = None
        self._last_scan_time = None
        self._last_responses_roster = None
        self._vote_counts = {}
        self.tree_responses.delete(*self.tree_responses.get_children())
        self.tree_pending.delete(*self.tree_pending.get_children())
        self.lbl_summary.config(text="No responses scanned yet.")
        for var in (self.var_kpi_total, self.var_kpi_yes, self.var_kpi_no, self.var_kpi_maybe,
                    self.var_kpi_pending):
            var.set("0")
        self.banner_deadline.set("")
        self.txt_reminder_body.delete("1.0", "end")

        # Tab 5 — the two bodies are regenerated from Tab 1 when the tab opens
        self._yes_emails = []
        self.list_yes.delete(0, "end")
        self._amount_paid_event = None
        self._set_amount_paid_quietly("0")
        self._adopt_attendance_owner(None)  # the table, its rounds, the first round's name
        self.combo_calendar_lang.current(0)
        self.var_appt_body_default = ""
        self.txt_appt_body.delete("1.0", "end")
        self.combo_thankyou_lang.current(0)
        self.var_thankyou_body_default = ""
        self.txt_thankyou_body.delete("1.0", "end")

        # Tab 6
        self._gift_event = None
        self._gift_roster = {}
        self._adopt_gift_item(None)
        self.var_gift_search.set("")
        self.tree_gift.delete(*self.tree_gift.get_children())
        self._update_gift_contributed_count()
        self.txt_gift_reminder_body.delete("1.0", "end")
        self.txt_gift_report_body.delete("1.0", "end")

        self._update_scan_status_banner()

    def _reset_for_new_event(self):
        """Đưa TOÀN BỘ form về trạng thái mặc định — như vừa mở app lần đầu
        — để bắt đầu nhập 1 sự kiện HOÀN TOÀN MỚI."""
        # Tab 1 — Event ID first, so nothing below can be saved under it
        self.var_event_id.set("")
        self.var_event_name.set("")
        self.var_location.set("")
        self.var_budget.set("")
        self.var_gift_budget.set("")
        self.var_organizer.set("")
        self.var_guest_of_honor.set("")
        self.combo_event_mode.current(0)  # về lại "Event"
        set_date_str(self.date_event, (datetime.now() + timedelta(days=21)).strftime("%d/%m/%Y"))
        set_date_str(self.date_deadline, (datetime.now() + timedelta(days=10)).strftime("%d/%m/%Y"))
        set_date_str(self.date_gift_deadline, (datetime.now() + timedelta(days=10)).strftime("%d/%m/%Y"))
        self.var_start_time.set("18:00")
        self.var_end_time.set("21:00")
        self.entry_note.delete("1.0", "end")
        self.entry_note.insert("1.0", "Please respond before the deadline so we can prepare accurate numbers.")

        self._group_expansion_cache = {}
        self._clear_event_state()
        self._refresh_compose_preview()

    def _update_history_from_tab1(self):
        """Ghi các trường Tab 1 hiện tại — và cả RecipientFile nếu có — vào
        dòng History khớp Event ID, mà KHÔNG cần gửi email gì cả. Dùng khi bạn
        chỉ cần sửa lại thông tin đã lưu (vd sửa nhầm địa điểm) mà không cần
        báo cho người nhận. Nếu Event ID chưa từng có trong History, sẽ tạo
        dòng mới."""
        event_id = self._save_event_from_tab1()
        if not event_id:
            return
        messagebox.showinfo(
            "History updated",
            f"Saved the Tab 1 details of Event ID '{event_id}' to History.\n\n"
            "Other columns (Sent Date, Calendar Sent, etc.) were KEPT AS-IS — "
            "nothing was cleared or overwritten.\n\n"
            "⚠️ This does NOT send any email to recipients — if you want to notify them of the "
            "change, use 'Send update invite' mode on Tab 3 instead."
        )

    def _refresh_history_combo_values(self):
        """Re-reads the past-event list from the database and filters it by
        the search box, keeping the current selection. Runs every time the
        dropdown opens."""
        try:
            self._history_all = db.load_history(self.history_path.get())
        except Exception:
            self._history_all = []
        return self._apply_history_search(keep_selection=True)

    def _refresh_history_combo(self):
        """Like _refresh_history_combo_values(), then selects the newest event."""
        labels = self._refresh_history_combo_values()
        if labels and not self.var_history_search.get().strip():
            self.combo_load_history.current(len(labels) - 1)

    def _clear_history_search(self):
        self.var_history_search_field.set(HISTORY_SEARCH_ANY)
        self.var_history_search.set("")

    def _apply_history_search(self, keep_selection=False):
        """Fills the past-event list with the events matching the search:
        case-insensitive, every typed word must appear (in any order), in the
        chosen field or in any field. self._history_records is the list AS
        SHOWN, so the selected index always maps to the event the user sees.
        Returns the labels."""
        records = getattr(self, "_history_all", [])
        words = self.var_history_search.get().lower().split()
        field_label = self.var_history_search_field.get()
        key = next((c for c, h, _w in HISTORY_TABLE_COLUMNS if h == field_label), None)
        matched = []
        for rec in records:
            if key:
                haystack = str(rec.get(key) or "").lower()
            else:
                haystack = " ".join(str(rec.get(c) or "") for c in HISTORY_TABLE_KEYS).lower()
            if all(w in haystack for w in words):
                matched.append(rec)
        selected = self.combo_load_history.get()
        self._history_records = matched
        labels = []
        for rec in matched:
            label = f'{rec["EventID"]} — {rec["EventName"]}'
            # Searching one field: show its value, so it is clear why a row matched.
            if words and key and key not in ("EventID", "EventName"):
                value = str(rec.get(key) or "").strip().replace("\n", " ")
                if value:
                    label += f"   [{field_label}: {value[:40]}{'…' if len(value) > 40 else ''}]"
            labels.append(label)
        self.combo_load_history["values"] = labels
        if keep_selection and selected in labels:
            self.combo_load_history.current(labels.index(selected))
        elif words and labels:
            self.combo_load_history.current(0)
        elif not keep_selection and labels:
            self.combo_load_history.current(len(labels) - 1)
        else:
            self.combo_load_history.set("")
        if words:
            where = "any field" if key is None else field_label
            self.lbl_history_search_result.configure(
                text=(f"{len(matched)} of {len(records)} event(s) match in {where}." if matched
                      else f"No event matches in {where}."))
        else:
            self.lbl_history_search_result.configure(text=f"{len(records)} saved event(s).")
        return labels

    def _unsaved_work_if_switching(self):
        """What would be lost by replacing the current event in memory — work
        that exists in no database table, or in one the past-event list can
        never reach again."""
        lost = self._hand_edited_drafts()
        if any(t.strip() for t in self.full_translations.values()) or \
                any(t.strip() for t in self.gift_full_translations.values()):
            lost.append("the Copilot translations saved on Tab 3")
        if self.txt_translation_paste.get("1.0", "end").strip():
            lost.append("the text in Tab 3's translation paste box")
        event_id = self.var_event_id.get().strip()
        if self.recipients and event_id and not self._event_in_history(event_id):
            lost.append(f"the Tab 2 recipient list of '{event_id}', which is not saved in History "
                        f"(save it on Tab 1 to keep it reachable)")
        return lost

    def _hand_edited_drafts(self):
        """The email texts edited by hand: each box that is not empty and no
        longer holds what the app last generated for it. Loading another
        event clears or rebuilds every one of them."""
        def edited(box, generated):
            text = box.get("1.0", "end").strip()
            return bool(text) and text != (generated or "").strip()

        found = []
        generated = getattr(self, "_compose_generated", None)
        if generated is not None and self._compose_box_texts() != generated:
            found.append("your edits to the email on Compose & send")
        for label, box, text in (
                ("the reminder email on Collect responses", self.txt_reminder_body,
                 self._generated_drafts.get("reminder")),
                ("the calendar invite text on Attendance & payment", self.txt_appt_body,
                 self.var_appt_body_default),
                ("the thank-you email on Attendance & payment", self.txt_thankyou_body,
                 self.var_thankyou_body_default),
                ("the gift reminder email on Gift contribution", self.txt_gift_reminder_body,
                 self._generated_drafts.get("gift_reminder")),
                ("the contribution report email on Gift contribution", self.txt_gift_report_body,
                 self._generated_drafts.get("gift_report"))):
            if edited(box, text):
                found.append(f"your edits to {label}")
        return found

    def _event_in_history(self, event_id):
        try:
            return any(r.get("EventID") == event_id for r in db.load_history(self.history_path.get()))
        except Exception:
            return False

    def _load_from_history(self):
        # History's money columns are recalculated only after a save; nothing
        # here saves, but a half-loaded event must never be synced either.
        self._suspend_money_sync = True
        try:
            self._load_from_history_now()
        finally:
            self._suspend_money_sync = False

    def _load_from_history_now(self):
        idx = self.combo_load_history.current()
        if idx < 0 or idx >= len(self._history_records):
            messagebox.showinfo("Nothing selected", "Please select a past event from the list first.")
            return

        # Đọc lại TRỰC TIẾP từ database theo đúng Event ID vừa chọn — luôn
        # dùng dữ liệu MỚI NHẤT, không dùng bản cache của danh sách.
        stale_rec = self._history_records[idx]
        event_id_to_load = stale_rec.get("EventID")
        try:
            fresh_records = db.load_history(self.history_path.get())
        except Exception:
            fresh_records = self._history_records
        matches = [r for r in fresh_records if r.get("EventID") == event_id_to_load]
        rec = matches[0] if matches else stale_rec

        lost = self._unsaved_work_if_switching()
        if lost and not messagebox.askyesno(
                "Replace the current event?",
                f"Loading '{event_id_to_load}' replaces everything the app holds for the current "
                "event. This would be lost:\n\n• " + "\n• ".join(lost) + "\n\nContinue?"):
            return

        # Clear first, Event ID included, so no auto-save can write the old
        # event's data under the new ID (or the reverse) while loading.
        self.var_event_id.set("")
        self._clear_event_state()

        self.var_event_id.set(rec.get("EventID") or "")
        self.var_event_name.set(rec.get("EventName") or "")
        self.var_location.set(rec.get("Location") or "")
        self.var_budget.set(rec.get("Budget") or "")
        set_date_str(self.date_event, rec.get("EventDate"))
        set_date_str(self.date_deadline, rec.get("Deadline"))
        self.entry_note.delete("1.0", "end")
        self.entry_note.insert("1.0", rec.get("OrganizerNote") or "")

        # Dòng History CŨ (tạo trước khi các cột này tồn tại) có giá trị
        # None — dùng mặc định. GiftDeadline thiếu -> dùng Deadline (RSVP).
        event_mode = rec.get("EventMode") or "Event"
        if event_mode in self.combo_event_mode["values"]:
            self.var_event_mode.set(event_mode)
        self.var_organizer.set(rec.get("Organizer") or "")
        self.var_guest_of_honor.set(rec.get("GuestOfHonor") or "")
        self.var_gift_budget.set(rec.get("GiftBudget") or "")
        self.var_start_time.set(rec.get("StartTime") or "18:00")
        self.var_end_time.set(rec.get("EndTime") or "21:00")
        set_date_str(self.date_gift_deadline, rec.get("GiftDeadline") or rec.get("Deadline"))

        # "Amount paid" (Tab 5) đã lưu cho sự kiện này.
        self._amount_paid_event = event_id_to_load
        self._set_amount_paid_quietly(rec.get("AmountPaid") or "0")

        loaded_people = db.load_recipients(event_id_to_load, self.history_path.get())
        if loaded_people:
            self.recipients = loaded_people
            self._refresh_recipient_tree()
            recipient_note = f"\n\n📋 Auto-loaded {len(self.recipients)} people into Tab 2 from the database."
        else:
            recipient_note = "\n\n(This event has no recipient list saved yet — Tab 2 is empty.)"

        # Khôi phục kết quả đã lưu (Tab 4/5/6) — KHÔNG quét lại Outlook. Chỉ
        # khi bấm '📨 Scan Inbox for Vote results' ở Tab 4 mới thực sự quét lại.
        responded_loaded = self._load_responded_result_from_file(event_id_to_load)
        attendance_loaded = self._load_attendance_sheet_from_file(event_id_to_load)
        gift_loaded = self._load_gift_roster_from_db(event_id_to_load)
        restore_note = ""
        if responded_loaded:
            restore_note += ("\n\n📨 Restored the last-scanned Responded results into Tab 4 "
                              "(no live Outlook scan was performed — click "
                              "'📨 Scan Inbox for Vote results' there to refresh).")
        if attendance_loaded:
            restore_note += "\n\n📊 Restored the Attendance & Payment tracking table into Tab 5."
        if gift_loaded:
            restore_note += "\n\n🎁 Restored the Gift Contribution tracking table into Tab 6."
        self._refresh_compose_preview()

        messagebox.showinfo(
            "Loaded",
            "Event ID / Event Name / Location / Budget / Event Date / Deadline / Note "
            "have all been loaded from the past event."
            + recipient_note + restore_note +
            "\n\n⚠️ If you're creating a NEW event (using this one as a template): CHANGE the "
            "Event ID (and dates) before sending the invite, to avoid clashing with the old event. "
            "The recipient list on Tab 2 is saved under the new Event ID as soon as you edit it, "
            "save the event or send the invite. Votes, attendance, gift ticks and Amount paid stay "
            "with the original event.\n\n"
            "If you're reloading THIS SAME event just to Collect Responses / Send Calendar "
            "Invite, you can leave the Event ID as loaded."
        )

    # ══════════════════════════════════════════════════════════════════
    # TAB 2 — RECIPIENTS
    # ══════════════════════════════════════════════════════════════════
    def _build_tab_recipients(self):
        f = self.tab_recipients.body

        people = Card(f, "Recipient list",
                      "Import from Excel (Name in column A, Email in column B) or add people below. "
                      "The list is saved for the current Event ID after every change.")
        people.pack(fill="both", expand=True, pady=(0, 16))
        ttk.Button(people.actions, text="📊 Export to Excel", command=self._export_recipients_to_excel)\
            .pack(side="right")
        ttk.Button(people.actions, text="📥 Import from Excel...", style="Primary.TButton",
                   command=self._load_recipients).pack(side="right", padx=(0, 8))

        # Lọc TRỰC TIẾP bảng bên dưới khi gõ, trên CẢ Name lẫn Email.
        toolbar = ttk.Frame(people.body)
        toolbar.pack(fill="x", pady=(0, 12))
        self.var_recipient_search = tk.StringVar(value="")
        ttk.Label(toolbar, text="🔎", style="Muted.TLabel").pack(side="left", padx=(0, 6))
        ttk.Entry(toolbar, textvariable=self.var_recipient_search, width=36).pack(side="left")
        self.var_recipient_search.trace_add("write", lambda *a: self._apply_recipient_filter())
        ttk.Button(toolbar, text="✕ Clear", style="Ghost.TButton",
                   command=lambda: self.var_recipient_search.set("")).pack(side="left", padx=(6, 0))
        ttk.Button(toolbar, text="🔎 Expand group emails", command=self._expand_group_recipients)\
            .pack(side="right")

        cols = ("name", "email")
        tree_container, self.tree_recipients = make_scrollable_treeview(people.body, columns=cols, height=16)
        self.tree_recipients.heading("name", text="Name", anchor="w")
        self.tree_recipients.heading("email", text="Email", anchor="w")
        self.tree_recipients.column("name", width=280)
        self.tree_recipients.column("email", width=320)
        tree_container.pack(fill="both", expand=True)

        footer = ttk.Frame(people.body)
        footer.pack(fill="x", pady=(10, 0))
        self.lbl_recipient_count = ttk.Label(footer, text="No list loaded yet.", style="Muted.TLabel")
        self.lbl_recipient_count.pack(side="left")
        ttk.Label(footer, textvariable=self.recipient_file, style="Hint.TLabel").pack(side="right")
        WrapLabel(people.body, style="Hint.TLabel",
                  text="Expand group emails replaces a distribution list (e.g. 'EET Employees All') "
                       "with its real members, sub-groups included, so Collect responses can tell "
                       "exactly who has not answered.").pack(fill="x", pady=(8, 0))

        add = Card(f, "Add a person")
        add.pack(fill="x")
        g = add.body
        g.columnconfigure((0, 1), weight=1, uniform="form")
        self.var_new_name = tk.StringVar()
        self.var_new_email = tk.StringVar()
        field(g, 0, 0, "Name", lambda p: ttk.Entry(p, textvariable=self.var_new_name))
        field(g, 0, 1, "Email", lambda p: ttk.Entry(p, textvariable=self.var_new_email))
        buttons = ttk.Frame(g)
        buttons.grid(row=1, column=0, columnspan=2, sticky="w")
        ttk.Button(buttons, text="➕ Add person", style="Primary.TButton",
                   command=self._add_recipient_row).pack(side="left")
        ttk.Button(buttons, text="🗑 Delete selected row", command=self._delete_recipient_row)\
            .pack(side="left", padx=(8, 0))

    def _load_recipients(self):
        """Asks for an Excel file and replaces Tab 2's list with it. Reads
        the "DanhSach" sheet (or the first one): Name in column A, Email in
        column B, on any row - title and header rows are skipped because
        they have no address. Example rows and repeated addresses are
        dropped."""
        path = filedialog.askopenfilename(title="Import recipients from Excel",
                                          filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return False
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb["DanhSach"] if "DanhSach" in wb.sheetnames else wb.active
            rows = list(ws.iter_rows(min_row=1, values_only=True))
        except Exception as e:
            messagebox.showerror("Error", f"Couldn't read the file:\n{path}\n\n{e}")
            return False
        people, seen, duplicates = [], set(), 0
        for row in rows:
            name, email = (tuple(row) + (None, None))[:2]
            if not email or "@" not in str(email):
                continue
            if "(ví dụ)" in str(name or "") or "(example)" in str(name or "").lower():
                continue
            email = str(email).strip()
            if email.lower() in seen:
                duplicates += 1
                continue
            seen.add(email.lower())
            people.append((str(name or "").strip(), email))
        if not people:
            messagebox.showwarning(
                "No addresses found",
                f"No email address was found in column B of:\n{path}\n\nThe list was not changed.")
            return False
        self.recipients = people
        self.recipient_file.set(path)
        self._refresh_recipient_tree()
        self._save_recipients_to_db(silent=True)
        note = f"\n\n{duplicates} repeated address(es) were skipped." if duplicates else ""
        messagebox.showinfo("Imported", f"Imported {len(people)} people from:\n{path}{note}")
        return True

    def _save_recipients_to_db(self, silent=False):
        """Lưu self.recipients (danh sách người nhận đang có trên Tab 2)
        vào database cho Event ID hiện tại — đây là hàm PERSIST THẬT SỰ
        trong kiến trúc mới (thay cho việc ghi ra Excel trước đây). Gọi
        NGAY sau mọi thay đổi (Import/Add/Delete row/Expand group), và
        trước khi đăng ký/gửi sự kiện, để đảm bảo dữ liệu trong DB luôn
        khớp với những gì đang hiển thị trên Tab 2."""
        event_id = self.var_event_id.get().strip()
        if not event_id:
            if not silent:
                messagebox.showwarning(
                    "Missing Event ID",
                    "Enter an Event ID on Tab 1 first — it's needed to know which event "
                    "this recipient list belongs to in the database."
                )
            return False
        if not self.recipients:
            if not silent:
                messagebox.showwarning("No list yet", "There are no recipients to save yet.")
            return False
        try:
            db.save_recipients(event_id, self.recipients, self.history_path.get())
        except Exception as e:
            if not silent:
                messagebox.showerror("Error", f"Couldn't save recipients to the database:\n{e}")
            return False
        if not silent:
            messagebox.showinfo(
                "Saved",
                f"Saved {len(self.recipients)} people to the database for Event ID '{event_id}'."
            )
        return True

    def _refresh_recipient_tree(self):
        # Giờ chỉ là 1 lớp mỏng gọi _apply_recipient_filter() — hàm đó mới
        # thực sự đọc self.recipients + lọc theo ô search (nếu có), để mọi
        # nơi gọi _refresh_recipient_tree() (Load list, Add/Delete row,
        # Expand group, reset form...) đều tự động tôn trọng từ khoá đang
        # tìm thay vì hiện lại TOÀN BỘ danh sách và làm mất kết quả search.
        self._apply_recipient_filter()

    def _apply_recipient_filter(self, *args):
        query = self.var_recipient_search.get().strip().lower() if hasattr(self, "var_recipient_search") else ""
        self.tree_recipients.delete(*self.tree_recipients.get_children())
        matched = 0
        for name, email in self.recipients:
            if query and query not in (name or "").lower() and query not in (email or "").lower():
                continue
            self.tree_recipients.insert("", "end", values=(name, email))
            matched += 1
        total = len(self.recipients)
        if query:
            self.lbl_recipient_count.config(text=f"Showing {matched} / {total} people (search: '{self.var_recipient_search.get()}')")
        else:
            self.lbl_recipient_count.config(text=f"Total: {total} people")

    def _add_recipient_row(self):
        name, email = self.var_new_name.get().strip(), self.var_new_email.get().strip()
        if not email or "@" not in email:
            messagebox.showwarning("Missing email", "Enter a valid email before adding.")
            return
        if any(e.lower() == email.lower() for _, e in self.recipients):
            messagebox.showinfo("Already in the list", f"{email} is already on the list.")
            return
        self.recipients.append((name, email))
        self.var_new_name.set("")
        self.var_new_email.set("")
        self._refresh_recipient_tree()
        self._save_recipients_to_db(silent=True)

    def _delete_recipient_row(self):
        sel = self.tree_recipients.selection()
        if not sel:
            return
        for item_id in sel:
            values = self.tree_recipients.item(item_id, "values")
            self.recipients = [r for r in self.recipients if not (r[0] == values[0] and r[1] == values[1])]
        self._refresh_recipient_tree()
        self._save_recipients_to_db(silent=True)

    def _export_recipients_to_excel(self):
        """Exports Tab 2's list to an Excel file the user chooses, in the same
        layout Import reads (sheet "DanhSach", header on row 3), so the file
        can be edited and imported again."""
        if not self.recipients:
            messagebox.showwarning("No data", "There is no recipient list to export yet.")
            return
        event_id = self.var_event_id.get().strip()
        imported = self.recipient_file.get().strip()
        path = filedialog.asksaveasfilename(
            title="Export recipients to Excel",
            defaultextension=".xlsx",
            initialdir=os.path.dirname(imported) if imported else None,
            initialfile=f"Participant_List_{event_id or 'event'}.xlsx",
            filetypes=[("Excel files", "*.xlsx")])
        if not path:
            return
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "DanhSach"
            ws.cell(row=3, column=1, value="Họ tên").font = Font(bold=True)
            ws.cell(row=3, column=2, value="Email").font = Font(bold=True)
            for i, (name, email) in enumerate(self.recipients, start=4):
                ws.cell(row=i, column=1, value=name)
                ws.cell(row=i, column=2, value=email)
            ws.column_dimensions["A"].width = 28
            ws.column_dimensions["B"].width = 32
            wb.save(path)
        except Exception as e:
            messagebox.showerror("Error", f"Couldn't export the file:\n{e}")
            return
        messagebox.showinfo("Exported", f"Exported {len(self.recipients)} people to:\n{path}")

    def _expand_group_recipients(self):
        """Với mỗi dòng trong danh sách hiện tại, kiểm tra xem đó có phải 1
        group email (Exchange Distribution List) không — nếu phải, thay dòng
        đó bằng TẤT CẢ thành viên thật của group (kể cả sub-group lồng bên
        trong, xem self.outlook.expand_group_members()). Dòng nào KHÔNG phải
        group thì giữ nguyên. Chạy trong background thread vì có thể mất vài
        giây/group khi query Exchange GAL."""
        if not self.recipients:
            messagebox.showwarning("No list", "Load or add recipients first.")
            return

        original = list(self.recipients)

        def worker():
            new_list = []
            seen_emails = set()
            groups_expanded = []
            failed_groups = []
            groups_kept = []
            groups_partial = []
            diag_lines = []
            errors = []
            for name, email in original:
                try:
                    members, failed, diag = self.outlook.expand_group_members_detailed(email)
                except Exception as e:
                    members, failed, diag = None, [], []
                    errors.append(f"{email}: {e}")
                failed_groups.extend(failed)
                diag_lines.extend(diag)
                if not members or failed:
                    # Not a group - or one that listed nobody (no member with
                    # an email address), or only part of its people (a
                    # sub-group Outlook could not list): keep the row. Saving
                    # the part would lose the rest for good - the group's
                    # address, the only way back to them, would be gone; and
                    # keeping both would invite those listed twice.
                    # Không phải group (hoặc không resolve được) -> giữ nguyên dòng gốc
                    if members is not None and not failed:
                        groups_kept.append(name or email)
                    elif members is not None:
                        groups_partial.append(name or email)
                    key = email.lower()
                    if key not in seen_emails:
                        seen_emails.add(key)
                        new_list.append((name, email))
                    continue
                groups_expanded.append((name or email, len(members)))
                for m_name, m_email in members:
                    key = m_email.lower()
                    if key in seen_emails:
                        continue
                    seen_emails.add(key)
                    new_list.append((m_name, m_email))

            def apply_result():
                self.recipients = new_list
                self._refresh_recipient_tree()
                self._save_recipients_to_db(silent=True)
                lines = [f"Total after expanding: {len(new_list)} people (was {len(original)} rows)."]
                if groups_expanded:
                    lines.append("\nGroups expanded:")
                    lines.extend(f"  • {n} → {c} members" for n, c in groups_expanded)
                else:
                    lines.append("\nNo group email detected in the list — everyone was already an "
                                  "individual address (or Outlook couldn't resolve them as a group; "
                                  "see notes below if that's unexpected).")
                if diag_lines:
                    lines.append("\nStructure detected:")
                    lines.extend(f"  • {d}" for d in dict.fromkeys(diag_lines))
                if groups_kept:
                    lines.append("\nThese groups listed nobody with an email address, so each was "
                                  "kept as one row:")
                    lines.extend(f"  • {g}" for g in groups_kept)
                if groups_partial:
                    lines.append("\n⚠️ Not every member of these groups could be listed, so each "
                                  "was kept as one row, not expanded:")
                    lines.extend(f"  • {g}" for g in groups_partial)
                if failed_groups:
                    lines.append("\n⚠️ These (sub-)groups were found but their members could NOT "
                                  "be listed:")
                    lines.extend(f"  • {g}" for g in dict.fromkeys(failed_groups))
                    lines.append(
                        "\nThis usually means Outlook is in Cached Exchange Mode and the Offline "
                        "Address Book doesn't have that group's membership yet. Try: Send/Receive "
                        "→ Send/Receive Groups → Download Address Book, then run Expand again. "
                        "Otherwise, add those members by hand.")
                if errors:
                    lines.append("\n⚠️ Some rows could not be checked (Outlook/COM error):")
                    lines.extend(f"  • {e}" for e in errors)
                messagebox.showinfo("Expand groups — done", "\n".join(lines))

            self.after(0, apply_result)

        threading.Thread(target=worker, daemon=True).start()

    # ══════════════════════════════════════════════════════════════════
    # TAB 3 — COMPOSE & SEND (VOTING BUTTONS + LANGUAGE + TRANSLATION)
    # ══════════════════════════════════════════════════════════════════
    def _build_tab_compose(self):
        f = self.tab_compose.body
        gap = {"fill": "x", "pady": (0, 16)}

        # ── send mode + language (TOP-MOST — important choice, made first) ──
        message = Card(f, "Message", "What kind of email this is, and in which language.")
        message.pack(**gap)
        g = message.body
        g.columnconfigure((0, 1), weight=1, uniform="form")
        self.combo_send_mode = field(
            g, 0, 0, "Send mode",
            lambda p: ttk.Combobox(p, state="readonly", values=[
                "Send first Invite", "Send update invite", "Send Gift Contribution Notice"]),
            hint="'Update invite' adds a change notice and lets people vote again. 'Gift Contribution "
                 "Notice' announces the gift from Event setup, without voting buttons.")
        self.combo_send_mode.current(0)  # default: first invite
        self.combo_send_mode.bind("<<ComboboxSelected>>", lambda e: (self._refresh_send_button_label(), self._refresh_compose_preview()))
        self.combo_email_lang = field(
            g, 0, 1, "Email language",
            lambda p: ttk.Combobox(p, state="readonly", values=[
                LANG_LABELS["en"], LANG_LABELS["ja"], LANG_LABELS["vi"], LANG_LABELS["bilingual"]]))
        self.combo_email_lang.current(0)  # default: English
        self.combo_email_lang.bind("<<ComboboxSelected>>", lambda e: self._refresh_compose_preview())
        self.var_subject_preview = tk.StringVar(value="")
        self.var_greeting_preview = tk.StringVar(value="")
        preview = bordered(g, bg=COLORS["muted"], color=COLORS["border"])
        preview.grid(row=1, column=0, columnspan=2, sticky="we")
        inner = tk.Frame(preview, bg=COLORS["muted"], padx=12, pady=8)
        inner.pack(fill="x")
        for var in (self.var_subject_preview, self.var_greeting_preview):
            tk.Label(inner, textvariable=var, bg=COLORS["muted"], fg=COLORS["foreground"],
                     font=self.fonts.body, anchor="w", justify="left").pack(fill="x")

        # Email is assembled in THIS order: Greeting → EDITABLE → FIXED.
        content = Card(f, "Email content",
                       "Greeting, then your note, then the fixed part. Both boxes can be edited before "
                       "sending; they are rebuilt from Event setup when you come back here, unless you "
                       "edited them.")
        content.pack(**gap)
        ttk.Label(content.body, text="Your note — editable, any language", style="Field.TLabel")\
            .pack(anchor="w", pady=(0, 4))
        editable_container, self.txt_editable_preview = make_scrollable_text(
            content.body, width=40, height=5, bg=COLORS["success_wash"])
        editable_container.pack(fill="x")
        self._editable_container = editable_container
        # Shown in Bilingual only; _sync_translate_controls() packs it.
        self.lbl_bilingual_note_hint = WrapLabel(
            content.body, style="Hint.TLabel",
            text="Bilingual: write only your note here, in any language. 'Translate with Copilot' "
                 "below turns it into Japanese and English in one go, either side of the divider "
                 "line. The greeting and the fixed part are already in both languages.")
        ttk.Label(content.body, text="Fixed part — event details and voting instructions", style="Field.TLabel")\
            .pack(anchor="w", pady=(14, 4))
        fixed_container, self.txt_fixed_preview = make_scrollable_text(
            content.body, width=40, height=9, bg=COLORS["warning_wash"])
        fixed_container.pack(fill="x")
        WrapLabel(content.body, style="Hint.TLabel",
                  text="You can edit the fixed part too, and keep your wording for future events with "
                       "'Save FIXED wording as default'.").pack(fill="x", pady=(6, 0))
        button_row(content.body,
                   ("🔄 Refresh preview from Tab 1 / Tab 2", self._refresh_compose_preview),
                   ("💾 Save FIXED wording as default for this language", self._save_fixed_default),
                   ("↺ Reset FIXED wording", self._reset_fixed_default))

        # Set once and rarely touched, so it starts collapsed: the send
        # controls below are a long scroll away already.
        prompt = Card(f, "Copilot prompt",
                      "The instructions sent to Copilot with the email, for single-language targets. "
                      "Icons are already built in.")
        prompt.pack(**gap)
        self.var_prompt_toggle = tk.StringVar(value="▸ Show / edit")
        ttk.Button(prompt.actions, textvariable=self.var_prompt_toggle, style="Small.TButton",
                   command=self._toggle_prompt_editor).pack()
        self.prompt_editor = ttk.Frame(prompt.body)
        WrapLabel(self.prompt_editor, style="Hint.TLabel",
                  text="This is exactly what will be sent to Copilot (the system default, with emoji "
                       "instructions ⏰📍💰📋👥). Edit it and Save to keep your version until you Reset. "
                       "Bilingual uses its own built-in prompt, which sends only your note.")\
            .pack(fill="x", pady=(0, 6))
        prompt_container, self.txt_custom_prompt = make_scrollable_text(
            self.prompt_editor, width=40, height=11, bg=COLORS["muted"])
        # Always show SOMETHING — the saved override if present, otherwise the
        # system default (with icon instructions) — so the user can see exactly
        # what will be sent, and edit it directly instead of starting from blank.
        current_prompt = self.prompt_overrides.get("single", "").strip() or DEFAULT_PROMPT_SINGLE
        self.txt_custom_prompt.insert("1.0", current_prompt)
        prompt_container.pack(fill="x")
        button_row(self.prompt_editor,
                   ("💾 Save custom prompt as default", self._save_custom_prompt),
                   ("↺ Reset prompt to system default", self._reset_custom_prompt),
                   ("📘 Show system default prompt", self._show_system_prompt))

        # ── translation helper (Copilot bridge) ──
        translate = Card(f, "Translate with Copilot",
                         "Copies the email and the instructions to the clipboard; paste Copilot's "
                         "answer back here. A single language translates the whole email from the "
                         "Email language above. Bilingual sends only your note, in any language, and "
                         "gets it back in Japanese and English at once.")
        translate.pack(**gap)
        row = ttk.Frame(translate.body)
        row.pack(fill="x")
        ttk.Label(row, text="Translate into", style="Field.TLabel").pack(side="left", padx=(0, 8))
        self.combo_translate_target = ttk.Combobox(row, width=28, state="readonly", values=SINGLE_TARGETS)
        self.combo_translate_target.current(1)  # default Japanese
        self.combo_translate_target.pack(side="left")
        self.var_copy_translation_label = tk.StringVar(value="📋 Copy full email + prompt")
        ttk.Button(row, textvariable=self.var_copy_translation_label, style="Primary.TButton",
                   command=self._copy_email_for_translation).pack(side="left", padx=(8, 0))
        ttk.Label(translate.body, text="Translated email from Copilot", style="Field.TLabel")\
            .pack(anchor="w", pady=(14, 4))
        paste_container, self.txt_translation_paste = make_scrollable_text(translate.body, width=40, height=6)
        paste_container.pack(fill="x")
        WrapLabel(translate.body, style="Hint.TLabel",
                  text="A single language: used as the complete email, exactly as pasted. Bilingual: "
                       "the [JA] and [EN] parts become your note in each language. Pasting into a box "
                       "that is not empty appends — click '🗑 Clear' first. If words come out glued "
                       "together (a Copilot copy quirk), click 'Clean up', then 'Save' again.")\
            .pack(fill="x", pady=(6, 0))
        button_row(translate.body,
                   ("💾 Save as translated version for this language", self._save_translated_email),
                   ("🗑 Clear saved translation", self._clear_translated_email),
                   ("🧹 Clean up spacing", self._cleanup_pasted_text))
        self.lbl_translation_status = ttk.Label(
            translate.body, style="Muted.TLabel",
            text="Translation saved: EN ❌  |  JA ❌  |  VI ❌  |  Bilingual note ❌")
        self.lbl_translation_status.pack(anchor="w", pady=(10, 0))

        # ── send controls ──
        send = Card(f, "Send", "Opens the email in Outlook for you to check and send, unless you tick "
                               "'Send immediately'.")
        send.pack(fill="x")
        g = send.body
        g.columnconfigure(0, weight=1)
        self.var_send_to_override = tk.StringVar(value="")
        field(g, 0, 0, "Send to (optional)", lambda p: ttk.Entry(p, textvariable=self.var_send_to_override),
              hint="A group address to send to instead of each person on Recipients. Votes are "
                   "tracked against the Recipients list either way.")
        self.var_auto_send = tk.BooleanVar(value=False)
        ttk.Checkbutton(g, text="Send immediately without review (unchecked = open Outlook for you to click Send)",
                        variable=self.var_auto_send).grid(row=1, column=0, sticky="w", pady=(0, 12))
        self.var_send_btn_label = tk.StringVar(value="✉ Send Invite via Outlook (Voting Buttons)")
        ttk.Button(g, textvariable=self.var_send_btn_label, style="Primary.TButton",
                   command=self._send_invite).grid(row=2, column=0, sticky="w")

        self._refresh_send_button_label()
        self._refresh_compose_preview()

    def _current_lang_code(self):
        label = self.combo_email_lang.get()
        return LANG_LABEL_TO_CODE.get(label, "en")

    def _is_update_mode(self):
        return self.combo_send_mode.get() == "Send update invite"

    def _is_gift_mode(self):
        return self.combo_send_mode.get() == "Send Gift Contribution Notice"

    def _refresh_send_button_label(self):
        if not hasattr(self, "var_send_btn_label"):
            return
        if self._is_gift_mode():
            self.var_send_btn_label.set("🎁 Send Gift Contribution Notice via Outlook")
        else:
            self.var_send_btn_label.set("✉ Send Invite via Outlook (Voting Buttons)")

    def _toggle_prompt_editor(self):
        if self.prompt_editor.winfo_manager():
            self.prompt_editor.pack_forget()
            self.var_prompt_toggle.set("▸ Show / edit")
        else:
            self.prompt_editor.pack(fill="x")
            self.var_prompt_toggle.set("▾ Hide")

    def _active_full_translations(self):
        """Dict lưu bản dịch Copilot đầy đủ ĐANG DÙNG — tách riêng cho mode
        Gift (self.gift_full_translations) và mode Invite/Update invite
        (self.full_translations), để 2 loại nội dung hoàn toàn khác nhau
        không ghi đè lẫn nhau khi đổi qua lại Send mode trên cùng 1 ngôn
        ngữ."""
        return self.gift_full_translations if self._is_gift_mode() else self.full_translations

    def _refresh_subject_preview_only(self):
        """Cập nhật CHỈ phần Subject preview (không đụng tới 2 ô nội dung
        EDITABLE/FIXED) — dùng khi Gửi mail hoặc Copy-để-dịch, để KHÔNG xoá
        mất nội dung bạn vừa gõ TAY TRỰC TIẾP vào ô EDITABLE trên Tab 3.

        BUG ĐÃ SỬA: trước đây cả '📨 Send' lẫn '📋 Copy full email' đều gọi
        _refresh_compose_preview() (bản ĐẦY ĐỦ) trước khi đọc nội dung —
        hàm đó NẠP LẠI ô Editable từ ghi chú gốc ở Tab 1 (self.entry_note),
        nên nếu bạn gõ thêm/sửa trực tiếp vào ô Editable trên Tab 3 (thay vì
        quay lại Tab 1) mà CHƯA đồng bộ ngược, phần vừa gõ đó bị XOÁ MẤT
        ngay trước khi gửi/copy — y hệt hiện tượng bạn gặp."""
        lang_code = self._current_lang_code()
        event_name = self.var_event_name.get()
        event_id = self.var_event_id.get().strip()
        if self._is_gift_mode():
            subject = build_gift_subject(lang_code, event_id, self.var_guest_of_honor.get())
        else:
            subject = build_subject(lang_code, event_id, event_name, is_update=self._is_update_mode())
        self._compose_subject = subject
        self.var_subject_preview.set(f"Subject preview: {subject}")

    def _editable_box_text_for_translation(self):
        """Đọc nội dung THẬT ĐANG CÓ trên ô EDITABLE của Tab 3 — đúng như
        bạn đang thấy/đã gõ trên màn hình (kể cả gõ tay trực tiếp vào đó) —
        bỏ đi hậu tố '[not yet translated...]' nếu có (hậu tố đó chỉ để
        hiển thị trong app, không phải nội dung thật cần dịch)."""
        text = self.txt_editable_preview.get("1.0", "end").strip()
        for flag in NOT_TRANSLATED_FLAG.values():
            flag = flag.strip()
            if flag and text.endswith(flag):
                text = text[: -len(flag)].rstrip()
                break
        return text

    def _source_note_text(self):
        return self.entry_note.get("1.0", "end").strip()

    def _lang_content(self, lang_code, event_name, event_date, location, deadline, budget):
        """Returns (fixed_text_or_None, editable_text, using_override) for a SINGLE
        language (en/ja/vi — not bilingual). fixed_text is None when a full Copilot-
        translated override is active for this language — in that case editable_text
        already holds the COMPLETE translated email (greeting + note + event details)."""
        override = self.full_translations.get(lang_code, "").strip()
        if override:
            return None, override, True
        fixed = self.fixed_overrides.get(lang_code, "").strip() or \
            build_fixed_block(lang_code, event_name, event_date, location, deadline, budget)
        editable = build_editable_block(lang_code, self._source_note_text(), False)
        return fixed, editable, False

    def _fixed_text(self, lang_code):
        """The fixed part in ONE language as Tab 1 describes the event now: the
        gift notice's in Gift mode, otherwise the saved FIXED wording or the
        built-in one. Never a saved Copilot translation."""
        event_date = get_date_str(self.date_event)
        location = self.var_location.get()
        if self._is_gift_mode():
            return build_gift_fixed_block(
                lang_code, self.var_guest_of_honor.get(), self.var_start_time.get(), event_date,
                location, self.var_organizer.get(), get_date_str(self.date_gift_deadline),
                self.var_gift_budget.get())
        return self.fixed_overrides.get(lang_code, "").strip() or build_fixed_block(
            lang_code, self.var_event_name.get(), event_date, location,
            get_date_str(self.date_deadline), self.var_budget.get())

    # ── MỚI: các hàm build nội dung riêng cho mode "Send Gift Contribution
    # Notice" — song song với _lang_content() ở
    # trên nhưng dùng build_gift_fixed_block() thay vì build_fixed_block(),
    # và đọc/ghi self.gift_full_translations (KHÔNG dùng self.full_translations,
    # để tránh 2 loại nội dung khác nhau ghi đè lẫn nhau khi đổi Send mode). ──
    def _gift_lang_content(self, lang_code, guest_of_honor, start_time, event_date, location,
                            organizer, deadline, gift_budget):
        override = self.gift_full_translations.get(lang_code, "").strip()
        if override:
            return None, override, True
        fixed = build_gift_fixed_block(lang_code, guest_of_honor, start_time, event_date, location,
                                        organizer, deadline, gift_budget)
        editable = build_editable_block(lang_code, self._source_note_text(), False)
        return fixed, editable, False

    def _refresh_compose_preview(self):
        self._fill_compose_preview()
        # What was generated, so a later refresh can tell hand edits apart.
        self._compose_generated = self._compose_box_texts()

    def _compose_box_texts(self):
        return (self.txt_editable_preview.get("1.0", "end"), self.txt_fixed_preview.get("1.0", "end"))

    def _refresh_compose_preview_unless_edited(self):
        """Runs when Tab 3 opens: rebuilds the preview from Tab 1, so it never
        sends details Tab 1 no longer holds - unless the EDITABLE or FIXED box
        was edited by hand, which a rebuild would throw away. Then only the
        subject is refreshed and the '🔄 Refresh preview' button stays the
        way to rebuild."""
        if self._compose_box_texts() == getattr(self, "_compose_generated", None):
            self._refresh_compose_preview()
        else:
            self._refresh_subject_preview_only()

    def _fill_compose_preview(self):
        lang_code = self._current_lang_code()
        event_name = self.var_event_name.get()
        event_id = self.var_event_id.get().strip()
        location = self.var_location.get()
        budget = self.var_budget.get()
        event_date = get_date_str(self.date_event)
        deadline = get_date_str(self.date_deadline)
        gift_deadline = get_date_str(self.date_gift_deadline)
        self._sync_translate_controls()

        # MỚI: nhánh RIÊNG hoàn toàn cho mode "Send Gift Contribution Notice"
        # — cùng cấu trúc UI (Subject/Greeting/Editable/Fixed/Bilingual) như
        # nhánh Invite bên dưới, nhưng dùng nội dung + subject + dict lưu bản
        # dịch Copilot RIÊNG cho Gift (xem _gift_lang_content()/
        # _fixed_text() ở trên) — return sớm, không chạy tiếp
        # xuống logic Invite/Update invite bên dưới. Dùng "Gift contribution
        # deadline" (RIÊNG, khác "Event response deadline" của RSVP) làm hạn
        # đóng góp trong nội dung email.
        if self._is_gift_mode():
            guest_of_honor = self.var_guest_of_honor.get()
            organizer = self.var_organizer.get()
            gift_budget = self.var_gift_budget.get()
            start_time = self.var_start_time.get()

            subject = build_gift_subject(lang_code, event_id, guest_of_honor)
            self._compose_subject = subject
            self.var_subject_preview.set(f"Subject preview: {subject}")

            if lang_code == "bilingual":
                self._fill_bilingual_preview()
                return
            self.var_greeting_preview.set(f"Greeting (auto, appears first): {build_greeting(lang_code)}")
            fixed_text, editable_text, override_used = self._gift_lang_content(
                lang_code, guest_of_honor, start_time, event_date, location, organizer, gift_deadline, gift_budget)
            fixed_display = fixed_text if fixed_text is not None else \
                "→ Using a full Copilot-translated email (see EDITABLE box below — this FIXED box is unused for this language)."
            self.txt_fixed_preview.config(state="normal")
            self.txt_fixed_preview.delete("1.0", "end")
            self.txt_fixed_preview.insert("1.0", fixed_display)
            if override_used:
                self.txt_fixed_preview.config(state="disabled")

            self.txt_editable_preview.delete("1.0", "end")
            self.txt_editable_preview.insert("1.0", editable_text)
            self._refresh_translation_status()
            return

        subject = build_subject(lang_code, event_id, event_name, is_update=self._is_update_mode())
        self._compose_subject = subject
        self.var_subject_preview.set(f"Subject preview: {subject}")

        if lang_code == "bilingual":
            self._fill_bilingual_preview()
            return
        self.var_greeting_preview.set(f"Greeting (auto, appears first): {build_greeting(lang_code)}")
        fixed_text, editable_text, override_used = self._lang_content(
            lang_code, event_name, event_date, location, deadline, budget)
        fixed_display = fixed_text if fixed_text is not None else \
            "→ Using a full Copilot-translated email (see EDITABLE box below — this FIXED box is unused for this language)."
        editable_display = editable_text
        if self._is_update_mode():
            # Chèn banner "thông tin đã thay đổi" NGAY ĐẦU ô Editable — vẫn
            # là văn bản có thể sửa/xoá tay như phần note bình thường, chỉ
            # là được tự động thêm sẵn khi chọn 'Send update invite'.
            editable_display = build_update_notice(lang_code) + "\n" + editable_display

        self.txt_fixed_preview.config(state="normal")
        self.txt_fixed_preview.delete("1.0", "end")
        self.txt_fixed_preview.insert("1.0", fixed_display)
        if override_used:
            self.txt_fixed_preview.config(state="disabled")
        # else: leave editable, per point 2 — user can hand-edit + save as new default

        self.txt_editable_preview.delete("1.0", "end")
        self.txt_editable_preview.insert("1.0", editable_display)

        self._refresh_translation_status()

    def _fill_bilingual_preview(self):
        """Bilingual, in both Send modes. The note box holds only your note -
        or, once Copilot's translation is saved, its Japanese and English
        versions either side of the divider line - and the fixed box holds the
        fixed part in Japanese and in English the same way. Both stay editable;
        _bilingual_body() assembles the email from them."""
        self.var_greeting_preview.set(
            f"Greeting (auto): {build_greeting('ja')} opens the Japanese half (first), "
            f"{build_greeting('en')} the English half (second).")
        halves = []
        for code in ("ja", "en"):
            fixed = self._fixed_text(code)
            if self._is_update_mode():
                fixed = build_update_notice(code) + "\n" + fixed
            halves.append(fixed)
        note = self._active_full_translations().get("bilingual", "").strip() or self._source_note_text()
        self.txt_fixed_preview.config(state="normal")
        self.txt_fixed_preview.delete("1.0", "end")
        self.txt_fixed_preview.insert("1.0", join_bilingual(*halves))
        self.txt_editable_preview.delete("1.0", "end")
        self.txt_editable_preview.insert("1.0", note)
        self._refresh_translation_status()

    def _sync_translate_controls(self):
        """Bilingual has one target - Japanese and English from your note, in
        whatever language it is written - so there is nothing to pick. A
        single language picks one of three and sends the whole email."""
        target = self.combo_translate_target
        if self._current_lang_code() == "bilingual":
            target.configure(values=[BILINGUAL_TARGET])
            target.set(BILINGUAL_TARGET)
            target.state(["disabled"])
            self.var_copy_translation_label.set("📋 Copy note + prompt (→ Japanese + English)")
            self.lbl_bilingual_note_hint.pack(fill="x", pady=(6, 0), after=self._editable_container)
        else:
            target.state(["!disabled"])
            target.configure(values=SINGLE_TARGETS)
            if target.get() not in SINGLE_TARGETS:
                target.set("Japanese")
            self.var_copy_translation_label.set("📋 Copy full email + prompt")
            self.lbl_bilingual_note_hint.pack_forget()

    def _refresh_translation_status(self):
        translations = self._active_full_translations()
        def mark(code):
            return "✅" if translations.get(code, "").strip() else "❌"
        self.lbl_translation_status.config(
            text=f"Translation saved: EN {mark('en')}  |  JA {mark('ja')}  |  "
                 f"VI {mark('vi')}  |  Bilingual note {mark('bilingual')}"
        )

    def _save_fixed_default(self):
        if self._is_gift_mode():
            messagebox.showinfo(
                "Not available for Gift mode",
                "The FIXED wording for 'Send Gift Contribution Notice' is always freshly built "
                "from Tab 1 (Guest of Honor/Organizer/Location/Deadline/Gift budget), so there's "
                "no cross-event default to save here.\n\n"
                "You can still hand-edit the FIXED box above before sending this specific email — "
                "that edit just won't be remembered for future events."
            )
            return
        lang_code = self._current_lang_code()
        if lang_code not in ("en", "ja", "vi"):
            messagebox.showinfo(
                "Switch language first",
                "Custom default wording is saved per single language.\n\n"
                "Switch 'Email language' to English, Japanese, or Vietnamese first."
            )
            return
        if self.full_translations.get(lang_code, "").strip():
            messagebox.showwarning(
                "Clear the Copilot translation first",
                "This language is currently using a full Copilot-translated email override, "
                "so the FIXED box isn't active. Clear that translation first (button below) "
                "if you want to edit and save the FIXED wording instead."
            )
            return
        text = self.txt_fixed_preview.get("1.0", "end").strip()
        if not text:
            messagebox.showwarning("Empty", "The FIXED box is empty.")
            return
        self.fixed_overrides[lang_code] = text
        settings.save_fixed_overrides(self.fixed_overrides)
        messagebox.showinfo(
            "Saved",
            f"Saved your edited wording as the new default FIXED text for "
            f"{self.combo_email_lang.get()}.\n\n"
            "It will be used automatically for this and future events (until you Reset it)."
        )

    def _reset_fixed_default(self):
        if self._is_gift_mode():
            messagebox.showinfo(
                "Not available for Gift mode",
                "There's no saved default to reset for Gift mode — the FIXED box is always freshly "
                "built from Tab 1. Click '🔄 Refresh preview from Tab 1 / Tab 2' instead to reload it."
            )
            return
        lang_code = self._current_lang_code()
        if lang_code not in ("en", "ja", "vi"):
            messagebox.showinfo(
                "Switch language first",
                "Switch 'Email language' to English, Japanese, or Vietnamese first."
            )
            return
        self.fixed_overrides[lang_code] = ""
        settings.save_fixed_overrides(self.fixed_overrides)
        self._refresh_compose_preview()
        messagebox.showinfo("Reset", f"{self.combo_email_lang.get()} FIXED wording reset to the system default.")

    def _save_custom_prompt(self):
        """Save the custom prompt template to disk for reuse."""
        custom_prompt = self.txt_custom_prompt.get("1.0", "end").strip()
        if not custom_prompt:
            messagebox.showwarning("Empty", "The prompt box is empty. Type a prompt, or click "
                                             "'Reset prompt to system default' to restore the built-in one.")
            return
        self.prompt_overrides["single"] = custom_prompt
        settings.save_prompt_overrides(self.prompt_overrides)
        messagebox.showinfo(
            "Saved",
            "Custom prompt template saved.\n\n"
            "It will be used for single-language translations (English/Japanese/Vietnamese "
            "as the TRANSLATE-INTO target) from now on — including this exact wording, icons, "
            "and instructions — until you Reset it."
        )

    def _reset_custom_prompt(self):
        """Reset prompt back to system default (with icon instructions) — NOT blank."""
        self.prompt_overrides["single"] = ""
        settings.save_prompt_overrides(self.prompt_overrides)
        self.txt_custom_prompt.config(state="normal")
        self.txt_custom_prompt.delete("1.0", "end")
        self.txt_custom_prompt.insert("1.0", DEFAULT_PROMPT_SINGLE)
        messagebox.showinfo("Reset", "Prompt reset to the system default (shown in the box above — "
                                      "this is exactly what gets sent to Copilot, icon instructions included).")

    def _show_system_prompt(self):
        """Display the full built-in default prompts (both single-language and bilingual),
        so the user can see exactly what's used when no custom override is saved."""
        messagebox.showinfo(
            "System Default Prompt — Single language (EN/JA/VI)",
            "Used when 'Translate into' = English / Japanese / Vietnamese, and no custom "
            "prompt is saved (this is also what's pre-filled in the editable box above):\n\n"
            + DEFAULT_PROMPT_SINGLE
        )
        messagebox.showinfo(
            "System Default Prompt — Bilingual (Japanese + English)",
            "Used when 'Email language' = Bilingual. It sends only your note, in any language, "
            "and asks for it in Japanese and English: the greeting and the fixed part are "
            "already written in both. This one is built in:\n\n"
            + DEFAULT_PROMPT_BILINGUAL
        )

    def _copy_email_for_translation(self):
        lang_code = self._current_lang_code()
        if lang_code == "bilingual":
            self._copy_note_for_bilingual_translation()
            return
        self._refresh_subject_preview_only()
        greeting = build_greeting(lang_code)
        # Rebuild FIXED content directly (bypass the preview box on purpose):
        # when a full-translation override is already saved for this language
        # (self.full_translations[lang_code]), _refresh_compose_preview() fills
        # txt_fixed_preview with a PLACEHOLDER string — "→ Using a full
        # Copilot-translated email (see EDITABLE box below...)" — not real
        # event details. Reading that placeholder here and sending it to
        # Copilot as "content to translate" produced garbage/stale-looking
        # results whenever you tried to Copy-for-translation again after
        # already having saved a translation for the current source language.
        # Always use the LIVE Tab-1 event details instead, regardless of
        # whatever override happens to be saved.
        fixed_text = self._fixed_text(lang_code)
        # BUG ĐÃ SỬA: trước đây đọc note từ Tab 1 (self._source_note_text()),
        # nghĩa là nếu bạn gõ thêm/sửa trực tiếp vào ô EDITABLE trên Tab 3
        # (thay vì quay lại Tab 1) thì phần đó bị BỎ QUA hoàn toàn khi Copy —
        # y hệt hiện tượng "nội dung đã gõ biến mất". Giờ đọc TRỰC TIẾP từ ô
        # Editable đang hiển thị trên Tab 3 — đúng nguyên văn những gì bạn
        # thấy trên màn hình lúc bấm Copy.
        editable_text = self._editable_box_text_for_translation()
        if not fixed_text and not editable_text:
            messagebox.showwarning("Nothing to translate", "The email preview is empty.")
            return

        target_label = self.combo_translate_target.get()

        # Email order is: Greeting -> organizer's note -> event details/voting instructions.
        note_part = editable_text if editable_text else "(no note was written for this event)"
        combined = f"{greeting}\n\n{note_part}\n\n{fixed_text}"

        # Build prompt.
        # IMPORTANT: for the single-language box, read directly from what's shown/edited
        # in txt_custom_prompt — that box is ALWAYS pre-filled with either the saved
        # override or the system default (which already includes icon instructions), so
        # "what you see in the box is exactly what gets sent" — no more silently falling
        # back to an old icon-less hardcoded prompt just because nothing was Saved yet.
        # (Bilingual never comes here: it has its own prompt, sending the note alone.)
        prompt_template = self.txt_custom_prompt.get("1.0", "end").strip() or DEFAULT_PROMPT_SINGLE
        # Replace [TARGET_LANGUAGE] placeholder with the actual language name
        prompt_template = prompt_template.replace("[TARGET_LANGUAGE]", target_label)
        prompt = prompt_template + "\n\n" + combined

        self.clipboard_clear()
        self.clipboard_append(prompt)

        stale_override_note = ""
        if self._active_full_translations().get(lang_code, "").strip():
            stale_override_note = (
                f"\n\n⚠️ Note: a translation is STILL SAVED for {self.combo_email_lang.get()} "
                "from earlier — the preview above will keep showing that OLD saved version "
                "until you paste and Save the NEW result below (or click '🗑 Clear saved "
                "translation' first if you don't want it replaced)."
            )

        messagebox.showinfo(
            "Copied",
            "The full email (greeting + note + event details) plus translation instructions "
            "were copied to the clipboard — freshly rebuilt from the CURRENT Tab 1/Tab 2 "
            "settings and note, regardless of any previously saved translation." + stale_override_note +
            "\n\nNext steps:\n"
            "1. Open Copilot (or any AI assistant)\n"
            "2. Paste (Ctrl+V) and send\n"
            "3. Copy the translated reply\n"
            "4. Come back here, paste it in the box below, and click "
            "'Save as translated version for this language'"
        )

    def _copy_note_for_bilingual_translation(self):
        """Bilingual: only the note goes to Copilot, in whatever language it is
        written, and comes back in Japanese and English in one answer. The
        greeting and the fixed part are already written in both, so they are
        not sent. A prompt saved in prompt_overrides["bilingual"] is not used:
        it was written for the old whole-email reply, which
        parse_bilingual_reply() cannot read."""
        note = self._editable_box_text_for_translation()
        if not note:
            messagebox.showinfo(
                "Nothing to translate",
                "Your note is empty. The greeting and the fixed part are already in Japanese and "
                "English, so this email needs no translation: you can send it as it is.")
            return
        self.clipboard_clear()
        self.clipboard_append(DEFAULT_PROMPT_BILINGUAL + "\n\n" + note)
        messagebox.showinfo(
            "Copied",
            "Your note and the instructions were copied to the clipboard. Copilot will answer "
            "with the note in Japanese after [JA] and in English after [EN], whatever language "
            "you wrote it in.\n\n"
            "Next steps:\n"
            "1. Open Copilot (or any AI assistant)\n"
            "2. Paste (Ctrl+V) and send\n"
            "3. Copy its answer\n"
            "4. Come back here, paste it in the box below, and click "
            "'Save as translated version for this language'\n\n"
            "The greeting and the fixed part are added in both languages automatically.")

    def _target_lang_code(self):
        target = self.combo_translate_target.get()
        return {
            "English": "en", "Japanese": "ja", "Vietnamese": "vi",
            "Bilingual (Japanese + English)": "bilingual",
        }.get(target, "en")

    def _cleanup_pasted_text(self):
        """Apply best-effort repair to the pasted box: first remove a
        duplicated whole-email copy if detected (see dedupe_pasted_translation
        docstring — this can come from Copilot's own reply, not just from
        re-pasting), then fix lost spaces/line breaks on what remains.
        Deduplication runs AUTOMATICALLY (no confirmation prompt) — duplicate
        content is never something you'd want to keep in the final email, and
        a Yes/No dialog here was too easy to accidentally dismiss/misread."""
        text = self.txt_translation_paste.get("1.0", "end").strip()
        if not text:
            messagebox.showwarning("Empty", "Nothing to clean up — the paste box is empty.")
            return

        working_text = text
        dedupe_note = ""
        if detect_possible_duplicate_paste(text):
            deduped = dedupe_pasted_translation(text)
            removed = len(text) - len(deduped)
            if removed > 20:
                working_text = deduped
                dedupe_note = (
                    f"\n\n🔁 Also detected and removed ~{removed} characters of "
                    "DUPLICATED content at the start (the whole email appeared "
                    "twice — this can happen in Copilot's own reply, e.g. a "
                    "malformed draft immediately followed by a self-corrected "
                    "version — not just from pasting twice). Kept the later copy."
                )

        cleaned = cleanup_pasted_translation(working_text)
        self.txt_translation_paste.delete("1.0", "end")
        self.txt_translation_paste.insert("1.0", cleaned)
        messagebox.showinfo(
            "Cleaned up — remember to Save",
            "Applied best-effort fixes: stray '**'/'***' markers removed, spacing "
            "added around icons, line breaks restored before '•' bullets and "
            "after Japanese '。' sentence endings." + dedupe_note +
            "\n\nThis can't perfectly reconstruct every lost space inside a plain "
            "sentence — please read through the result once before saving.\n\n"
            "⚠️ IMPORTANT: this only updates the box above. It does NOT change "
            "what actually gets sent until you click '💾 Save as translated "
            "version for this language' again."
        )

    def _save_translated_email(self):
        text = self.txt_translation_paste.get("1.0", "end").strip()
        if not text:
            messagebox.showwarning("Empty", "Paste the translated email first.")
            return

        dedupe_note = ""
        if detect_possible_duplicate_paste(text):
            deduped = dedupe_pasted_translation(text)
            removed = len(text) - len(deduped)
            if removed > 20:
                text = cleanup_pasted_translation(deduped)
                self.txt_translation_paste.delete("1.0", "end")
                self.txt_translation_paste.insert("1.0", text)
                dedupe_note = (
                    f"\n\n🔁 Also detected and automatically removed ~{removed} "
                    "characters of DUPLICATED content at the start before saving "
                    "(the whole email appeared twice — can happen in Copilot's own "
                    "reply, not just from pasting twice). Kept the later copy."
                )

        code = self._target_lang_code()
        if code == "bilingual":
            self._save_bilingual_note(text, dedupe_note)
            return
        self._active_full_translations()[code] = text

        # Auto-switch "Email language" to match what was just saved — this used
        # to be a manual step ("switch Email language to see it") that was easy
        # to forget, causing the tool to silently send the OLD default-template
        # email instead of the translation you just pasted. Now it's applied
        # immediately, so what you see in the preview is guaranteed to be what
        # gets sent.
        target_label = LANG_LABELS.get(code)
        if target_label:
            self.combo_email_lang.set(target_label)

        self._refresh_compose_preview()
        messagebox.showinfo(
            "Saved & Applied",
            f"Saved the {self.combo_translate_target.get()} translation, and switched "
            f"'Email language' to {target_label} so it's shown in the preview below "
            "and will be used when you click Send — no extra step needed." + dedupe_note
        )

    def _save_bilingual_note(self, text, dedupe_note):
        """Bilingual: Copilot's [JA] and [EN] parts become the note box's two
        halves. Nothing is saved unless both are found."""
        parts = parse_bilingual_reply(text)
        if parts is None:
            messagebox.showwarning(
                "Japanese and English parts not found",
                "Copilot's answer should have a line [JA] before the Japanese note and a line "
                "[EN] before the English one. If they are missing, type them into the box, then "
                "click Save again." + dedupe_note)
            return
        self._active_full_translations()["bilingual"] = join_bilingual(*parts)
        self._refresh_compose_preview()
        messagebox.showinfo(
            "Saved & Applied",
            "Your note is now in Japanese and English: the note box above shows both, either "
            "side of the divider line, and that is what Send uses.\n\n"
            "Each language gets its greeting and its fixed part automatically, Japanese first "
            "and English second." + dedupe_note)

    def _clear_translated_email(self):
        code = self._target_lang_code()
        self._active_full_translations()[code] = ""
        self.txt_translation_paste.delete("1.0", "end")
        self._refresh_compose_preview()

    def _compose_full_body(self):
        """Assemble the FINAL email body for sending.
        ALWAYS prioritize full translation override (from Copilot) if present.
        If no override exists, build manually from components."""
        lang_code = self._current_lang_code()
        if lang_code == "bilingual":
            # Its saved translation is only the note, already in the note box.
            return self._bilingual_body()

        # ✅ ALWAYS check for full translation override FIRST (from Copilot)
        # MỚI: dùng đúng dict theo mode hiện tại (Gift dùng self.gift_full_translations
        # riêng, không lẫn với self.full_translations của Invite/Update invite).
        full_override = self._active_full_translations().get(lang_code, "").strip()
        if full_override:
            # Use the COMPLETE translated email (greeting + note + details already in it)
            return full_override
        
        # If no override, build single language: greeting + editable + fixed
        # (đọc TRỰC TIẾP từ 2 ô đang hiển thị trên Tab 3 — đã đúng nội
        # dung Gift hay Invite tuỳ mode, vì _refresh_compose_preview() đã
        # điền đúng nội dung cho từng mode; không cần branch thêm ở đây)
        greeting = build_greeting(lang_code)
        fixed_text = self.txt_fixed_preview.get("1.0", "end").strip()
        editable_text = self.txt_editable_preview.get("1.0", "end").strip()
        return f"{greeting}\n\n{editable_text}\n\n{fixed_text}".strip()

    def _bilingual_body(self):
        """The bilingual email from Tab 3's two boxes as they show now. A note
        with no divider line is not translated and goes into both halves as
        written; _bilingual_ready_to_send() asks about that first, and refuses
        a fixed box that lost its divider line."""
        ja_note, en_note = split_bilingual(self.txt_editable_preview.get("1.0", "end"))
        ja_fixed, en_fixed = split_bilingual(self.txt_fixed_preview.get("1.0", "end"))
        return build_bilingual_body(ja_note, ja_note if en_note is None else en_note,
                                    ja_fixed, en_fixed or "")

    def _bilingual_ready_to_send(self):
        """False when the bilingual email cannot be put together, or when the
        note is not translated yet and you would rather translate it first."""
        if split_bilingual(self.txt_fixed_preview.get("1.0", "end"))[1] is None:
            messagebox.showwarning(
                "Divider line missing",
                "The fixed part no longer has the divider line (――――) between its Japanese and "
                "English halves, so the tool cannot tell which goes where.\n\n"
                "Put the line back, or click '🔄 Refresh preview from Tab 1 / Tab 2' to rebuild "
                "the fixed part.")
            return False
        note = self.txt_editable_preview.get("1.0", "end").strip()
        if note and split_bilingual(note)[1] is None:
            return messagebox.askyesno(
                "Note not translated",
                "Your note is not in Japanese and English yet, so it would appear exactly as "
                "written in both halves of the email.\n\n"
                "To translate it: '📋 Copy note + prompt' under Translate with Copilot, paste "
                "Copilot's answer into the box there, then Save.\n\n"
                "Continue with the note as written?")
        return True

    def _send_invite(self):
        if not self.recipients:
            messagebox.showwarning("No recipients", "Go to Tab 2 and load the recipient list first.")
            return
        event_id = self.var_event_id.get().strip()
        if not event_id:
            # Without it the subject carries no ID, so Scan Inbox could never
            # match the replies and nothing could be recorded in History.
            messagebox.showwarning("Missing Event ID", "Enter the Event ID on Tab 1 before sending.")
            return
        if self._current_lang_code() == "bilingual" and not self._bilingual_ready_to_send():
            return
        self._save_recipients_to_db(silent=True)
        self._refresh_subject_preview_only()
        subject = self._compose_subject
        body = self._compose_full_body()
        auto_send = self.var_auto_send.get()
        send_to_override = self.var_send_to_override.get().strip() or None
        is_update = self._is_update_mode()
        is_gift = self._is_gift_mode()
        # Everything the send needs is read from the widgets here, on the Tk
        # thread, before the worker starts.
        record = self._event_details_record()
        record["EmailLanguage"] = self.combo_email_lang.get()
        record["TotalInvited"] = len(self.recipients)
        if self.recipient_file.get():
            record["RecipientFile"] = self.recipient_file.get()
        request = InviteRequest(
            recipients=list(self.recipients),
            subject=subject,
            body=body,
            # Fixed: Scan Inbox, the Yes/Maybe calendar list, attendance and
            # reminders all read exactly these three answers.
            voting_options="Yes;No;Maybe",
            auto_send=auto_send,
            send_to_override=send_to_override,
            mode="gift" if is_gift else ("update" if is_update else "invite"),
            record=record,
        )
        history_path = self.history_path.get()

        def worker():
            try:
                # The send, and the History write that follows it, live in
                # rsvp/services/invite.py, where a test can check exactly
                # what reaches Outlook. The History row is written at send
                # time (SentDate or UpdateInviteDate by mode; neither for a
                # gift notice); later scans update its vote counts.
                result = send_invite(
                    self.outlook,
                    lambda record: db.save_event_record(record, history_path),
                    request,
                )
                if result.history_written:
                    self.after(0, self._refresh_history_tree)
                self.after(0, lambda: messagebox.showinfo("Done", result.message))
            except Exception as e:
                err_msg = str(e)
                self.after(0, lambda: messagebox.showerror(
                    "Error",
                    f"Could not send via Outlook:\n{err_msg}\n\n"
                    "Check: is Outlook desktop open & signed in? Is pywin32 installed?"
                ))

        threading.Thread(target=worker, daemon=True).start()

    # ══════════════════════════════════════════════════════════════════
    # TAB 4 — COLLECT RESPONSES
    # ══════════════════════════════════════════════════════════════════
    def _build_tab_collect(self):
        f = self.tab_collect.body
        gap = {"fill": "x", "pady": (0, 16)}

        scan = Card(f, "Scan Inbox",
                    "Finds every vote reply whose subject contains this Event ID — invite, update or "
                    "reminder alike — in every folder of your mailbox. Each person's latest vote counts.")
        scan.pack(**gap)
        ttk.Button(scan.actions, text="📨 Scan Inbox for Vote results", style="Primary.TButton",
                   command=self._collect_responses).pack()
        self.banner_scan = Banner(scan.body, self.fonts)
        self.banner_scan.pack(fill="x")
        WrapLabel(scan.body, style="Hint.TLabel",
                  text="The table only changes when you scan: after switching events, scan again (the line "
                       "above says which event the table shows). Results, manual corrections and the "
                       "Yes / No / Maybe counts in History are saved automatically.").pack(fill="x", pady=(8, 0))

        self.var_kpi_total = tk.StringVar(value="0")
        self.var_kpi_yes = tk.StringVar(value="0")
        self.var_kpi_no = tk.StringVar(value="0")
        self.var_kpi_maybe = tk.StringVar(value="0")
        self.var_kpi_pending = tk.StringVar(value="0")
        kpi_row(f, [
            ("Tracked", self.var_kpi_total, {}),
            ("Yes", self.var_kpi_yes, {"dot": COLORS["success_fill"]}),
            ("No", self.var_kpi_no, {"dot": COLORS["destructive"]}),
            ("Maybe", self.var_kpi_maybe, {"dot": COLORS["warning"]}),
            ("No response", self.var_kpi_pending, {"dot": COLORS["muted_foreground"]}),
        ]).pack(**gap)

        responses = Card(f, "Responses",
                         "Everyone on Recipients (group addresses expanded) and their latest vote.")
        responses.pack(**gap)
        # Tick "Manual edit" (or double-click "Vote") to correct a vote by hand
        # — see _on_response_manual_check_click()/_commit_response_vote_edit().
        WrapLabel(responses.body, style="Hint.TLabel",
                  text="Tick ✏️ (or double-click a Vote) to correct a vote by hand, e.g. someone told you "
                       "in person. Corrected rows are tinted blue and survive later scans, unless that "
                       "person votes again by email; untick to let the next scan update them. Yellow rows "
                       "voted but are not on the list.").pack(fill="x", pady=(0, 10))
        cols = ("manual_edit", "name", "email", "vote", "received")
        tree_container, self.tree_responses = make_scrollable_treeview(responses.body, columns=cols, height=10)
        headers = ["✏️", "Name", "Email", "Vote", "Received At"]
        widths = [56, 220, 260, 110, 150]
        for c, label, w in zip(cols, headers, widths):
            anchor = "center" if c == "manual_edit" else "w"
            self.tree_responses.heading(c, text=label, anchor=anchor)
            self.tree_responses.column(c, width=w, anchor=anchor, stretch=c != "manual_edit")
        self.tree_responses.bind("<Double-1>", self._on_response_tree_double_click)
        self.tree_responses.bind("<Button-1>", self._on_response_manual_check_click)
        tree_container.pack(fill="both", expand=True)
        self.lbl_summary = WrapLabel(responses.body, text="No responses scanned yet.", style="Muted.TLabel")
        self.lbl_summary.pack(fill="x", pady=(10, 0))

        # ══════════════════════════════════════════════════════════════
        # CHƯA PHẢN HỒI (pending), + soạn/gửi email nhắc nhở
        # ══════════════════════════════════════════════════════════════
        pending = Card(f, "Not yet responded",
                       "Updated by every scan. Remind them with an email that keeps the voting buttons.")
        pending.pack(fill="x")
        self.banner_deadline = Banner(pending.body, self.fonts)
        self.banner_deadline.pack(fill="x", pady=(0, 10))
        pending_cols = ("name", "email")
        pending_container, self.tree_pending = make_scrollable_treeview(pending.body, columns=pending_cols, height=6)
        for c, label, w in zip(pending_cols, ["Name", "Email"], [280, 320]):
            self.tree_pending.heading(c, text=label, anchor="w")
            self.tree_pending.column(c, width=w)
        pending_container.pack(fill="x")

        ttk.Label(pending.body, text="Reminder email", style="CardTitle.TLabel").pack(anchor="w", pady=(18, 8))
        rbtn = ttk.Frame(pending.body)
        rbtn.pack(fill="x")
        ttk.Label(rbtn, text="Language", style="Field.TLabel").pack(side="left", padx=(0, 8))
        self.combo_reminder_lang = ttk.Combobox(
            rbtn, width=26, state="readonly",
            values=[LANG_LABELS["en"], LANG_LABELS["ja"], LANG_LABELS["vi"], LANG_LABELS["bilingual"]],
        )
        self.combo_reminder_lang.current(0)  # default: English
        self.combo_reminder_lang.pack(side="left")
        # Changing the language regenerates the draft right away, like Tab 3.
        self.combo_reminder_lang.bind("<<ComboboxSelected>>", lambda e: self._generate_reminder_draft())
        ttk.Button(rbtn, text="🔄 Regenerate text", command=self._generate_reminder_draft)\
            .pack(side="left", padx=(8, 0))
        self.var_reminder_attach = tk.BooleanVar(value=True)
        ttk.Checkbutton(pending.body, text="📎 Attach original invite email (looked up in Sent Items)",
                        variable=self.var_reminder_attach).pack(anchor="w", pady=(10, 8))
        # Every reminder (RSVP here, Gift on Tab 6) only opens Outlook for
        # review; there is deliberately no auto-send option.
        reminder_text_container, self.txt_reminder_body = make_scrollable_text(
            pending.body, width=40, height=7)
        reminder_text_container.pack(fill="x")
        self.lbl_reminder_send = ttk.Button(
            pending.body, text="📨 Send reminder email to 0 people who haven't responded",
            style="Primary.TButton", command=self._send_reminder)
        self.lbl_reminder_send.pack(anchor="w", pady=(12, 0))

        self._update_scan_status_banner()

    def _collect_responses(self):
        if not self.recipients:
            messagebox.showwarning("No list", "Go to Tab 2 and load the recipient list first.")
            return
        event_id = self.var_event_id.get().strip()
        if not event_id:
            # An empty ID is contained in every subject, so the scan would
            # count the votes of every voting email in the mailbox.
            messagebox.showwarning("Missing Event ID",
                                   "Enter the Event ID on Tab 1 first — Scan Inbox looks for it "
                                   "in the Subject of the vote replies.")
            return
        recipients = list(self.recipients)

        def worker():
            # Only the Outlook calls run on this thread; the results are
            # merged and shown on the Tk thread by _apply_scan_results().
            try:
                # Every folder of the mailbox, not just the Inbox.
                responses, skipped = self.outlook.scan_voting_responses(event_id, scan_all=True)
                # Group rows on Tab 2 are tracked as their real members.
                roster = self._build_effective_roster(recipients)
            except Exception as e:
                err_msg = str(e)
                self.after(0, lambda: messagebox.showerror(
                    "Error", f"Could not scan folders:\n{err_msg}\n\n"
                             "Check that Outlook is open & signed in, and pywin32 is installed."
                ))
                return
            self.after(0, lambda: self._apply_scan_results(event_id, responses, skipped, roster))

        threading.Thread(target=worker, daemon=True).start()

    def _apply_scan_results(self, event_id, scanned, skipped, roster):
        """Merges a finished Scan Inbox into self.responses, redraws Tab 4 and
        saves it. A vote corrected by hand (see _commit_response_vote_edit())
        is kept unless a NEWER email vote from that person was found; without
        a time on either side the manual vote wins, rather than silently
        losing a manual correction."""
        # Outlook returns timezone-aware datetimes, while manual edits and
        # rows loaded from the database are naive local times; comparing the
        # two raises TypeError. The wall-clock value is already local time.
        for info in scanned.values():
            received = info.get("received")
            if received is not None and getattr(received, "tzinfo", None) is not None:
                info["received"] = received.replace(tzinfo=None)
        # The baseline is what is known for THIS event: the table in memory if
        # it was scanned for it, else what was saved for it (e.g. after a
        # restart). Never another event's votes; never nothing when votes are
        # saved, or the first scan after a restart would discard saved manual
        # corrections, and an empty one would wipe every saved vote.
        if self._last_scanned_event_id == event_id:
            previous = self.responses
        else:
            try:
                previous = db.load_responses(event_id, self.history_path.get())[0]
            except Exception:
                previous = {}
        found_before = sum(1 for info in previous.values() if not info.get("manual"))
        if not scanned and found_before:
            # Far more likely a lagging search index than every vote email
            # vanishing at once - and saving would wipe the stored votes.
            messagebox.showwarning(
                "No votes found",
                f"Scan Inbox found no vote replies for '{event_id}' this time, but {found_before} "
                "were found before. The previous results were kept and nothing was saved.\n\n"
                "Outlook's search index may still be updating — try again in a minute.")
            return
        merged = dict(scanned)
        for email_key, prev in previous.items():
            if not prev.get("manual"):
                continue
            new_scan = merged.get(email_key)
            prev_time = prev.get("received")
            new_time = new_scan.get("received") if new_scan else None
            if new_scan is None or prev_time is None or new_time is None or new_time <= prev_time:
                merged[email_key] = prev
        self.responses = merged
        # The banner at the top of Tab 4 reports which Event ID the table
        # belongs to, so stale results of another event are never mistaken
        # for the current one.
        self._last_scanned_event_id = event_id
        self._last_scan_time = datetime.now()
        self._refresh_response_tree(skipped, roster)
        self._autosave_scan_results()
        self._update_scan_status_banner()

    def _update_scan_status_banner(self):
        """Cập nhật dòng trạng thái ở đầu Tab 4 — cho biết bảng bên dưới
        đang hiển thị kết quả quét của Event ID nào, quét lúc nào, và CẢNH
        BÁO RÕ nếu Event ID ở Tab 1 đã đổi khác so với lần quét gần nhất
        (tức dữ liệu đang hiển thị là CŨ, chưa phản ánh sự kiện hiện tại)."""
        current_id = self.var_event_id.get().strip()
        last_id = getattr(self, "_last_scanned_event_id", None)
        last_time = getattr(self, "_last_scan_time", None)
        last_time_str = last_time.strftime("%Y-%m-%d %H:%M") if last_time else "?"

        if not current_id:
            text = "🔍 Current Event ID: (not entered on Tab 1 yet)"
            tone = "warning"
        elif last_id is None:
            text = (f"🔍 Current Event ID: '{current_id}' — NOT scanned yet this session. "
                    f"Click '📨 Scan Inbox for Vote results' to fetch results.")
            tone = "warning"
        elif last_id != current_id:
            text = (f"⚠️ The table below is showing results for the OLD Event ID ('{last_id}', scanned at "
                    f"{last_time_str}) — NOT the current Event ID ('{current_id}'). Click Scan Inbox "
                    f"again to refresh!")
            tone = "danger"
        elif not getattr(self, "_scan_in_history", True):
            text = (f"✅ Showing results for Event ID '{current_id}' — scanned at {last_time_str}. "
                    f"⚠️ This event is not in History yet, so its Yes/No/Maybe counts are not recorded "
                    f"there: save it on Tab 1 ('💾 Save event details').")
            tone = "warning"
        else:
            text = f"✅ Showing results for Event ID '{current_id}' — scanned at {last_time_str}."
            tone = "success"
        self.banner_scan.set(text, tone)

    def _table_belongs_to_tab1_event(self, what):
        """True when Tab 4's table was scanned for the Event ID on Tab 1.
        Everything derived from it - the not-yet-responded list, the
        Yes/Maybe list - names people of THAT event, so a {what} built from a
        stale table would reach the wrong audience. Warns and returns False
        otherwise."""
        current_id = self.var_event_id.get().strip()
        if current_id and current_id == self._last_scanned_event_id:
            return True
        scanned = self._last_scanned_event_id or "(nothing scanned yet)"
        messagebox.showwarning(
            "Scan this event first",
            f"The vote table on Tab 4 belongs to Event ID '{scanned}', not '{current_id}' on Tab 1, "
            f"so the {what} would go to the wrong people.\n\n"
            "Click '📨 Scan Inbox for Vote results' on Tab 4 first.")
        return False

    def _build_effective_roster(self, recipients=None):
        """Trả về roster THỰC TẾ để đối chiếu vote ở Tab 4 — mỗi dòng trong
        Tab 2 là group email (Exchange Distribution List) được TỰ ĐỘNG thay
        bằng các thành viên thật (đệ quy qua sub-group, xem
        self.outlook.expand_group_members()), để khung 'Đã/Chưa phản hồi'
        liệt kê đúng TỪNG NGƯỜI thay vì 1 dòng group email mơ hồ.
        `recipients` defaults to Tab 2's list; the scan passes a copy taken
        on the Tk thread.

        KHÔNG sửa self.recipients (Tab 2 vẫn giữ nguyên như đã lưu) — chỉ áp
        dụng cho việc THEO DÕI/HIỂN THỊ ở Tab 4. Muốn áp dụng vĩnh viễn vào
        chính Tab 2 (vd để lần gửi mời SAU tự đúng luôn từ đầu), dùng nút
        '🔎 Expand group emails' ở Tab 2.

        Kết quả mỗi group được CACHE lại (self._group_expansion_cache) để
        không phải hỏi lại Exchange GAL mỗi lần bấm Scan Inbox trong cùng
        phiên làm việc — chỉ query lần đầu tiên gặp mỗi group email."""
        # The merge and de-duplication rules live in rsvp/domain/roster.py,
        # which is testable without Outlook. Expansion itself needs the
        # address book, so it is passed in rather than imported there.
        # A group that listed nobody stays one row to chase - tracking nobody
        # for it would hide that its people never answered. One missing a
        # sub-group gives the people it could list AND keeps its own row for
        # the rest. Neither result is cached: the next scan asks again
        # (Outlook's offline address book may have been downloaded meanwhile).
        recipients = self.recipients if recipients is None else recipients
        names = {(email or "").lower(): name for name, email in recipients}
        ask_again = set()

        def expand(email):
            key = (email or "").lower()
            members, failed, _diag = self.outlook.expand_group_members_detailed(email)
            if members is not None and (failed or not members):
                ask_again.add(key)
            if members and failed:
                return list(members) + [(names.get(key) or email, email)]
            return members or None

        roster = merge_expanded_roster(recipients, expand, cache=self._group_expansion_cache)
        for key in ask_again:
            self._group_expansion_cache.pop(key, None)
        return roster

    def _refresh_response_tree(self, skipped=0, roster=None):
        roster = roster if roster is not None else self.recipients
        # MỚI: nhớ lại roster vừa dùng để vẽ bảng — dùng khi sửa tay 1 ô Vote
        # (_commit_response_vote_edit()) cần vẽ lại toàn bộ bảng mà KHÔNG phải
        # tính lại _build_effective_roster() (có thể gọi COM để mở rộng group
        # email, không nên chạy lại chỉ vì sửa 1 ô).
        self._last_responses_roster = roster
        self.tree_responses.delete(*self.tree_responses.get_children())
        self.tree_responses.tag_configure("extra", background=COLORS["warning_soft"])
        # MỚI: tô nền xanh nhạt cho các dòng có phiếu vote đã được SỬA TAY
        # (double-click ô Vote — xem _commit_response_vote_edit()), để phân
        # biệt trực quan với phiếu quét được thật sự từ email.
        self.tree_responses.tag_configure("manual", background=COLORS["info_soft"])
        counts = {"Yes": 0, "No": 0, "Maybe": 0, "No response": 0}
        matched_emails = set()
        pending = []  # list[(name, email)] — người trong roster mà CHƯA vote gì
        for name, email in roster:
            email_key = email.lower()
            matched_emails.add(email_key)
            r = self.responses.get(email_key)
            if r:
                vote = r["vote"]
                received = r["received"].strftime("%Y-%m-%d %H:%M") if r.get("received") else ""
                display_name = r.get("name") or name
                counts[vote] = counts.get(vote, 0) + 1
                tags = ("manual",) if r.get("manual") else ()
            else:
                vote, received, display_name = "No response", "", name
                counts["No response"] += 1
                pending.append((name, email))
                tags = ()
            manual_display = "✅" if (r and r.get("manual")) else "⬜"
            self.tree_responses.insert("", "end", values=(manual_display, display_name, email, vote, received),
                                        tags=tags)

        # BUG ĐÃ SỬA: trước đây, bất kỳ ai bấm vote mà KHÔNG có mặt trong danh
        # sách Recipients (Tab 2) — vd: thành viên của 1 group email đã gửi
        # tới qua "Send to override" ở Tab 3, mỗi người trả lời bằng chính hộp
        # thư cá nhân của họ chứ không phải địa chỉ group — thì phiếu vote đó
        # ĐÃ được tool tìm thấy (nằm sẵn trong self.responses) nhưng bị ÂM
        # THẦM BỎ QUA lúc hiển thị, vì vòng lặp trên chỉ duyệt qua roster.
        # Giờ `roster` đã tự MỞ RỘNG group email thành từng thành viên thật
        # (xem _build_effective_roster), nên phần lớn trường hợp này KHÔNG
        # còn xảy ra nữa — nhưng vẫn giữ lại khối bên dưới làm lưới an toàn
        # cho những phiếu vote thật sự "lạ" (vd trả lời từ 1 địa chỉ hoàn
        # toàn không liên quan tới roster/group đã biết). Hiển thị thêm các
        # phiếu "ngoài danh sách" này (tô màu vàng nhạt để phân biệt) và
        # CỘNG luôn vào Yes/No/Maybe — vì đó vẫn là phiếu vote thật, hợp lệ,
        # khớp đúng Event ID — để tổng số phản ánh đúng thực tế.
        extra_rows = []
        for email, r in self.responses.items():
            if email in matched_emails:
                continue
            vote = r["vote"]
            received = r["received"].strftime("%Y-%m-%d %H:%M") if r.get("received") else ""
            display_name = (r.get("name") or email) + "  ⚠ (outside known list/group)"
            counts[vote] = counts.get(vote, 0) + 1
            # Dòng "extra" (ngoài danh sách) ưu tiên tag "extra" (vàng nhạt) —
            # nếu vừa "extra" vừa "manual", vẫn hiện "extra" vì đó là tín hiệu
            # quan trọng hơn cần chú ý (ai đó không có trong roster).
            tags = ("manual",) if r.get("manual") else ("extra",)
            manual_display = "✅" if r.get("manual") else "⬜"
            extra_rows.append((email, manual_display, display_name, vote, received, tags))
        for email, manual_display, display_name, vote, received, tags in extra_rows:
            self.tree_responses.insert("", "end", values=(manual_display, display_name, email, vote, received),
                                        tags=tags)

        self._vote_counts = dict(counts)
        self.var_kpi_total.set(str(len(roster)))
        self.var_kpi_yes.set(str(counts.get("Yes", 0)))
        self.var_kpi_no.set(str(counts.get("No", 0)))
        self.var_kpi_maybe.set(str(counts.get("Maybe", 0)))
        self.var_kpi_pending.set(str(counts.get("No response", 0)))
        extra_note = (f"  |  ➕ {len(extra_rows)} people outside the list/group also responded "
                      f"(added to the counts below)") if extra_rows else ""
        group_note = "" if len(roster) == len(self.recipients) else \
            f"  (auto-expanded {len(self.recipients)} Tab 2 row(s) → {len(roster)} actual people)"
        self.lbl_summary.config(
            text=(f"Total tracked: {len(roster)}{group_note}  |  ✅ Yes: {counts.get('Yes',0)}  |  "
                  f"❌ No: {counts.get('No',0)}  |  ❔ Maybe: {counts.get('Maybe',0)}  |  "
                  f"⬜ No response: {counts.get('No response',0)}"
                  + extra_note
                  + (f"  |  (⚠ {skipped} emails matched the Event ID but were not valid votes)" if skipped else ""))
        )
        self._refresh_calendar_yes_list()
        self._refresh_pending_panel(pending)

    def _on_response_tree_double_click(self, event):
        """MỚI: double-click ô "Vote" trên bảng Tab 4 mở 1 dropdown Yes/No/
        Maybe để sửa tay — dùng cho những người chỉ đổi ý hoặc báo miệng
        thay vì bấm lại nút Vote trong email. Các cột khác (Name/Email/
        Received At) KHÔNG cho sửa trực tiếp ở đây (Name/Email vốn lấy từ
        Tab 2, sửa nhầm dễ làm lệch với danh sách gốc)."""
        tree = self.tree_responses
        region = tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = tree.identify_column(event.x)
        row_id = tree.identify_row(event.y)
        if not row_id or not col:
            return
        columns = tree["columns"]
        try:
            col_index = int(col.replace("#", "")) - 1
        except ValueError:
            return
        if col_index < 0 or col_index >= len(columns):
            return
        col_name = columns[col_index]
        if col_name != "vote":
            return
        self._begin_cell_edit_combobox(tree, row_id, col_name, ["Yes", "No", "Maybe"],
                                        self._commit_response_vote_edit)

    def _commit_response_vote_edit(self, row_id, col_name, new_value):
        """Ghi lại 1 phiếu vote SỬA TAY vào self.responses (dict nguồn dữ
        liệu chung, xem đầu class) rồi vẽ lại TOÀN BỘ Tab 4 — vì mọi nơi
        khác đọc vote (Tab 5 Yes/Maybe list qua _refresh_calendar_yes_list(),
        Tab 5 Attendance & Payment qua _refresh_attendance_list(), khung
        "Chưa phản hồi"/banner deadline qua _refresh_pending_panel()) đều
        LẤY DỮ LIỆU TỪ CHÍNH self.tree_responses/self.responses, nên chỉ cần
        vẽ lại đúng 1 chỗ này là toàn bộ các tab liên quan tự động khớp
        theo — KHÔNG cần sửa gì thêm ở nơi khác. Sửa tay cũng auto-save
        xuống database ngay (giống mọi auto-save khác trong app), và đánh
        dấu "manual": True để lần "📨 Scan Inbox for Vote results" tiếp
        theo KHÔNG vô tình ghi đè mất phiếu sửa tay này (xem
        _collect_responses() — chỉ ghi đè nếu tìm thấy email vote MỚI HƠN
        thời điểm sửa tay)."""
        new_value = (new_value or "").strip()
        if new_value not in ("Yes", "No", "Maybe"):
            return
        # Đọc lại email/name TRỰC TIẾP từ chính ô đang hiển thị trên dòng đó
        # (không dựa vào row_id/iid) — đơn giản và luôn đúng bất kể iid được
        # Tkinter tự sinh ra là gì.
        email = self.tree_responses.set(row_id, "email").strip()
        if not email:
            return
        email_key = email.lower()
        display_name = self.tree_responses.set(row_id, "name")
        # Bỏ hậu tố cảnh báo "⚠ (outside known list/group)" (nếu có, xem
        # _refresh_response_tree()) trước khi lưu — hậu tố đó chỉ để HIỂN
        # THỊ, không phải 1 phần của tên thật.
        display_name = display_name.split("  ⚠")[0].strip()
        existing = self.responses.get(email_key, {})
        self.responses[email_key] = {
            "name": existing.get("name") or display_name or email,
            "vote": new_value,
            "received": datetime.now(),  # thời điểm SỬA TAY — dùng để so sánh "mới hơn hay cũ hơn" ở lần Scan sau
            "manual": True,
        }
        self._refresh_response_tree(roster=getattr(self, "_last_responses_roster", None))
        self._autosave_scan_results()

    def _on_response_manual_check_click(self, event):
        """MỚI: single-click ô "Manual edit" (cột checkbox ✅/⬜, đầu tiên
        bên trái "Name") trên Tab 4:
          - Dòng CHƯA sửa tay (⬜) -> mở NGAY dropdown Yes/No/Maybe ở ô
            "Vote" cùng dòng đó để chọn giá trị mới (tick + sửa cùng lúc,
            đúng yêu cầu "tick chọn rồi manual cho phần vote") — commit
            qua chính _commit_response_vote_edit() ở trên, hàm đó tự đặt
            "manual": True nên ô checkbox sẽ tự chuyển thành ✅ sau khi
            bảng được vẽ lại.
          - Dòng ĐÃ sửa tay (✅) -> bấm lại để BỎ tick, chỉ xoá cờ "manual"
            (giữ nguyên giá trị Vote hiện tại) — để lần Scan Inbox tiếp
            theo được tự do cập nhật lại phiếu này từ email thật, không
            còn bị khoá bởi bản sửa tay cũ nữa."""
        tree = self.tree_responses
        region = tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = tree.identify_column(event.x)
        row_id = tree.identify_row(event.y)
        if not row_id or col != "#1":  # "#1" = cột "manual_edit" (cột đầu tiên)
            return
        email = tree.set(row_id, "email").strip()
        if not email:
            return
        email_key = email.lower()
        info = self.responses.get(email_key)
        if info and info.get("manual"):
            # Đang ✅ -> bấm để bỏ tick: chỉ xoá cờ manual, GIỮ NGUYÊN vote.
            info["manual"] = False
            self._refresh_response_tree(roster=getattr(self, "_last_responses_roster", None))
            self._autosave_scan_results()
        else:
            # Đang ⬜ -> bấm để tick: mở dropdown Vote ngay để chọn giá trị
            # sửa tay (commit sẽ tự đặt manual=True, xem _commit_response_vote_edit()).
            self._begin_cell_edit_combobox(tree, row_id, "vote", ["Yes", "No", "Maybe"],
                                            self._commit_response_vote_edit)

    def _refresh_pending_panel(self, pending):
        """Cập nhật khung 'Chưa phản hồi' phía dưới + banner hạn phản hồi +
        nút gửi nhắc nhở, dựa trên `pending` (list[(name,email)]) vừa tính
        được từ _refresh_response_tree(). Gọi mỗi lần Scan Inbox xong, đúng
        như yêu cầu 'khi user bấm Scan inbox, cập nhật cả 2 khung'."""
        self._pending_recipients = pending

        self.tree_pending.delete(*self.tree_pending.get_children())
        for name, email in pending:
            self.tree_pending.insert("", "end", values=(name, email))

        # ── banner hạn phản hồi (Response Deadline ở Tab 1) ──
        deadline_str = get_date_str(self.date_deadline)
        overdue = False
        try:
            d, m, y = deadline_str.split("/")
            deadline_date = datetime(int(y), int(m), int(d)).date()
            overdue = datetime.now().date() >= deadline_date
        except Exception:
            deadline_date = None

        if not pending:
            banner, tone = f"Response deadline {deadline_str} — everyone on Recipients has responded.", "success"
        elif deadline_date is None:
            banner, tone = (f"{len(pending)} people still haven't responded. (Couldn't read the deadline "
                            f"'{deadline_str}'.)"), "warning"
        elif overdue:
            banner, tone = (f"Response deadline reached ({deadline_str}) — {len(pending)} people still haven't "
                            f"replied. Review the reminder below, then send it."), "danger"
        else:
            days_left = (deadline_date - datetime.now().date()).days
            banner, tone = (f"Response deadline {deadline_str} ({days_left} days left) — "
                            f"{len(pending)} people still haven't replied."), "warning"
        self.banner_deadline.set(banner, tone)

        self.lbl_reminder_send.config(
            text=f"📨 Send reminder email to {len(pending)} people who haven't responded")

        # Tự động tạo sẵn nội dung nhắc nhở lần đầu (nếu ô đang trống), để
        # người dùng có ngay bản nháp mà không cần bấm "Tạo lại" trước.
        if not self.txt_reminder_body.get("1.0", "end").strip():
            self._generate_reminder_draft()

    def _lookup_sent_date_hint(self, event_id):
        """Tra cột SentDate trong RSVP_History.xlsx theo EventID — dùng làm
        hint để chọn ĐÚNG email trong Sent Items khi có nhiều email cùng
        chứa event_id trong Subject (vd: gửi nhắc nhiều lần). Dùng chung cho
        cả '📅 Send Calendar Invite' (Tab 5) và '📨 Gửi nhắc nhở' (Tab 4).
        Trả về datetime hoặc None nếu không tìm thấy/không đọc được."""
        if not event_id:
            return None
        try:
            for rec in db.load_history(self.history_path.get()):
                if rec.get("EventID") == event_id:
                    raw = rec.get("SentDate")
                    if raw:
                        raw_str = str(raw).strip()
                        for fmt in ("%Y-%m-%d %H:%M", "%Y-%m-%d %H:%M:%S"):
                            try:
                                return datetime.strptime(raw_str, fmt)
                            except ValueError:
                                continue
                    break
        except Exception:
            pass  # best-effort — History không đọc được thì bỏ qua hint
        return None

    def _generate_reminder_draft(self):
        """Tự động soạn sẵn nội dung email nhắc nhở dựa trên thông tin sự
        kiện đang có ở Tab 1 (Event Name/Date/Location/Deadline/Budget) +
        ngôn ngữ đang chọn ở combo_reminder_lang (English/Japanese/
        Vietnamese/Bilingual — giống 4 lựa chọn ở Tab 3). Người dùng xem/sửa
        tay trong ô bên dưới trước khi gửi — hàm này chỉ tạo BẢN NHÁP, không
        tự gửi gì cả."""
        event_name = self.var_event_name.get()
        event_date = get_date_str(self.date_event)
        location = self.var_location.get()
        deadline = get_date_str(self.date_deadline)
        budget = self.var_budget.get()

        lang_label = self.combo_reminder_lang.get() or LANG_LABELS["en"]
        lang_code = LANG_LABEL_TO_CODE.get(lang_label, "en")

        draft = build_reminder_body(lang_code, event_name, event_date, location, deadline, budget)
        self._generated_drafts["reminder"] = draft
        self.txt_reminder_body.delete("1.0", "end")
        self.txt_reminder_body.insert("1.0", draft)

    def _send_reminder(self):
        pending = getattr(self, "_pending_recipients", []) or []
        if not pending:
            messagebox.showinfo(
                "No one left to remind",
                "Everyone in Tab 2 has already responded (Yes/No/Maybe) — no need to send a reminder.\n\n"
                "If you think there are still people who haven't replied, click "
                "'📨 Scan Inbox for Vote results' again first to refresh."
            )
            return

        body = self.txt_reminder_body.get("1.0", "end").strip()
        if not body:
            messagebox.showwarning(
                "Empty content",
                "Click '🔄 Regenerate text' or type the content by hand before sending.")
            return

        event_id = self.var_event_id.get().strip()
        if not self._table_belongs_to_tab1_event("reminder"):
            return
        event_name = self.var_event_name.get()
        lang_label = self.combo_reminder_lang.get() or LANG_LABELS["en"]
        lang_code = LANG_LABEL_TO_CODE.get(lang_label, "en")
        # Subject PHẢI vẫn chứa đúng event_id (giống mail gốc) — để lần Scan
        # Inbox sau còn quét/khớp được các phiếu vote mới trả lời trên chính
        # email nhắc nhở này (cơ chế quét chỉ dựa vào event_id nằm trong
        # Subject, không phân biệt đó là mail mời gốc hay mail nhắc, và
        # không phân biệt ngôn ngữ — build_reminder_subject() luôn giữ
        # nguyên "[Reminder-{event_id}]"/"【リマインド-{event_id}】" ở đầu).
        subject = build_reminder_subject(lang_code, event_id, event_name)

        attach = self.var_reminder_attach.get()
        sent_date_hint = self._lookup_sent_date_hint(event_id) if attach else None
        pending_count = len(pending)
        history_path = self.history_path.get()

        def worker():
            try:
                mail, attached = self.outlook.send_reminder_email(
                    pending, subject, body,
                    auto_send=False,  # a reminder is only ever opened for review
                    attach_event_id=(event_id if attach else None),
                    attach_hint_datetime=sent_date_hint,
                )
                if attach and attached:
                    attach_note = "\n\n📎 Found and attached the original invite email."
                elif attach:
                    attach_note = (
                        "\n\n⚠️ Couldn't find the original invite email in Sent Items to attach "
                        "(it may have been deleted/moved, or sent from a different account) — the "
                        "reminder email was still created, just without an attachment."
                    )
                else:
                    attach_note = ""
                action = "opened"
                review_note = " Review it, then click Send in Outlook."

                # Ghi lại thời điểm mở email nhắc nhở vào "LastReminderSentDate"
                # — giống cách _send_invite() ghi SentDate ngay lúc gửi/mở — để
                # Tab 7 (Event History) có dữ liệu để xem lại.
                history_log_note = ""
                try:
                    if db.update_event(
                            event_id,
                            {"LastReminderSentDate": datetime.now().strftime("%Y-%m-%d %H:%M")},
                            history_path):
                        self.after(0, self._refresh_history_tree)
                except Exception:
                    history_log_note = ("\n\n⚠️ Couldn't write the reminder-sent time to "
                                         "the database — "
                                         "the email was still " + action + " normally.")

                self.after(0, lambda: messagebox.showinfo(
                    "Done",
                    f"{action.capitalize()} a reminder email for {pending_count} people who haven't responded."
                    + review_note + attach_note + history_log_note))
            except Exception as e:
                err_msg = str(e)
                self.after(0, lambda: messagebox.showerror(
                    "Error", f"Couldn't send the reminder email:\n{err_msg}"))

        threading.Thread(target=worker, daemon=True).start()

    # ══════════════════════════════════════════════════════════════════
    # TAB 5 — SEND MEETING (CALENDAR INVITE) TO EVERYONE WHO VOTED "YES"
    # ══════════════════════════════════════════════════════════════════
    def _build_tab_gift(self):
        f = self.tab_gift.body
        gap = {"fill": "x", "pady": (0, 16)}

        self.var_gift_contributed_count = tk.StringVar(value="0 / 0")
        self.var_gift_total_amount = tk.StringVar(value="0")
        kpi_row(f, [
            ("Contributors", self.var_gift_contributed_count, {"dot": COLORS["success_fill"]}),
            ("Total collected", self.var_gift_total_amount, {}),
        ]).pack(**gap)

        tracking = Card(f, "Contributions",
                        "Everyone on Recipients. Tick who has contributed; it is saved as you click, and "
                        "comes back when you reopen the event.")
        tracking.pack(**gap)
        ttk.Button(tracking.actions, text="📊 Export to Excel", command=self._export_gift_contribution_list)\
            .pack(side="right")
        # Importing a file (e.g. one a colleague edited in Excel) OVERWRITES
        # the checked state and amount of EVERYONE found in it.
        ttk.Button(tracking.actions, text="📂 Load from file", command=self._load_gift_list_from_file)\
            .pack(side="right", padx=(0, 8))

        # Lọc TRỰC TIẾP bảng khi gõ; trạng thái tick vẫn nằm trong
        # self._gift_roster nên lọc không làm mất tick của người bị ẩn.
        toolbar = ttk.Frame(tracking.body)
        toolbar.pack(fill="x", pady=(0, 12))
        self.var_gift_search = tk.StringVar(value="")
        ttk.Label(toolbar, text="🔎", style="Muted.TLabel").pack(side="left", padx=(0, 6))
        ttk.Entry(toolbar, textvariable=self.var_gift_search, width=36).pack(side="left")
        self.var_gift_search.trace_add("write", lambda *a: self._apply_gift_filter())
        ttk.Button(toolbar, text="✕ Clear", style="Ghost.TButton",
                   command=lambda: self.var_gift_search.set("")).pack(side="left", padx=(6, 0))

        # Column order: send_email, check, No., Name, Email, Amount. "Send
        # email" (who gets the report) is independent of "Contributed". Click
        # a cell to toggle it, or a column header to toggle every row shown
        # (see _on_gift_tree_click() / _toggle_all_gift_column()).
        cols = ("send_email", "check", "no", "name", "email", "amount")
        tree_container, self.tree_gift = make_scrollable_treeview(tracking.body, columns=cols, height=14)
        self.tree_gift.heading("send_email", text="⬜ Send email",
                                command=lambda: self._toggle_all_gift_column("send_email"))
        self.tree_gift.heading("check", text="⬜ Contributed",
                                command=lambda: self._toggle_all_gift_column("check"))
        self.tree_gift.heading("no", text="No.")
        self.tree_gift.heading("name", text="Name", anchor="w")
        self.tree_gift.heading("email", text="Email", anchor="w")
        self.tree_gift.heading("amount", text="Amount", anchor="e")
        self.tree_gift.column("send_email", width=110, anchor="center", stretch=False)
        self.tree_gift.column("check", width=110, anchor="center", stretch=False)
        self.tree_gift.column("no", width=44, anchor="center", stretch=False)
        self.tree_gift.column("name", width=220)
        self.tree_gift.column("email", width=260)
        self.tree_gift.column("amount", width=110, anchor="e", stretch=False)
        tree_container.pack(fill="both", expand=True)
        self.tree_gift.bind("<Button-1>", self._on_gift_tree_click)
        # Double-click Amount to type what someone actually gave.
        self.tree_gift.bind("<Double-1>", self._on_gift_tree_double_click)
        WrapLabel(tracking.body, style="Hint.TLabel",
                  text="Ticking Contributed fills Amount from the expected gift budget; double-click an "
                       "Amount to type what that person actually gave - a typed amount (shown in blue) is "
                       "kept when Contributed is ticked again. Click a column header to tick or untick "
                       "everyone shown.").pack(fill="x", pady=(8, 0))
        self.tree_gift.tag_configure("manual_amt", foreground=COLORS["info"])

        # ── the gift actually bought, and where the money stands. Saved with
        # the event as you type (GiftItem* columns of its History row). ──
        item = Card(f, "Gift item & money",
                    "What was bought, and what is left. Optionally adds the party's money from "
                    "Attendance & payment for one Event + Gift summary - the same figures go into the "
                    "contribution report.")
        item.pack(**gap)
        self.var_gift_item_name = tk.StringVar(value="")
        self.var_gift_item_link = tk.StringVar(value="")
        self.var_gift_item_price = tk.StringVar(value="")
        fields = ttk.Frame(item.body, style="Card.TFrame")
        fields.pack(fill="x")
        fields.columnconfigure((0, 1), weight=1, uniform="gift")
        field(fields, 0, 0, "Gift name", lambda p: ttk.Entry(p, textvariable=self.var_gift_item_name))
        field(fields, 0, 1, "Gift price",
                           lambda p: ttk.Entry(p, textvariable=self.var_gift_item_price, width=16),
                           hint='e.g. "3,570 JPY" - only the number is used.')
        link_row = ttk.Frame(fields, style="Card.TFrame")
        link_row.grid(row=2, column=0, columnspan=2, sticky="we", pady=(4, 0))
        ttk.Label(link_row, text="Order link (URL)", style="Field.TLabel").pack(anchor="w", pady=(0, 4))
        link_entry_row = ttk.Frame(link_row, style="Card.TFrame")
        link_entry_row.pack(fill="x")
        ttk.Entry(link_entry_row, textvariable=self.var_gift_item_link).pack(side="left", fill="x", expand=True)
        ttk.Button(link_entry_row, text="🌐 Open link", command=self._open_gift_item_link)\
            .pack(side="left", padx=(8, 0))
        for var in (self.var_gift_item_name, self.var_gift_item_link, self.var_gift_item_price):
            var.trace_add("write", lambda *a: self._on_gift_item_changed())
        self._gift_fields_frame = fields
        self.lbl_gift_price_warning = WrapLabel(item.body, style="Danger.TLabel", text="")

        self.var_gift_cost_display = tk.StringVar(value="0")
        self.var_gift_remaining = tk.StringVar(value="0")
        kpi_row(item.body, [
            ("Collected (gift)", self.var_gift_total_amount, {}),
            ("Gift cost", self.var_gift_cost_display, {}),
            ("Gift remaining", self.var_gift_remaining, {}),
        ]).pack(fill="x", pady=(16, 0))

        self.var_gift_link_event = tk.BooleanVar(value=False)
        ttk.Checkbutton(item.body, text="🔗 Add the party's money from Attendance & payment "
                                        "(Event + Gift totals)",
                        variable=self.var_gift_link_event,
                        command=self._on_gift_link_event_toggled).pack(anchor="w", pady=(14, 4))
        self.lbl_gift_event_figures = WrapLabel(item.body, style="Hint.TLabel", text="")
        self.lbl_gift_event_figures.pack(fill="x")
        self.var_gift_grand_collected = tk.StringVar(value="0")
        self.var_gift_grand_paid = tk.StringVar(value="0")
        self.var_gift_grand_remaining = tk.StringVar(value="0")
        kpi_row(item.body, [
            ("Total collected", self.var_gift_grand_collected, {}),
            ("Total paid / spent", self.var_gift_grand_paid, {}),
            ("Total remaining", self.var_gift_grand_remaining, {}),
        ]).pack(fill="x", pady=(10, 0))

        # ── reminder to people who haven't contributed — always opens Outlook
        # for review (mail.Display()); there is no auto-send here. ──
        reminder = Card(f, "Reminder", "To everyone not ticked as Contributed. Opens in Outlook for you to "
                                       "check and send.")
        reminder.pack(**gap)
        grbtn = ttk.Frame(reminder.body)
        grbtn.pack(fill="x")
        ttk.Label(grbtn, text="Language", style="Field.TLabel").pack(side="left", padx=(0, 8))
        self.combo_gift_reminder_lang = ttk.Combobox(
            grbtn, width=26, state="readonly",
            values=[LANG_LABELS["en"], LANG_LABELS["ja"], LANG_LABELS["vi"], LANG_LABELS["bilingual"]],
        )
        self.combo_gift_reminder_lang.current(0)  # default: English
        self.combo_gift_reminder_lang.pack(side="left")
        self.combo_gift_reminder_lang.bind(
            "<<ComboboxSelected>>", lambda e: self._generate_gift_reminder_draft())
        ttk.Button(grbtn, text="🔄 Regenerate text", command=self._generate_gift_reminder_draft)\
            .pack(side="left", padx=(8, 0))
        self.var_gift_reminder_attach = tk.BooleanVar(value=True)
        ttk.Checkbutton(reminder.body, text="📎 Attach original Gift email (looked up from Sent Items)",
                        variable=self.var_gift_reminder_attach).pack(anchor="w", pady=(10, 8))
        gift_reminder_text_container, self.txt_gift_reminder_body = make_scrollable_text(
            reminder.body, width=40, height=7)
        gift_reminder_text_container.pack(fill="x")
        self.lbl_gift_reminder_send = ttk.Button(
            reminder.body, text="📨 Open reminder email for 0 people who haven't contributed",
            style="Primary.TButton", command=self._send_gift_reminder)
        self.lbl_gift_reminder_send.pack(anchor="w", pady=(12, 0))

        # ── contribution report to the people ticked in "Send email": a short
        # summary in the body, the list of contributors attached as Excel
        # (see reports.gift_report_workbook()/build_gift_report_body()). ──
        report = Card(f, "Contribution report",
                      "To everyone ticked in Send email: the gift, how many contributed and the money "
                      "(with the party's when linked above), and the list of contributors attached as an "
                      "Excel file. Opens in Outlook first.")
        report.pack(fill="x")
        rbtn = ttk.Frame(report.body)
        rbtn.pack(fill="x", pady=(0, 10))
        ttk.Label(rbtn, text="Language", style="Field.TLabel").pack(side="left", padx=(0, 8))
        self.combo_gift_report_lang = ttk.Combobox(
            rbtn, width=26, state="readonly",
            values=[LANG_LABELS["en"], LANG_LABELS["ja"], LANG_LABELS["vi"], LANG_LABELS["bilingual"]],
        )
        self.combo_gift_report_lang.current(0)  # default: English
        self.combo_gift_report_lang.pack(side="left")
        self.combo_gift_report_lang.bind("<<ComboboxSelected>>", lambda e: self._generate_gift_report_draft())
        ttk.Button(rbtn, text="🔄 Regenerate text", command=self._generate_gift_report_draft)\
            .pack(side="left", padx=(8, 0))
        report_text_container, self.txt_gift_report_body = make_scrollable_text(
            report.body, width=40, height=10)
        report_text_container.pack(fill="x")
        self.lbl_gift_report_send = ttk.Button(
            report.body, text="📧 Send report to 0 selected people", style="Primary.TButton",
            command=self._send_gift_report_email)
        self.lbl_gift_report_send.pack(anchor="w", pady=(12, 0))

    def _pending_gift_contributors(self):
        """Returns list[(name, email)] of everyone NOT yet checked as
        "Contributed" in self._gift_roster (the source of truth for Tab 6 —
        independent of whatever the search box is currently filtering)."""
        roster = getattr(self, "_gift_roster", None) or {}
        return [(info["name"], email) for email, info in roster.items() if not info["checked"]]

    def _update_gift_reminder_button_label(self):
        if not hasattr(self, "lbl_gift_reminder_send"):
            return
        count = len(self._pending_gift_contributors())
        self.lbl_gift_reminder_send.config(text=f"📨 Open reminder email for {count} people who haven't contributed")

    def _generate_gift_reminder_draft(self):
        """Auto-composes a draft gift-contribution reminder email using the
        event info currently on Tab 1 (Guest of Honor/Location/Start time/
        Event date/Organizer/Gift deadline/Gift budget) plus the language
        selected in combo_gift_reminder_lang. Only produces a DRAFT — nothing
        is sent — the user reviews/edits it in the box below before sending."""
        guest_of_honor = self.var_guest_of_honor.get()
        start_time = self.var_start_time.get()
        event_date = get_date_str(self.date_event)
        location = self.var_location.get()
        organizer = self.var_organizer.get()
        deadline = get_date_str(self.date_gift_deadline)
        gift_budget = self.var_gift_budget.get()

        lang_label = self.combo_gift_reminder_lang.get() or LANG_LABELS["en"]
        lang_code = LANG_LABEL_TO_CODE.get(lang_label, "en")

        draft = build_gift_reminder_body(lang_code, guest_of_honor, start_time, event_date,
                                          location, organizer, deadline, gift_budget)
        self._generated_drafts["gift_reminder"] = draft
        self.txt_gift_reminder_body.delete("1.0", "end")
        self.txt_gift_reminder_body.insert("1.0", draft)

    def _send_gift_reminder(self):
        pending = self._pending_gift_contributors()
        if not pending:
            messagebox.showinfo(
                "No one left to remind",
                "Everyone in the Tab 6 list is already checked as 'Contributed' — no need to send a reminder."
            )
            return

        body = self.txt_gift_reminder_body.get("1.0", "end").strip()
        if not body:
            messagebox.showwarning(
                "Empty content",
                "Click '🔄 Regenerate text' or type the content by hand before sending.")
            return

        event_id = self.var_event_id.get().strip()
        guest_of_honor = self.var_guest_of_honor.get()
        lang_label = self.combo_gift_reminder_lang.get() or LANG_LABELS["en"]
        lang_code = LANG_LABEL_TO_CODE.get(lang_label, "en")
        # Uses its OWN subject prefix "[Reminder-Gift-{event_id}]" — doesn't
        # need to match any Scan Inbox logic (Gift mode has no Voting
        # Buttons / isn't vote-scanned), just needs to be clearly distinct
        # from the original Gift email.
        subject = build_gift_reminder_subject(lang_code, event_id, guest_of_honor)

        attach = self.var_gift_reminder_attach.get()
        pending_count = len(pending)

        def worker():
            try:
                # ALWAYS auto_send=False — only opens the Outlook window for
                # review, there is no direct-send option (unlike Tab 4's
                # RSVP reminder, which used to have a "Send now" checkbox;
                # here it's deliberately absent, per the requirement that
                # every reminder — RSVP or Gift — must be sent by the user
                # clicking Send themselves).
                mail, attached = self.outlook.send_gift_reminder_email(
                    pending, subject, body,
                    auto_send=False,
                    attach_event_id=(event_id if attach else None),
                )
                if attach and attached:
                    attach_note = "\n\n📎 Found and attached the original Gift email."
                elif attach:
                    attach_note = (
                        "\n\n⚠️ Couldn't find the original Gift email in Sent Items to attach "
                        "(it may have been deleted/moved, or sent from a different account) — the "
                        "reminder email was still created, just without an attachment."
                    )
                else:
                    attach_note = ""

                self.after(0, lambda: messagebox.showinfo(
                    "Reminder email opened",
                    f"Opened a reminder email compose window for {pending_count} people who haven't contributed.\n"
                    "Review it, then click Send in Outlook." + attach_note))
            except Exception as e:
                err_msg = str(e)
                self.after(0, lambda: messagebox.showerror(
                    "Error", f"Couldn't open the reminder email:\n{err_msg}"))

        threading.Thread(target=worker, daemon=True).start()

    def _load_gift_list_from_file(self):
        """Lets you pick ANY Excel file by hand (typically an older
        'Gift_Contribution_List_*.xlsx', e.g. exported before switching to
        the database, or a copy edited by a colleague outside the app),
        then MERGES/UPDATES the contribution state into self._gift_roster
        and immediately saves it to the database:
        - Accepts ANY .xlsx file (doesn't have to match the current Event
          ID's standard export filename).
        - OVERWRITES the checked state (and amount) for EVERYONE found in
          the chosen file (not just people who aren't already in the
          current list, like the normal Tab-open reload does).
        - Automatically adds NEW people to the list if the file contains
          someone not currently in Tab 2 (e.g. the Tab 2 list changed since
          the file was exported)."""
        path = filedialog.askopenfilename(
            title="Select a Gift Contribution List file",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not path:
            return
        if not hasattr(self, "_gift_roster"):
            self._gift_roster = {}
        try:
            rows = read_gift_contribution_rows(path)
        except Exception as e:
            messagebox.showerror("Error", f"Couldn't read the file:\n{e}")
            return

        for row in rows:
            existing = self._gift_roster.get(row["email"], {})
            self._gift_roster[row["email"]] = {
                "name": row["name"] or existing.get("name", row["email"]),
                "checked": row["checked"],
                "amount": row["amount"],
                # What the file says someone gave: ticking Contributed later
                # must not replace it with the expected budget.
                "manual_amount": row["amount"] > 0,
                # "send_email" không có trong file Excel (chỉ là lựa chọn
                # riêng của app, xem cột "Send email") — GIỮ NGUYÊN giá trị
                # đã có (nếu người này đã từng được tick chọn nhận báo cáo
                # trước đó), không để việc import file vô tình xoá mất.
                "send_email": existing.get("send_email", False),
            }

        self._apply_gift_filter()
        # Lưu ngay vào database (thay vì ghi lại Excel như kiến trúc cũ) —
        # để lần mở lại Tab 6 sau này tự nạp đúng những gì vừa import.
        self._save_gift_roster_to_db(silent=True)
        messagebox.showinfo(
            "Loaded",
            f"Loaded/updated {len(rows)} people from file:\n{path}\n\n"
            f"Total in list now: {len(self._gift_roster)} people."
        )

    # ── Tab 6: the gift item and the money ──

    def _adopt_gift_item(self, event_id):
        """Shows event_id's saved gift item (or none), without saving it back.
        When it cannot be read, the boxes are shown empty and belong to no
        event, so typing in them saves nothing over the stored item."""
        if not hasattr(self, "var_gift_item_name"):
            return
        rec = {}
        if event_id:
            try:
                rec = self._history_record(event_id) or {}
            except Exception as exc:
                messagebox.showerror(
                    "Gift item not loaded",
                    f"Couldn't read the gift item of '{event_id}':\n\n{exc}\n\n"
                    "It is shown empty and will not be saved until it can be read.")
                rec, event_id = {}, None
        self._gift_item_event = event_id or None
        self._gift_item_quiet = True
        try:
            self.var_gift_item_name.set(rec.get("GiftItemName") or "")
            self.var_gift_item_link.set(rec.get("GiftItemLink") or "")
            self.var_gift_item_price.set(rec.get("GiftItemPrice") or "")
            self.var_gift_link_event.set((rec.get("GiftLinkEvent") or "").strip().lower() == "yes")
        finally:
            self._gift_item_quiet = False
        self._update_gift_summary()

    def _gift_item_fields(self):
        return {"GiftItemName": self.var_gift_item_name.get(),
                "GiftItemLink": self.var_gift_item_link.get(),
                "GiftItemPrice": self.var_gift_item_price.get(),
                "GiftLinkEvent": "Yes" if self.var_gift_link_event.get() else "No"}

    def _on_gift_item_changed(self):
        """Typing in the gift item: figures follow, and it is saved to the
        event the gift list belongs to (UPDATE only - see Tab 1's save)."""
        if self._gift_item_quiet:
            return
        self._update_gift_summary()
        event_id = self._gift_item_event
        if not event_id:
            return
        was_tracked = self._was_money_tracked(event_id)
        if self._save_event_fields(event_id, self._gift_item_fields(), "The gift item"):
            self._sync_event_money(event_id, was_tracked=was_tracked)

    def _on_gift_link_event_toggled(self):
        self._on_gift_item_changed()

    def _open_gift_item_link(self):
        """Opens the order link in the browser - http(s) only, so text pasted
        into the box by mistake cannot open a file or run anything."""
        url = (self.var_gift_item_link.get() or "").strip()
        if not url:
            messagebox.showinfo("No link", "Type the order link (URL) in the box first.")
            return
        if not re.match(r"^https?://", url, re.I):
            messagebox.showwarning("Not a web link",
                                   f"Only http:// or https:// links can be opened.\n\nCurrent value: {url}")
            return
        try:
            webbrowser.open(url)
        except Exception as e:
            messagebox.showerror("Error", f"Couldn't open the link:\n{e}")

    def _event_figures_for(self, event_id):
        """Every round's figures of event_id: from the screen when Tab 5 holds
        that event, else from the database (which may raise)."""
        if event_id and event_id == self._attendance_event == self._amount_paid_event:
            return self._attendance_figures()
        if not event_id:
            return payment_rounds([], ROUND1_DEFAULT_LABEL, "", [])
        path = self.history_path.get()
        rec = self._history_record(event_id) or {}
        return payment_rounds(db.load_attendance_roster(event_id, path).values(),
                              rec.get("Round1Label") or ROUND1_DEFAULT_LABEL, rec.get("AmountPaid"),
                              db.load_attendance_rounds(event_id, path))

    def _gift_money(self):
        """The gift's money, plus - when linked - the party's, for the event
        the gift list belongs to. One source for the screen and the report.
        "unreadable" holds the error when the party's figures could not be
        read: zeros in their place would report the party as costing nothing."""
        linked = bool(self.var_gift_link_event.get())
        event_totals, unreadable = (0.0, 0.0, 0.0), None
        if linked:
            try:
                event_totals = round_totals(self._event_figures_for(self._gift_event))
            except Exception as exc:
                unreadable = exc
        fig = gift_figures(contributed_total(self._gift_roster.values()),
                           self.var_gift_item_price.get(), linked, event_totals)
        fig["unreadable"] = unreadable
        return fig

    @staticmethod
    def _gift_amount_text(fig, key):
        """An amount for the screen or the report; "?" for one that includes
        the party's figures when those could not be read."""
        if fig["unreadable"] is not None and (key.startswith("event_") or key.startswith("grand_")):
            return "?"
        return format_amount(fig[key])

    def _update_gift_summary(self):
        if not hasattr(self, "var_gift_grand_remaining"):
            return  # Tab 6 is still being built
        fig = self._gift_money()
        price = self.var_gift_item_price.get()
        self._show_typed_amount_warning(
            self.lbl_gift_price_warning,
            [f"Gift price: “{price}” is read as {format_amount(fig['gift_cost'])}"]
            if unclear_typed_amount(price) else [], self._gift_fields_frame)
        self.var_gift_cost_display.set(format_amount(fig["gift_cost"]))
        self.var_gift_remaining.set(format_amount(fig["gift_remaining"]))
        self.var_gift_grand_collected.set(self._gift_amount_text(fig, "grand_collected"))
        self.var_gift_grand_paid.set(self._gift_amount_text(fig, "grand_paid"))
        self.var_gift_grand_remaining.set(self._gift_amount_text(fig, "grand_remaining"))
        if fig["unreadable"] is not None:
            text = (f"⚠ Attendance & payment of {self._gift_event} could not be read "
                    f"({fig['unreadable']}) - the totals below are unknown. Untick this, or try again.")
        elif fig["linked"]:
            text = (f"From Attendance & payment - collected {format_amount(fig['event_collected'])}, "
                    f"paid out {format_amount(fig['event_paid'])}, remaining {format_amount(fig['event_remaining'])}.")
        else:
            text = "Not linked - the totals below cover the gift only."
        self.lbl_gift_event_figures.configure(text=text)
        # The report quotes these figures: rebuild it unless edited by hand.
        current = self.txt_gift_report_body.get("1.0", "end").strip() if hasattr(
            self, "txt_gift_report_body") else None
        if current is not None and current == (self._generated_drafts.get("gift_report") or "").strip() \
                and current:
            self._generate_gift_report_draft()

    def _gift_report_body_args(self):
        """Everything build_gift_report_body() quotes: the gift, how many
        contributed, the gift's money and - when linked - a Party / Gift table
        with the Event + Gift totals."""
        contributor_count = sum(1 for info in self._gift_roster.values() if info.get("checked"))
        fig = self._gift_money()
        lang_label = self.combo_gift_report_lang.get() or LANG_LABELS["en"]
        lang_code = LANG_LABEL_TO_CODE.get(lang_label, "en")
        L = GIFT_REPORT_LABELS.get("en" if lang_code == "bilingual" else lang_code, GIFT_REPORT_LABELS["en"])
        rows = []
        if fig["linked"]:
            rows.append((L["row_event"], self._gift_amount_text(fig, "event_collected"),
                         self._gift_amount_text(fig, "event_paid"),
                         self._gift_amount_text(fig, "event_remaining")))
        rows.append((L["row_gift"], format_amount(fig["gift_collected"]),
                     format_amount(fig["gift_cost"]), format_amount(fig["gift_remaining"])))
        return (
            self.var_guest_of_honor.get(),
            self.var_event_name.get(),
            str(contributor_count),
            self.var_gift_total_amount.get(),
            self.var_gift_item_name.get(),
            self.var_gift_item_link.get(),
            format_amount(fig["gift_cost"]) if self.var_gift_item_price.get().strip() else "",
            rows,
            self._gift_amount_text(fig, "grand_collected"),
            self._gift_amount_text(fig, "grand_paid"),
            self._gift_amount_text(fig, "grand_remaining"),
            fig["linked"],
            format_amount(fig["gift_remaining"]),
        )

    def _generate_gift_report_draft(self):
        """Auto-composes a draft contribution-report email using the event
        info on Tab 1 plus the CURRENT Gift Contribution table (Tab 6) and
        the language selected in combo_gift_report_lang. Only produces a
        DRAFT — nothing is sent — the user reviews/edits it in the box
        below before sending."""
        lang_label = self.combo_gift_report_lang.get() or LANG_LABELS["en"]
        lang_code = LANG_LABEL_TO_CODE.get(lang_label, "en")
        draft = build_gift_report_body(lang_code, *self._gift_report_body_args())
        self._generated_drafts["gift_report"] = draft
        self.txt_gift_report_body.delete("1.0", "end")
        self.txt_gift_report_body.insert("1.0", draft)

    def _send_gift_report_email(self):
        """MỚI: gửi email báo cáo số tiền đã quyên góp được tới những ai đã
        tick chọn ở cột "Send email" (self._gift_roster[...]["send_email"])
        — HOÀN TOÀN ĐỘC LẬP với cột "Contributed" (người nhận báo cáo
        không nhất thiết phải là người đã đóng góp). Nội dung mail chỉ là
        1 thông báo TỔNG QUAN (đã thu được bao nhiêu người/bao nhiêu tiền)
        — danh sách chi tiết từng người ĐÃ đóng góp (đánh số lại từ 1,
        KHÔNG gồm 2 cột checkbox) nằm trong file Excel TỰ ĐỘNG ĐÍNH KÈM
        (xem reports.gift_report_workbook()). Không có Voting Buttons (dùng
        self.outlook.send_gift_report_email(), hàm gửi thông báo thuần tuý
        + đính kèm file — KHÁC với send_voting_invite() vốn không hỗ trợ
        đính kèm)."""
        recipients = [
            (info.get("name") or email, email)
            for email, info in getattr(self, "_gift_roster", {}).items()
            if info.get("send_email")
        ]
        if not recipients:
            messagebox.showwarning(
                "No recipients selected",
                "Tick the \"Send email\" checkbox (top-left column) for at least one person first — "
                "click the ⬜/✅ mark in the column header to select everyone currently shown at once."
            )
            return

        unreadable = self._gift_money()["unreadable"]
        if unreadable is not None:
            messagebox.showwarning(
                "Party figures unknown",
                f"Attendance & payment of {self._gift_event} could not be read, so the report's "
                f"Event + Gift totals would be wrong:\n\n{unreadable}\n\n"
                "Untick \"Add the party's money from Attendance & payment\" to report the gift "
                "alone, or try again.")
            return

        body = self.txt_gift_report_body.get("1.0", "end").strip()
        # Generated text is rebuilt from the current figures first; text
        # edited by hand is sent as it is.
        if not body or body == (self._generated_drafts.get("gift_report") or "").strip():
            self._generate_gift_report_draft()
            body = self.txt_gift_report_body.get("1.0", "end").strip()
        if not body:
            messagebox.showwarning(
                "Empty content",
                "Click '🔄 Regenerate text' or type the content by hand before sending.")
            return

        event_id = self.var_event_id.get().strip()
        guest_of_honor = self.var_guest_of_honor.get()
        lang_label = self.combo_gift_report_lang.get() or LANG_LABELS["en"]
        lang_code = LANG_LABEL_TO_CODE.get(lang_label, "en")
        subject = build_gift_report_subject(lang_code, event_id, guest_of_honor)

        # Chuẩn bị file Excel đính kèm TRƯỚC (trong main thread, vì cần đọc
        # self._gift_roster/các StringVar UI — an toàn hơn đọc chúng từ
        # worker thread), lưu vào 1 file tạm — không phải nơi lưu trữ
        # chính (dữ liệu thật vẫn ở database), chỉ để đính kèm email.
        try:
            wb = reports.gift_report_workbook(self._gift_roster)
            excel_path = os.path.join(
                tempfile.gettempdir(), f"Gift_Contribution_Report_{event_id or 'event'}.xlsx")
            wb.save(excel_path)
        except Exception as e:
            messagebox.showerror("Error", f"Couldn't prepare the contribution report attachment:\n{e}")
            return

        recipient_count = len(recipients)
        # HTML, so the Party / Gift table lines up and the order link opens.
        html_body = text_body_to_html(body)

        def worker():
            try:
                mail, attached = self.outlook.send_gift_report_email(
                    recipients, subject, body, excel_path=excel_path, html_body=html_body)
                attach_note = ("\n\n📎 Attached the contribution report." if attached else
                                "\n\n⚠️ Couldn't attach the contribution report file — the email was "
                                "still created without it.")
                self.after(0, lambda: messagebox.showinfo(
                    "Done",
                    f"Contribution report email opened for {recipient_count} selected people. "
                    "Review it, then click Send in Outlook." + attach_note))
            except Exception as e:
                err_msg = str(e)
                self.after(0, lambda: messagebox.showerror(
                    "Error", f"Could not create the report email:\n{err_msg}"))

        threading.Thread(target=worker, daemon=True).start()

    def _refresh_gift_contribution_list(self):
        """Reloads the Name/Email list from Tab 2 (self.recipients) into
        self._gift_roster — the SOURCE OF TRUTH (not the Treeview itself —
        the Treeview is only a filterable VIEW, see _apply_gift_filter()).
        Priority order for each person's checked state and amount:
        1. Keep whatever is already in self._gift_roster (if it was already
           loaded this session — e.g. a new person just appeared on Tab 2),
           otherwise
        2. If that person has never been in self._gift_roster yet, try
           reading from the database (bảng gift_contributions) — to RESTORE
           the checked state (and amount) from last time, even after
           closing/reopening the app."""
        event_id = self.var_event_id.get().strip()
        if self._gift_event != (event_id or None):
            # The roster in memory belongs to another event: start from what
            # is saved for this one, never from the other event's ticks - or
            # its gift item.
            self._gift_roster = {}  # email -> {"name", "checked", "amount", "send_email", "manual_amount"}
            self._gift_event = event_id or None
            self._adopt_gift_item(self._gift_event)

        # Step 1: read the saved state from the database as a fallback —
        # ONLY used for people not already in self._gift_roster.
        saved_state = {}
        if event_id:
            try:
                saved_state = db.load_gift_roster(event_id, self.history_path.get())
            except Exception:
                pass  # best-effort

        new_roster = {}
        for name, email in self.recipients:
            prev = self._gift_roster.get(email) or saved_state.get(email, {})
            new_roster[email] = {"name": name, "checked": prev.get("checked", False),
                                 "amount": prev.get("amount", 0.0),
                                 "send_email": prev.get("send_email", False),
                                 "manual_amount": prev.get("manual_amount", False)}
        self._gift_roster = new_roster
        self._apply_gift_filter()

    def _apply_gift_filter(self, *args):
        """Redraws the Treeview from self._gift_roster, only showing rows
        that match the search box (if any) — this does NOT touch the actual
        checked state (which lives in self._gift_roster), so searching/
        filtering never loses anyone's checked state. The "No." column shows
        each visible row's position (1, 2, 3, ...) in the CURRENTLY SHOWN
        list, so it stays contiguous even while filtering."""
        if not hasattr(self, "tree_gift") or not hasattr(self, "_gift_roster"):
            return
        query = self.var_gift_search.get().strip().lower() if hasattr(self, "var_gift_search") else ""
        self.tree_gift.delete(*self.tree_gift.get_children())
        seq = 0
        for email, info in self._gift_roster.items():
            name = info["name"]
            if query and query not in (name or "").lower() and query not in email.lower():
                continue
            seq += 1
            send_email_display = "✅" if info.get("send_email") else "⬜"
            check = "✅" if info["checked"] else "⬜"
            amount_display = f"{info.get('amount', 0.0):,.0f}" if info.get("amount") else ""
            # Use EMAIL as the Treeview row's iid -> look up/toggle the
            # right person in self._gift_roster directly, without matching
            # displayed Name/Email strings (which could collide).
            self.tree_gift.insert("", "end", iid=email,
                                   tags=("manual_amt",) if info.get("manual_amount") else (),
                                   values=(send_email_display, check, seq, name, email, amount_display))
        self._update_gift_contributed_count()
        self._update_gift_header_checkmarks()

    def _update_gift_header_checkmarks(self):
        """Cập nhật dấu ✅/⬜ hiển thị NGAY TRÊN TIÊU ĐỀ cột "Send email"/
        "Contributed" — ✅ khi TẤT CẢ dòng ĐANG HIỂN THỊ (tôn trọng ô tìm
        kiếm) đều đã tick ở cột đó, ⬜ nếu không (kể cả khi danh sách đang
        hiển thị rỗng). Gọi lại sau mỗi lần vẽ bảng (_apply_gift_filter())
        để tiêu đề luôn phản ánh đúng trạng thái hiện tại."""
        if not hasattr(self, "tree_gift"):
            return
        shown = self.tree_gift.get_children()
        all_send = bool(shown) and all(
            self._gift_roster.get(iid, {}).get("send_email") for iid in shown)
        all_checked = bool(shown) and all(
            self._gift_roster.get(iid, {}).get("checked") for iid in shown)
        self.tree_gift.heading("send_email", text=("✅" if all_send else "⬜") + " Send email")
        self.tree_gift.heading("check", text=("✅" if all_checked else "⬜") + " Contributed")

    def _on_gift_tree_click(self, event):
        region = self.tree_gift.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.tree_gift.identify_column(event.x)
        iid = self.tree_gift.identify_row(event.y)
        if not iid:
            return
        email = iid  # iid IS the email — see _apply_gift_filter()
        if email not in self._gift_roster:
            return
        if col == "#1":  # "send_email" — cột checkbox chọn người nhận báo cáo, ĐỘC LẬP với "checked"
            self._gift_roster[email]["send_email"] = not self._gift_roster[email].get("send_email", False)
        elif col == "#2":  # "check" — cột "Contributed"
            new_checked = not self._gift_roster[email]["checked"]
            self._gift_roster[email]["checked"] = new_checked
            # Auto-fill the Amount from Tab 1's "Expected gift budget" the
            # moment someone is checked as contributed; clear it back to 0 if
            # unchecked, so the total only ever counts people currently marked
            # as having contributed. A different actual amount is typed by
            # double-clicking the Amount cell (_on_gift_tree_double_click()).
            # A typed amount (_commit_gift_amount_edit()) is never replaced
            # by the budget; unticking clears it and forgets that it was typed.
            if new_checked:
                if not self._gift_roster[email].get("manual_amount"):
                    self._gift_roster[email]["amount"] = parse_amount_from_text(self.var_gift_budget.get())
            else:
                self._gift_roster[email]["amount"] = 0.0
                self._gift_roster[email]["manual_amount"] = False
        else:
            return  # cột khác (No./Name/Email/Amount) không tick được bằng click
        self._apply_gift_filter()
        self.tree_gift.see(iid)
        # Auto-save to the database right after every check/uncheck — no
        # separate button needed, so progress isn't lost if you forget to
        # save manually before closing the app. Saves silently (no
        # messagebox) so it doesn't interrupt every single click.
        self._save_gift_roster_to_db(silent=True)

    def _on_gift_tree_double_click(self, event):
        if self.tree_gift.identify("region", event.x, event.y) != "cell":
            return
        iid = self.tree_gift.identify_row(event.y)
        if iid and self.tree_gift.identify_column(event.x) == "#6":  # "amount"
            self._begin_cell_edit(self.tree_gift, iid, "amount", self._commit_gift_amount_edit)

    def _commit_gift_amount_edit(self, row_id, col_name, new_value):
        """A typed Amount on Tab 6. Empty or 0 resets it (and forgets that it
        was typed). A positive amount is kept as typed - ticking Contributed
        again will not replace it with the budget - and marks the person as
        Contributed, since the total counts every amount. Text that is not
        one number ("abc", "1 000"), or a negative amount, is refused and the
        old value kept: reading it as 0 or 1 would silently change what that
        person gave."""
        info = self._gift_roster.get(row_id)
        if info is None:
            return
        text = (new_value or "").strip()
        if not text:
            amount = 0.0
        elif unclear_typed_amount(text):
            messagebox.showwarning(
                "Not a number",
                f"“{text}” is not one number, so the amount was left unchanged.\n\n"
                "Type just the figure (e.g. 1000 or 1,000), or clear the cell to reset it to 0.")
            return
        elif parse_typed_amount(text) < 0:
            messagebox.showwarning(
                "Negative amount", "A contribution amount can't be negative - the amount was left unchanged.")
            return
        else:
            amount = parse_amount_from_text(text)
        info["amount"] = amount
        info["manual_amount"] = amount > 0
        if amount > 0:
            info["checked"] = True
        self._apply_gift_filter()
        if self.tree_gift.exists(row_id):
            self.tree_gift.see(row_id)
        self._save_gift_roster_to_db(silent=True)

    def _toggle_all_gift_column(self, col_key):
        """MỚI: thay thế 2 nút "☑ Check all"/"☐ Uncheck all" cũ — bấm vào
        TIÊU ĐỀ cột "send_email" hoặc "check" (xem tree.heading(...,
        command=...) trong _build_tab_gift()) để chọn/bỏ chọn TẤT CẢ các
        dòng ĐANG HIỂN THỊ (tôn trọng ô tìm kiếm, giống hệt hành vi 2 nút
        cũ) cho ĐÚNG cột đó — 2 cột hoạt động HOÀN TOÀN ĐỘC LẬP với nhau.
        Trạng thái MỚI luôn là ĐẢO NGƯỢC của trạng thái hiện tại: nếu tất
        cả dòng đang hiển thị đã tick hết -> bỏ tick hết; ngược lại (kể cả
        khi chỉ tick 1 phần) -> tick hết."""
        if col_key not in ("send_email", "check"):
            return
        shown = self.tree_gift.get_children()
        if not shown:
            return
        currently_all = all(self._gift_roster.get(iid, {}).get(
            "checked" if col_key == "check" else "send_email") for iid in shown)
        new_state = not currently_all
        for iid in shown:
            if iid not in self._gift_roster:
                continue
            if col_key == "check":
                if bool(self._gift_roster[iid].get("checked")) == new_state:
                    continue  # already ticked: its amount stays as it is
                self._gift_roster[iid]["checked"] = new_state
                # The same rule as a single tick: one header click must not
                # wipe out every amount typed by hand.
                if new_state:
                    if not self._gift_roster[iid].get("manual_amount"):
                        self._gift_roster[iid]["amount"] = parse_amount_from_text(
                            self.var_gift_budget.get())
                else:
                    self._gift_roster[iid]["amount"] = 0.0
                    self._gift_roster[iid]["manual_amount"] = False
            else:
                self._gift_roster[iid]["send_email"] = new_state
        self._apply_gift_filter()
        self._save_gift_roster_to_db(silent=True)

    def _update_gift_contributed_count(self):
        total = len(self._gift_roster) if hasattr(self, "_gift_roster") else 0
        contributed = sum(1 for info in self._gift_roster.values() if info["checked"]) if hasattr(self, "_gift_roster") else 0
        total_amount = contributed_total(self._gift_roster.values()) if hasattr(self, "_gift_roster") else 0.0
        shown = len(self.tree_gift.get_children())
        if shown != total:
            self.var_gift_contributed_count.set(f"{contributed} / {total}  (showing {shown}/{total} due to search)")
        else:
            self.var_gift_contributed_count.set(f"{contributed} / {total}")
        if hasattr(self, "var_gift_total_amount"):
            self.var_gift_total_amount.set(f"{total_amount:,.0f}")
        self._update_gift_reminder_button_label()
        self._update_gift_report_button_label()
        self._update_gift_summary()

    def _update_gift_report_button_label(self):
        """Cập nhật số người trên nút "📧 Send report to N selected people"
        — đếm theo cột "Send email" (self._gift_roster[...]["send_email"]),
        KHÔNG phải cột "Contributed" — 2 khái niệm độc lập, xem
        _build_tab_gift()."""
        if not hasattr(self, "lbl_gift_report_send"):
            return
        count = sum(1 for info in getattr(self, "_gift_roster", {}).values() if info.get("send_email"))
        self.lbl_gift_report_send.config(text=f"📧 Send report to {count} selected people")

    def _save_gift_roster_to_db(self, silent=True):
        """Lưu self._gift_roster vào database (bảng gift_contributions) —
        gọi TỰ ĐỘNG sau mọi tick/bỏ tick, "Check all"/"Uncheck all", hoặc
        import từ file. Đây là hàm PERSIST THẬT SỰ trong kiến trúc mới
        (thay cho việc ghi Excel liên tục trước đây). Writes to the event the
        roster belongs to, not whatever Tab 1 shows."""
        event_id = self._gift_event
        if not event_id or not self._gift_roster:
            return
        was_tracked = self._was_money_tracked(event_id)
        try:
            db.save_gift_roster(event_id, self._gift_roster, self.history_path.get())
        except Exception as exc:
            if not silent:
                raise
            # Money: a lost save is reported (once per run), not swallowed.
            self._report_save_failure("gift", "Gift contribution", self._unwritable_reason(exc))
            return
        self._save_failures_reported.discard("gift")
        self._sync_event_money(event_id, was_tracked=was_tracked)

    def _load_gift_roster_from_db(self, event_id):
        """Đọc gift roster đã lưu trong database cho event_id, dùng bởi
        '⬅ Load setup from selected event' ở Tab 1. Trả về True nếu có dữ
        liệu để nạp."""
        try:
            roster = db.load_gift_roster(event_id, self.history_path.get())
        except Exception:
            roster = {}
        self._gift_roster = roster
        self._gift_event = event_id
        self._adopt_gift_item(event_id)
        self._apply_gift_filter()
        return bool(roster)

    def _export_gift_contribution_list(self, silent=False):
        """Exports the FULL self._gift_roster (not just rows currently shown
        due to search) to Excel — CHỈ khi bấm nút "📊 Export to Excel"
        (không còn tự động ghi liên tục — dữ liệu thật sự sống trong
        database, xem _save_gift_roster_to_db()). Writes EVERYONE, including
        people who haven't contributed yet (the 'Contributed' column =
        Yes/No), so it's easy to compare total invited vs. total
        contributed."""
        event_id = self.var_event_id.get().strip()
        if not event_id:
            if not silent:
                messagebox.showwarning("Missing Event ID", "Enter an Event ID on Tab 1 before exporting the file.")
            return
        if not getattr(self, "_gift_roster", None):
            if not silent:
                messagebox.showwarning("List is empty",
                                        "No one in the list yet — add recipients on Tab 2 first.")
            return

        out_path = filedialog.asksaveasfilename(
            title="Export Gift Contribution List to Excel",
            defaultextension=".xlsx",
            initialfile=f"Gift_Contribution_List_{event_id}.xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not out_path:
            return

        try:
            reports.gift_contribution_workbook(self._gift_roster).save(out_path)
        except Exception as e:
            if not silent:
                messagebox.showerror("Error", f"Couldn't export the file:\n{e}")
            return
        contributed_count = sum(1 for info in self._gift_roster.values() if info["checked"])
        total_amount = sum(info.get("amount", 0.0) for info in self._gift_roster.values() if info["checked"])

        if not silent:
            total = len(self._gift_roster)
            messagebox.showinfo(
                "File exported",
                f"Saved the contribution list ({contributed_count}/{total} people contributed, "
                f"total collected: {total_amount:,.0f}) to:\n{out_path}"
            )

    def _build_tab_calendar(self):
        f = self.tab_calendar.body
        gap = {"fill": "x", "pady": (0, 16)}

        # Yes AND Maybe get the Calendar Invite: a Maybe should still have
        # the event in Outlook in case they decide to come.
        calendar = Card(f, "Calendar invite",
                        "For everyone who voted Yes or Maybe on Collect responses. Date, time and place "
                        "come from Event setup; the original invite email is attached.")
        calendar.pack(**gap)
        ttk.Label(calendar.body, text="Invitees", style="Field.TLabel").pack(anchor="w", pady=(0, 4))
        yes_box = bordered(calendar.body)
        yes_box.pack(fill="x")
        self.list_yes = tk.Listbox(yes_box, height=7)
        yes_vscroll = ttk.Scrollbar(yes_box, orient="vertical", command=self.list_yes.yview)
        self.list_yes.configure(yscrollcommand=autohide(yes_vscroll))
        self.list_yes.grid(row=0, column=0, sticky="nsew", padx=(8, 0), pady=6)
        yes_vscroll.grid(row=0, column=1, sticky="ns")
        yes_box.columnconfigure(0, weight=1)
        # Kept in step with Tab 4 by _refresh_response_tree(); no refresh needed.

        appt_header = ttk.Frame(calendar.body)
        appt_header.pack(fill="x", pady=(14, 4))
        ttk.Label(appt_header, text="Appointment text", style="Field.TLabel").pack(side="left")
        self.combo_calendar_lang = ttk.Combobox(
            appt_header, width=26, state="readonly",
            values=[LANG_LABELS["en"], LANG_LABELS["ja"], LANG_LABELS["vi"], LANG_LABELS["bilingual"]],
        )
        self.combo_calendar_lang.current(0)  # default: English
        self.combo_calendar_lang.pack(side="right")
        ttk.Label(appt_header, text="Language", style="Muted.TLabel").pack(side="right", padx=(0, 8))
        # Đổi ngôn ngữ -> THAY THẾ nội dung bằng bản mặc định của ngôn ngữ mới.
        self.combo_calendar_lang.bind("<<ComboboxSelected>>", lambda e: self._apply_calendar_body_lang())
        appt_container, self.txt_appt_body = make_scrollable_text(calendar.body, width=40, height=4)
        default_body = build_calendar_body("en", *self._calendar_body_args())
        self.txt_appt_body.insert("1.0", default_body)
        self.var_appt_body_default = default_body  # Keep track of default for reset
        appt_container.pack(fill="x")
        ttk.Button(calendar.body, text="📅 Send Calendar Invite to the list above", style="Primary.TButton",
                   command=self._send_calendar).pack(anchor="w", pady=(12, 0))

        # ── after the event: totals, payment rounds, the per-person table ──
        # Round 1 is the main event; a follow-up (second venue, another
        # evening) is another round with its own Attend and amount columns.
        self.var_total_actual_attend = tk.StringVar(value="0")   # round 1
        self.var_round1_collected = tk.StringVar(value="0")
        self.var_remaining_amount = tk.StringVar(value="0")      # round 1
        # "Amount paid" of round 1 — what was actually paid out (a restaurant
        # deposit...), typed by hand and saved per event (AmountPaid).
        self.var_amount_paid = tk.StringVar(value="0")
        self.var_total_collected_amount = tk.StringVar(value="0")  # all rounds
        self.var_total_paid_amount = tk.StringVar(value="0")
        self.var_total_remaining_amount = tk.StringVar(value="0")
        kpi_row(f, [
            ("Attended", self.var_total_actual_attend,
             {"dot": COLORS["success_fill"], "footnote": "Main event (round 1)"}),
            ("Collected", self.var_total_collected_amount, {"footnote": "All rounds"}),
            ("Paid out", self.var_total_paid_amount, {"footnote": "All rounds"}),
            ("Remaining", self.var_total_remaining_amount, {"footnote": "Collected − paid out"}),
        ]).pack(**gap)
        self.var_amount_paid.trace_add("write", lambda *a: self._on_amount_paid_changed())

        rounds = Card(f, "Payment rounds",
                      "Round 1 is the main event. Add a round for a follow-up — a second venue, another "
                      "evening — and it gets its own Attend and amount columns in the table below. Type "
                      "what was actually paid out in each round; Remaining is worked out for you.")
        rounds.pack(**gap)
        ttk.Button(rounds.actions, text="➕ Add round", command=self._add_amount_round).pack()
        self.rounds_container = ttk.Frame(rounds.body, style="Card.TFrame")
        self.rounds_container.pack(fill="x")
        # Shown only while a paid box holds text not read as what it shows.
        self.lbl_rounds_warning = WrapLabel(rounds.body, style="Danger.TLabel", text="")

        attendance = Card(f, "Attendance & payment",
                          "Fill in after the event. The list follows the Yes / Maybe votes each time this "
                          "page opens, and every change is saved automatically.")
        attendance.pack(**gap)
        # Every edit is saved to the database already; this only writes a
        # file to share or archive.
        ttk.Button(attendance.actions, text="📊 Export to Excel", command=self._export_attendance_to_excel)\
            .pack()
        WrapLabel(attendance.body, style="Hint.TLabel",
                  text="Click an Attend or Free cell to tick it, or its column header to tick everyone. "
                       "Ticking Attend fills that round's amount from the expected event budget (0 if Free; "
                       "Free counts for every round), and you can type over it. Double-click Name, Vote or "
                       "an amount to edit it, or an amount column's header to rename its round. Ctrl+C / "
                       "Ctrl+V copy to and from Excel; Delete clears a row's tracking.").pack(
                           fill="x", pady=(0, 10))
        # Columns: No., Name, Vote, round 1's Attend/Free/amount, then an
        # Attend + amount pair per further round (see
        # _rebuild_attendance_tree_columns()).
        attend_container, self.tree_attendance = make_scrollable_treeview(
            attendance.body, columns=("no", "name", "vote", "actual_attend", "free", "amount"),
            height=12, horizontal=True)
        attend_container.pack(fill="both", expand=True)
        self.tree_attendance.bind("<Double-1>", self._on_attendance_tree_double_click)
        self.tree_attendance.bind("<Button-1>", self._on_attendance_free_click)
        # Delete/Backspace clears the row's tracking fields (Attend, Free and
        # every amount); Name and Vote come from Tab 4 and are left alone.
        self.tree_attendance.bind("<Delete>", self._on_attendance_delete_key)
        self.tree_attendance.bind("<BackSpace>", self._on_attendance_delete_key)
        self._rebuild_attendance_tree_columns()
        self._rebuild_rounds_ui()

        # ── Thank You email: to everyone with "Actual Attend" = Yes; refreshes
        # from Tab 1 until hand-edited (see _refresh_thankyou_body_display()). ──
        thankyou = Card(f, "Thank-you email",
                        "To everyone who attended the main event (round 1). Includes a table of every "
                        "round's money, and attaches the attendance & payment report and the Calendar "
                        "Invite if it can be found in your Calendar.")
        thankyou.pack(fill="x")
        thankyou_header = ttk.Frame(thankyou.body)
        thankyou_header.pack(fill="x", pady=(0, 8))
        ttk.Label(thankyou_header, text="Language", style="Field.TLabel").pack(side="left", padx=(0, 8))
        self.combo_thankyou_lang = ttk.Combobox(
            thankyou_header, width=26, state="readonly",
            values=[LANG_LABELS["en"], LANG_LABELS["ja"], LANG_LABELS["vi"], LANG_LABELS["bilingual"]],
        )
        self.combo_thankyou_lang.current(0)  # default: English
        self.combo_thankyou_lang.pack(side="left")
        self.combo_thankyou_lang.bind("<<ComboboxSelected>>", lambda e: self._apply_thankyou_body_lang())
        # The text is generated when the page opens; after editing the table,
        # this rebuilds it from the current figures.
        ttk.Button(thankyou_header, text="🔄 Update from table",
                   command=self._update_thankyou_body_from_table).pack(side="left", padx=(8, 0))
        thankyou_container, self.txt_thankyou_body = make_scrollable_text(thankyou.body, width=40, height=10)
        default_thankyou_body = build_thankyou_body("en", *self._thankyou_body_args())
        self.txt_thankyou_body.insert("1.0", default_thankyou_body)
        self.var_thankyou_body_default = default_thankyou_body  # Keep track of default for reset
        thankyou_container.pack(fill="x")
        ttk.Button(thankyou.body, text="📧 Send Thank You email to confirmed attendees", style="Primary.TButton",
                   command=self._send_thank_you_email).pack(anchor="w", pady=(12, 0))

    # ── Tab 5: who owns the table ──

    def _history_record(self, event_id):
        """event_id's History row, or None. Raises when the database cannot
        be read, so a caller can tell "no row" from "could not look"."""
        return next((r for r in db.load_history(self.history_path.get())
                     if r.get("EventID") == event_id), None)

    def _adopt_attendance_owner(self, event_id):
        """Makes event_id the owner of Tab 5's table and loads everything it
        has - the people, the payment rounds, the first round's name and the
        next round key - together, or clears all of them when event_id is
        empty. The only place the owner changes, so one event's rounds can
        never stay on screen with another event's table and be saved under
        it. When the saved table cannot be read, nothing is owned (so nothing
        is saved over it) and the user is told. Returns True when the event
        had a table saved."""
        event_id = event_id or None
        roster, rounds, label, next_index = {}, [], "", 2
        if event_id:
            path = self.history_path.get()
            try:
                roster = db.load_attendance_roster(event_id, path)
                rounds = db.load_attendance_rounds(event_id, path)
                next_index = db.next_round_index(event_id, path)
                label = (self._history_record(event_id) or {}).get("Round1Label") or ""
            except Exception as exc:
                messagebox.showerror(
                    "Attendance not loaded",
                    f"Couldn't read the Attendance & Payment table of '{event_id}':\n\n{exc}\n\n"
                    "It is shown empty and nothing on this page will be saved until it can be read.")
                event_id, roster, rounds, label, next_index = None, {}, [], "", 2
        self._attendance_event = event_id
        self._attendance_roster = roster
        self._extra_rounds = rounds
        self._round1_label = label
        self._next_round_index = next_index
        self._rebuild_attendance_tree_columns()
        self._rebuild_rounds_ui()
        self._render_attendance_tree()
        return bool(roster or rounds)

    def _round1_label_text(self):
        return (self._round1_label or "").strip() or ROUND1_DEFAULT_LABEL

    def _attendance_figures(self):
        """Every round's figures for the table on screen, round 1 first."""
        return payment_rounds(self._attendance_roster.values(), self._round1_label_text(),
                              self.var_amount_paid.get(), self._extra_rounds)

    # ── Tab 5: the table ──

    def _rebuild_attendance_tree_columns(self):
        """Sets the table's columns: No., Name, Vote, round 1's Attend / Free /
        amount, then an Attend + amount pair per further round, headed with
        the rounds' names. Column widths are fixed, so the horizontal
        scrollbar appears when the rounds run off screen."""
        tree = self.tree_attendance
        r1 = self._round1_label_text()
        cols = ["no", "name", "vote", "actual_attend", "free", "amount"]
        heads = ["No.", "Name", "Vote", f"{r1} Attend", "Free", r1]
        widths = [44, 240, 70, max(130, 8 * len(r1) + 90), 70, max(100, 8 * len(r1) + 30)]
        for r in self._extra_rounds:
            cols += [f"attend_{r['key']}", f"extra_{r['key']}"]
            heads += [f"{r['label']} Attend", r["label"]]
            widths += [max(130, 8 * len(r["label"]) + 90), max(100, 8 * len(r["label"]) + 30)]
        tree["columns"] = cols
        for c, h, w in zip(cols, heads, widths):
            checkbox = c in ("actual_attend", "free") or c.startswith("attend_")
            anchor = "e" if (c == "amount" or c.startswith("extra_")) else (
                "w" if c == "name" else "center")
            if checkbox:
                # Click the header to tick or untick every row.
                tree.heading(c, text=f"⬜ {h}", anchor=anchor,
                             command=lambda k=c: self._toggle_all_attendance_column(k))
            else:
                tree.heading(c, text=h, anchor=anchor, command="")
            tree.column(c, width=w, minwidth=40, anchor=anchor, stretch=c == "name")
        extra_cols = {c for c in cols if c.startswith(("attend_", "extra_"))}
        self._enable_treeview_copy_paste(
            tree, on_commit=self._commit_attendance_edit,
            editable_cols={"name", "vote", "actual_attend", "free", "amount"} | extra_cols)

    def _rebuild_rounds_ui(self):
        """The "Payment rounds" card: one row per round with its attendees,
        collected, paid (typed) and remaining, and Rename / Remove."""
        if not hasattr(self, "rounds_container"):
            return
        grid = self.rounds_container
        for child in grid.winfo_children():
            child.destroy()
        self._round_vars = {}
        for i, head in enumerate(("Round", "Attendees", "Collected", "Paid out", "Remaining")):
            ttk.Label(grid, text=head, style="Field.TLabel").grid(
                row=0, column=i, sticky="w" if i == 0 else "e", padx=(0 if i == 0 else 16, 0), pady=(0, 6))
        grid.columnconfigure(0, weight=1)

        def row(r, label, attendees, collected, paid_var, remaining, key, removable):
            ttk.Label(grid, text=label, style="Strong.TLabel").grid(row=r, column=0, sticky="w", pady=3)
            ttk.Label(grid, textvariable=attendees).grid(row=r, column=1, sticky="e", padx=(16, 0))
            ttk.Label(grid, textvariable=collected).grid(row=r, column=2, sticky="e", padx=(16, 0))
            ttk.Entry(grid, textvariable=paid_var, width=12, justify="right").grid(
                row=r, column=3, sticky="e", padx=(16, 0))
            ttk.Label(grid, textvariable=remaining, style="Strong.TLabel").grid(
                row=r, column=4, sticky="e", padx=(16, 0))
            buttons = ttk.Frame(grid, style="Card.TFrame")
            buttons.grid(row=r, column=5, sticky="e", padx=(16, 0))
            ttk.Button(buttons, text="✎ Rename", style="Small.TButton",
                       command=lambda k=key: self._rename_round(k)).pack(side="left")
            if removable:
                ttk.Button(buttons, text="✖ Remove", style="Small.TButton",
                           command=lambda k=key: self._remove_amount_round(k)).pack(side="left", padx=(6, 0))

        row(1, self._round1_label_text(), self.var_total_actual_attend, self.var_round1_collected,
            self.var_amount_paid, self.var_remaining_amount, None, False)
        for i, r in enumerate(self._extra_rounds, start=2):
            v = {"attend": tk.StringVar(value="0"), "collected": tk.StringVar(value="0"),
                 "remaining": tk.StringVar(value="0"),
                 "paid": tk.StringVar(value=r.get("amount_paid") or "0")}
            v["paid"].trace_add("write", lambda *a, k=r["key"]: self._on_round_paid_changed(k))
            self._round_vars[r["key"]] = v
            row(i, r["label"], v["attend"], v["collected"], v["paid"], v["remaining"], r["key"], True)
        self._update_attendance_totals()

    def _render_attendance_tree(self):
        tree = self.tree_attendance
        tree.delete(*tree.get_children())
        for i, (email, info) in enumerate(self._attendance_roster.items(), start=1):
            amount = info.get("amount", 0.0)
            values = [i, info["name"], info["vote"],
                      "✅" if is_yes(info.get("actual_attend")) else "⬜",
                      "✅" if info.get("free") else "⬜",
                      f"{amount:,.0f}" if amount else ""]
            for r in self._extra_rounds:
                key = r["key"]
                values.append("✅" if is_yes((info.get("extra_attends") or {}).get(key)) else "⬜")
                v = (info.get("extra_amounts") or {}).get(key, 0.0)
                values.append(f"{v:,.0f}" if v else "")
            tree.insert("", "end", iid=email, values=values)
        self._update_attendance_totals()
        self._update_attendance_header_checkmarks()

    def _update_attendance_header_checkmarks(self):
        """✅ on a checkbox column's header when every row is ticked."""
        tree = self.tree_attendance
        shown = [self._attendance_roster.get(iid, {}) for iid in tree.get_children()]

        def mark(ticked):
            return "✅" if shown and all(ticked(info) for info in shown) else "⬜"

        tree.heading("actual_attend",
                     text=f"{mark(lambda i: is_yes(i.get('actual_attend')))} {self._round1_label_text()} Attend")
        tree.heading("free", text=f"{mark(lambda i: i.get('free'))} Free")
        for r in self._extra_rounds:
            key = r["key"]
            tree.heading(f"attend_{key}", text=(
                f"{mark(lambda i, k=key: is_yes((i.get('extra_attends') or {}).get(k)))} {r['label']} Attend"))

    def _sync_attendance_amount(self, email):
        """Round 1's amount from Attend / Free: the expected event budget when
        attending and not Free, else 0. A typed amount stays until Attend or
        Free changes again."""
        info = self._attendance_roster[email]
        info["amount"] = amount_for(is_yes(info.get("actual_attend")), info.get("free"),
                                    parse_amount_from_text(self.var_budget.get()))

    def _sync_round_amount(self, email, round_key):
        """The same rule for a further round, from that round's Attend."""
        info = self._attendance_roster[email]
        attending = is_yes((info.get("extra_attends") or {}).get(round_key))
        info.setdefault("extra_amounts", {})[round_key] = amount_for(
            attending, info.get("free"), parse_amount_from_text(self.var_budget.get()))

    def _commit_attendance_edit(self, row_id, col_name, new_value):
        """A typed or pasted value. Attend and Free accept ✅/⬜ (copied from
        this table) as well as Yes/No/TRUE/1 (from Excel)."""
        email = row_id
        if email not in self._attendance_roster:
            return
        info = self._attendance_roster[email]
        new_value = new_value.strip()
        truthy = new_value.lower() in TRUTHY
        # An empty Attend stays empty (not attending is "No", not blank).
        attend_value = "Yes" if truthy else ("No" if new_value else "")
        if (col_name == "amount" or col_name.startswith("extra_")) and new_value and (
                unclear_typed_amount(new_value) or parse_typed_amount(new_value) < 0):
            # Reading it as 0, a negative as positive or "1 000" as 1 would
            # silently change what that person paid.
            if self._paste_refusals is not None:
                self._paste_refusals.append(new_value)
                return
            messagebox.showwarning(
                "Amount not changed",
                f"“{new_value}” is not an amount someone paid, so the cell was left unchanged.\n\n"
                "Type just the figure (e.g. 6000 or 6,000), or clear the cell for 0.")
            return
        if col_name == "name":
            info["name"] = new_value
        elif col_name == "vote":
            info["vote"] = new_value
        # Amounts follow Attend and Free only when they change: pasting back
        # a row whose ticks are as they were keeps its typed amounts.
        elif col_name == "actual_attend":
            if attend_value != (info.get("actual_attend") or ""):
                info["actual_attend"] = attend_value
                self._sync_attendance_amount(email)
        elif col_name == "free":
            if truthy != bool(info.get("free")):
                info["free"] = truthy
                # Free exempts the person from every round.
                self._sync_attendance_amount(email)
                for r in self._extra_rounds:
                    self._sync_round_amount(email, r["key"])
        elif col_name == "amount":
            info["amount"] = parse_amount_from_text(new_value)
        elif col_name.startswith("attend_"):
            key = col_name[len("attend_"):]
            attends = info.setdefault("extra_attends", {})
            if attend_value != (attends.get(key) or ""):
                attends[key] = attend_value
                self._sync_round_amount(email, key)
        elif col_name.startswith("extra_"):
            info.setdefault("extra_amounts", {})[col_name[len("extra_"):]] = parse_amount_from_text(new_value)
        self._render_attendance_tree()
        self._save_attendance_sheet_to_file(silent=True)

    def _on_attendance_tree_double_click(self, event):
        """On a cell: Name, Vote and the amount columns open an edit box (the
        checkbox columns toggle on a single click instead). On the header of
        an amount column: rename that round. The Attend headers are left to
        their single-click "tick everyone" - a double-click there would fire
        it twice and turn a partly ticked column into an unticked one."""
        tree = self.tree_attendance
        region = tree.identify("region", event.x, event.y)
        col_name = self._tree_column_name_at(tree, tree.identify_column(event.x))
        if col_name is None:
            return
        if region == "heading":
            if col_name == "amount":
                self._rename_round(None)
            elif col_name.startswith("extra_"):
                self._rename_round(col_name[len("extra_"):])
            return
        row_id = tree.identify_row(event.y)
        if region == "cell" and row_id and (
                col_name in ("name", "vote", "amount") or col_name.startswith("extra_")):
            self._begin_cell_edit(tree, row_id, col_name, self._commit_attendance_edit)

    def _on_attendance_free_click(self, event):
        """Single click on a checkbox cell - Attend (any round) or Free -
        toggles it. Free exempts the person from every round; Attend fills
        that round's amount from the expected event budget, or empties it."""
        tree = self.tree_attendance
        if tree.identify("region", event.x, event.y) != "cell":
            return
        row_id = tree.identify_row(event.y)
        col_name = self._tree_column_name_at(tree, tree.identify_column(event.x))
        if not row_id or row_id not in self._attendance_roster or col_name is None:
            return
        info = self._attendance_roster[row_id]
        if col_name == "free":
            info["free"] = not info.get("free", False)
            self._sync_attendance_amount(row_id)
            for r in self._extra_rounds:
                self._sync_round_amount(row_id, r["key"])
        elif col_name == "actual_attend":
            info["actual_attend"] = "No" if is_yes(info.get("actual_attend")) else "Yes"
            self._sync_attendance_amount(row_id)
        elif col_name.startswith("attend_"):
            key = col_name[len("attend_"):]
            attends = info.setdefault("extra_attends", {})
            attends[key] = "No" if is_yes(attends.get(key)) else "Yes"
            self._sync_round_amount(row_id, key)
        else:
            return
        self._render_attendance_tree()
        self._save_attendance_sheet_to_file(silent=True)

    def _toggle_all_attendance_column(self, col_name):
        """A checkbox column's header: ticks every row, or unticks every row
        when all are ticked already. Amounts follow as for a single tick."""
        roster = self._attendance_roster
        shown = [iid for iid in self.tree_attendance.get_children() if iid in roster]
        if not shown:
            return
        # Rows already as the header sets them keep their typed amounts.
        if col_name == "free":
            new_state = not all(roster[iid].get("free") for iid in shown)
            for iid in shown:
                if bool(roster[iid].get("free")) == new_state:
                    continue
                roster[iid]["free"] = new_state
                self._sync_attendance_amount(iid)
                for r in self._extra_rounds:
                    self._sync_round_amount(iid, r["key"])
        elif col_name == "actual_attend":
            value = "No" if all(is_yes(roster[iid].get("actual_attend")) for iid in shown) else "Yes"
            for iid in shown:
                if roster[iid].get("actual_attend") == value:
                    continue
                roster[iid]["actual_attend"] = value
                self._sync_attendance_amount(iid)
        elif col_name.startswith("attend_"):
            key = col_name[len("attend_"):]
            value = "No" if all(is_yes((roster[iid].get("extra_attends") or {}).get(key))
                                for iid in shown) else "Yes"
            for iid in shown:
                attends = roster[iid].setdefault("extra_attends", {})
                if attends.get(key) == value:
                    continue
                attends[key] = value
                self._sync_round_amount(iid, key)
        else:
            return
        self._render_attendance_tree()
        self._save_attendance_sheet_to_file(silent=True)

    def _on_attendance_delete_key(self, event):
        """Delete/Backspace on selected rows clears their tracking fields -
        Attend, Free and the amount of every round - back to untouched. Name
        and Vote come from Tab 4 and are left alone. (A Treeview selects
        whole rows, so the row's tracking fields are cleared together.)"""
        changed = False
        for row_id in self.tree_attendance.selection():
            info = self._attendance_roster.get(row_id)
            if info is None:
                continue
            info.update(actual_attend="", free=False, amount=0.0, extra_amounts={}, extra_attends={})
            changed = True
        if changed:
            self._render_attendance_tree()
            self._save_attendance_sheet_to_file(silent=True)
        return "break"

    def _update_attendance_totals(self):
        """Every figure on Tab 5 from the table and the paid boxes: each
        round's row in the rounds card, and the totals over all rounds."""
        if not hasattr(self, "lbl_rounds_warning"):
            return  # Tab 5 is still being built
        figures = self._attendance_figures()
        first = figures[0]
        self.var_total_actual_attend.set(str(first.attendees))
        self.var_round1_collected.set(format_amount(first.collected))
        self.var_remaining_amount.set(format_amount(first.remaining))
        for f in figures[1:]:
            v = self._round_vars.get(f.key)
            if v is not None:
                v["attend"].set(str(f.attendees))
                v["collected"].set(format_amount(f.collected))
                v["remaining"].set(format_amount(f.remaining))
        typed = {r["key"]: r.get("amount_paid") for r in self._extra_rounds}
        typed[first.key] = self.var_amount_paid.get()
        unclear = [f"{f.label}: “{typed.get(f.key)}” is read as {format_amount(f.paid)}"
                   for f in figures if unclear_typed_amount(typed.get(f.key))]
        self._show_typed_amount_warning(self.lbl_rounds_warning, unclear, self.rounds_container)
        collected, paid, remaining = round_totals(figures)
        self.var_total_collected_amount.set(format_amount(collected))
        self.var_total_paid_amount.set(format_amount(paid))
        self.var_total_remaining_amount.set(format_amount(remaining))
        # Tab 6's Event + Gift figures use these when linked.
        self._update_gift_summary()

    @staticmethod
    def _show_typed_amount_warning(label, unclear, after):
        """Shows, under `after`, which typed amounts are not read as they
        look - "1 000" counts as 1 everywhere it is used (totals, History,
        the emails, Excel) - or hides the warning when there are none."""
        if unclear:
            label.configure(text="⚠ Type just the figure (e.g. 1000 or 1,000). " + "; ".join(unclear) + ".")
            label.pack(fill="x", pady=(8, 0), after=after)
        else:
            label.configure(text="")
            label.pack_forget()

    def _refresh_remaining_amount(self):
        """Kept for its callers: every Remaining moves together now."""
        self._update_attendance_totals()

    # ── Tab 5: payment rounds ──

    def _add_amount_round(self):
        """"➕ Add round": asks for a name and adds the round - an Attend and
        an amount column in the table and a row in the rounds card."""
        event_id = self._attendance_event
        if not event_id:
            messagebox.showwarning(
                "No event", "Open this page with an Event ID on Tab 1 first - rounds are saved with "
                            "the event the table belongs to.")
            return
        label = simpledialog.askstring(
            "Add round", "Name for the new payment round (e.g. \"Karaoke\"):", parent=self)
        if label is None:
            return  # Cancel
        label = label.strip() or f"Round {len(self._extra_rounds) + 2}"
        key = f"round_{self._next_round_index}"
        self._next_round_index += 1
        self._extra_rounds.append({"key": key, "label": label, "amount_paid": "0"})
        self._after_rounds_changed()

    def _rename_round(self, round_key):
        """Renames a round (None = round 1). The name follows into the table,
        the rounds card, the Excel report and the Thank-you email."""
        if round_key is None:
            current = self._round1_label_text()
        else:
            current = next((r["label"] for r in self._extra_rounds if r["key"] == round_key), None)
            if current is None:
                return
        label = simpledialog.askstring(
            "Rename round", "New name for this payment round:", initialvalue=current, parent=self)
        if label is None or not label.strip() or label.strip() == current:
            return
        label = label.strip()
        if round_key is None:
            self._round1_label = label
            if self._attendance_event:
                self._save_event_fields(self._attendance_event, {"Round1Label": label},
                                        "The name of the first round")
        else:
            for r in self._extra_rounds:
                if r["key"] == round_key:
                    r["label"] = label
        self._after_rounds_changed(save=round_key is not None)

    def _remove_amount_round(self, round_key):
        """Removes a further round and everyone's values for it, after a
        confirmation. Round 1 cannot be removed."""
        label = next((r["label"] for r in self._extra_rounds if r["key"] == round_key), round_key)
        if not messagebox.askyesno(
                "Remove round",
                f"Remove the \"{label}\" round? Everyone's Attend and amount for it will be deleted."):
            return
        self._extra_rounds = [r for r in self._extra_rounds if r["key"] != round_key]
        for info in self._attendance_roster.values():
            (info.get("extra_amounts") or {}).pop(round_key, None)
            (info.get("extra_attends") or {}).pop(round_key, None)
        self._after_rounds_changed()

    def _after_rounds_changed(self, save=True):
        self._rebuild_attendance_tree_columns()
        self._rebuild_rounds_ui()
        self._render_attendance_tree()
        if save and not self._save_attendance_sheet_to_file(silent=True) and self._attendance_event:
            self._report_save_failure(
                "rounds", "The payment rounds",
                f"they could not be written to {self.history_path.get()}. Check that the file is "
                "writable and not open in another program, then change them again.")
        # The Thank-you text names the rounds; rebuild it unless hand-edited.
        self._refresh_thankyou_body_display()

    def _on_round_paid_changed(self, round_key):
        """A further round's "paid" box: totals follow, and it is saved with
        the rounds."""
        v = self._round_vars.get(round_key)
        if v is None:
            return
        for r in self._extra_rounds:
            if r["key"] == round_key:
                r["amount_paid"] = v["paid"].get()
        self._update_attendance_totals()
        self._save_attendance_sheet_to_file(silent=True)

    def _on_amount_paid_changed(self):
        """Recomputes Remaining amount and auto-saves "Amount paid" as the user
        types, to the event it belongs to (self._amount_paid_event) — never
        to whatever Tab 1 shows. It is money, so a save that fails is
        reported, once per run rather than on every keystroke: it used to
        swallow every exception, losing the figure while the UI still showed
        it as entered. It only UPDATEs: an event that is not in History yet
        gets no row of its own, and the user is told to save it first."""
        self._update_attendance_totals()
        if self._amount_paid_quiet:
            return
        event_id = self._amount_paid_event
        if not event_id:
            return
        try:
            saved = db.update_event(event_id, {"AmountPaid": self.var_amount_paid.get()},
                                    self.history_path.get())
        except Exception as exc:
            saved, reason = False, (f"It could not be written to the database:\n\n{exc}\n\n"
                                    f"Check that {self.history_path.get()} is writable and not open "
                                    f"in another program, then re-enter it.")
        else:
            reason = (f"Event ID '{event_id}' is not in History yet. Save it on Tab 1 "
                      f"('💾 Save event details'), then re-enter Amount paid.")
        if saved:
            self._amount_paid_save_failed = False
            # Typing what was paid out is working on this event's money -
            # even with nobody in the table (a cancelled event's costs) - so
            # History follows it, as for every other Tab 5 change.
            self._sync_event_money(event_id, was_tracked=True)
        elif not self._amount_paid_save_failed:
            self._amount_paid_save_failed = True
            messagebox.showwarning("Amount paid not saved",
                                   f"The value on screen is NOT saved. {reason}")

    def _set_amount_paid_quietly(self, value):
        """Shows `value` as Amount paid without the trace saving it anywhere."""
        self._amount_paid_quiet = True
        try:
            self.var_amount_paid.set(value)
        finally:
            self._amount_paid_quiet = False
        self._update_attendance_totals()

    def _refresh_amount_paid_for_current_event(self):
        """Tab 5 opened: when the Amount paid on screen belongs to another
        event than Tab 1's, show the one saved for Tab 1's event instead."""
        current_id = self.var_event_id.get().strip()
        if self._amount_paid_event == (current_id or None):
            return
        saved = "0"
        if current_id:
            try:
                rec = self._history_record(current_id)
            except Exception as exc:
                # Shown as 0 but owned by no event, so neither typing here nor
                # Tab 1's save writes the 0 over the real figure; the next
                # visit to this page reads it again.
                messagebox.showerror(
                    "Amount paid not loaded",
                    f"Couldn't read the Amount paid of '{current_id}':\n\n{exc}\n\n"
                    "It is shown as 0 and will not be saved until it can be read.")
                self._amount_paid_event = None
                self._set_amount_paid_quietly(saved)
                return
            if rec and rec.get("AmountPaid"):
                saved = rec["AmountPaid"]
        self._amount_paid_event = current_id or None
        self._set_amount_paid_quietly(saved)

    def _refresh_attendance_list(self):
        """Rebuilds the table from the Yes/Maybe voters scanned on Tab 4 - the
        same people as the Calendar Invite list. Marks and amounts already in
        the table are kept, every round's included; new voters start as
        attending round 1, with its amount filled in.

        Runs when Tab 5 opens. The table belongs to one event
        (self._attendance_event): when Tab 1 now shows another, that event's
        own table and rounds are loaded first (_adopt_attendance_owner()), so
        marks and rounds never move between events; votes are merged only from
        a Tab 4 table that was scanned for that same event."""
        current_id = self.var_event_id.get().strip()
        if self._attendance_event != (current_id or None):
            self._adopt_attendance_owner(current_id)
        if not current_id or current_id != self._last_scanned_event_id:
            self._render_attendance_tree()
            return
        new_roster = {}
        newly_added = []
        for iid in self.tree_responses.get_children():
            _manual, name, email, vote, _received = self.tree_responses.item(iid, "values")
            if vote not in ("Yes", "Maybe"):
                continue
            prev = self._attendance_roster.get(email)
            if prev is None:
                newly_added.append(email)
                prev = {}
            new_roster[email] = {
                "name": name,
                "vote": vote,
                "actual_attend": prev.get("actual_attend", "Yes"),
                "free": prev.get("free", False),
                "amount": prev.get("amount", 0.0),
                "extra_amounts": dict(prev.get("extra_amounts") or {}),
                "extra_attends": dict(prev.get("extra_attends") or {}),
            }
        people_changed = set(new_roster) != set(self._attendance_roster)
        self._attendance_roster = new_roster
        # Only new people get round 1's amount filled in; anyone already in
        # the table keeps theirs, even a 0 chosen by hand.
        for email in newly_added:
            self._sync_attendance_amount(email)
        self._render_attendance_tree()
        # Votes added or dropped people: save the table as shown, so what is
        # saved - and History's money computed from it - matches the screen.
        # Rebuilt from this event's votes, an empty table means nobody comes.
        if people_changed:
            self._save_attendance_sheet_to_file(silent=True, reconciled=True)

    # ── Attendance & Payment / Responded result — lưu vào database ──
    # Every change on Tab 5 is saved as it is made; '⬅ Load setup from
    # selected event' reads it back. "📊 Export to Excel" only writes a copy.

    def _save_attendance_sheet_to_file(self, silent=True, reconciled=False):
        """Saves the table and its payment rounds, in one transaction, to the
        event the table belongs to; then History's money columns follow.
        Returns True when saved. An empty table on screen replaces a saved
        one only when `reconciled` - rebuilt from the event's own votes, so
        nobody is attending; otherwise it may simply not have been loaded."""
        event_id = self._attendance_event
        if not event_id:
            return False
        cleared = reconciled and not self._attendance_roster
        was_tracked = self._was_money_tracked(event_id)
        try:
            db.save_attendance(event_id, self._attendance_roster or ({} if cleared else None),
                               self._extra_rounds, self.history_path.get())
        except Exception as exc:
            if not silent:
                raise
            # Money: a lost save is reported (once per run), not swallowed.
            self._report_save_failure("attendance", "Attendance & payment", self._unwritable_reason(exc))
            return False
        self._save_failures_reported.discard("attendance")
        self._sync_event_money(event_id, was_tracked=was_tracked)
        return True

    def _load_attendance_sheet_from_file(self, event_id):
        """'⬅ Load setup from selected event': the event's saved table and
        rounds. True when it had any."""
        return self._adopt_attendance_owner(event_id)

    # ── History's money columns ──

    def _money_tracked(self, event_id):
        """Whether what is saved for event_id gives History's money columns:
        an attendance table, a payment round, someone marked as having
        contributed to the gift, or a gift price. A gift list's rows do not
        count on their own: they exist as soon as Tab 6 opens or "Send email"
        is ticked. Raises when the database cannot be read."""
        path = self.history_path.get()
        rec = self._history_record(event_id) or {}
        return bool(db.load_attendance_roster(event_id, path)
                    or db.load_attendance_rounds(event_id, path)
                    or any(info.get("checked") for info in db.load_gift_roster(event_id, path).values())
                    or (rec.get("GiftItemPrice") or "").strip())

    def _was_money_tracked(self, event_id):
        """_money_tracked() before a save, for _sync_event_money() after it;
        False when it cannot be read (the save and the sync report that)."""
        if self._suspend_money_sync or not event_id:
            return False  # no sync follows
        try:
            return self._money_tracked(event_id)
        except Exception:
            return False

    def _sync_event_money(self, event_id, was_tracked=False):
        """Recomputes event_id's money columns in History - Actual Att.
        (main), Cost/Person, Income, Expense, Balance - from what is SAVED for
        it: its attendance table and rounds, Amount paid, gift contributions
        and gift price. Called after each successful save, so the figures can
        only describe one event's saved data, never a mix of tables on
        screen. An event with nothing tracked here (_money_tracked()) keeps
        whatever History holds - figures typed in the other copy of the app,
        for instance - unless `was_tracked`: it was tracked before the save
        that called this, so the save removed the last of it (the table
        emptied, the last contributor unticked, the gift price cleared) and
        the figures left behind would count money no longer there. One
        without a History row is left alone too."""
        if self._suspend_money_sync or not event_id:
            return
        path = self.history_path.get()
        try:
            rec = self._history_record(event_id)
            if rec is None:
                return
            if not was_tracked and not self._money_tracked(event_id):
                return
            roster = db.load_attendance_roster(event_id, path)
            rounds = db.load_attendance_rounds(event_id, path)
            gift = db.load_gift_roster(event_id, path)
            figures = payment_rounds(roster.values(), rec.get("Round1Label") or ROUND1_DEFAULT_LABEL,
                                     rec.get("AmountPaid"), rounds)
            fields = history_figures(figures, contributed_total(gift.values()), rec.get("GiftItemPrice"))
            changed = {k: v for k, v in fields.items() if str(rec.get(k) or "") != v}
            if changed:
                db.update_event(event_id, changed, path)
                if hasattr(self, "tree_history"):
                    self._refresh_history_tree()
        except Exception as exc:
            self._report_save_failure(
                "money", "History's money columns",
                f"they could not be updated:\n\n{exc}\n\nThey are recalculated at the next change "
                "on Attendance & payment or Gift contribution.")

    def _unwritable_reason(self, exc):
        return (f"it could not be written to the database:\n\n{exc}\n\nCheck that "
                f"{self.history_path.get()} is writable and not open in another program, "
                "then change it again.")

    def _report_save_failure(self, key, what, reason):
        """Tells the user once per run that `what` was not saved."""
        if key in self._save_failures_reported:
            return
        self._save_failures_reported.add(key)
        messagebox.showwarning("Not saved", f"{what}: the value on screen is NOT saved - {reason}")

    def _save_event_fields(self, event_id, fields, what):
        """Writes `fields` to event_id's History row - UPDATE only, like every
        automatic write - and reports once when that is impossible. Returns
        True when saved."""
        key = f"fields:{what}"
        try:
            saved = db.update_event(event_id, fields, self.history_path.get())
        except Exception as exc:
            saved, reason = False, (f"it could not be written to the database:\n\n{exc}\n\nCheck that "
                                    f"{self.history_path.get()} is writable, then change it again.")
        else:
            reason = (f"Event ID '{event_id}' is not in History yet. Save it on Tab 1 "
                      "('💾 Save event details'), which also saves this.")
        if saved:
            self._save_failures_reported.discard(key)
        else:
            self._report_save_failure(key, what, reason)
        return saved

    def _vote_counts_record(self):
        """The Tab 4 table's counts, keyed by their History columns.
        TotalInvited is the number of people tracked - Tab 2's rows with
        group addresses expanded - the same denominator the table uses."""
        counts = self._vote_counts
        return {
            "TotalInvited": len(self._last_responses_roster or []),
            "Yes": counts.get("Yes", 0),
            "No": counts.get("No", 0),
            "Maybe": counts.get("Maybe", 0),
            "NoResponse": counts.get("No response", 0),
        }

    def _autosave_scan_results(self):
        """Saves Tab 4 to the database after every scan and every manual vote
        edit: the responses (so '⬅ Load setup from selected event' restores
        them without re-scanning Outlook) and the Yes/No/Maybe counts on the
        event's History row. Keyed on the Event ID the table was scanned
        for, which differs from Tab 1's while the status banner shows a stale
        table; with no scan behind the table there is nothing to key on and
        nothing is saved. Best-effort and silent, like every other auto-save."""
        event_id = self._last_scanned_event_id
        if not event_id:
            return
        path = self.history_path.get()
        try:
            db.save_responses(event_id, self.responses, self._last_scan_time, path)
            # UPDATE only: a scan never creates a History row of its own. The
            # banner says so when the event has none yet.
            self._scan_in_history = db.update_event(event_id, self._vote_counts_record(), path)
        except Exception:
            return
        if self._scan_in_history:
            self._refresh_history_tree()

    def _load_responded_result_from_file(self, event_id):
        """Đọc kết quả Responded đã lưu trong database cho event_id, nạp
        vào self.responses và vẽ lại bảng Tab 4 — dùng bởi '⬅ Load setup
        from selected event'. Trả về True nếu có dữ liệu để nạp."""
        if not hasattr(self, "tree_responses"):
            return False
        try:
            responses, last_scan = db.load_responses(event_id, self.history_path.get())
        except Exception:
            return False
        if not responses:
            return False
        self.responses = responses
        self._last_scanned_event_id = event_id
        self._last_scan_time = last_scan or datetime.now()
        self._scan_in_history = True  # loaded from the past-event list, so it has a row
        self._refresh_response_tree(0, self._build_effective_roster())
        self._update_scan_status_banner()
        return True

    def _build_attendance_workbook(self):
        """The Attendance & Payment report for the Export button and the
        Thank-you email's attachment (see reports.attendance_workbook())."""
        return reports.attendance_workbook(self._attendance_roster, self._attendance_figures())

    def _export_attendance_to_excel(self):
        """Xuất bảng Attendance & Payment ra 1 file Excel — CHỈ khi bấm nút
        "📊 Export to Excel" (thay cho nút "Open folder"/"Export Excel
        report" cũ), vì giờ dữ liệu thật sự sống trong database, không tự
        ghi Excel liên tục nữa."""
        event_id = self.var_event_id.get().strip()
        if not event_id:
            messagebox.showwarning("Missing Event ID", "Enter an Event ID on Tab 1 first.")
            return
        if not getattr(self, "_attendance_roster", None):
            messagebox.showwarning("List is empty",
                                    "No one in the list yet — it lists the Yes/Maybe votes scanned on Tab 4.")
            return
        out_path = filedialog.asksaveasfilename(
            title="Export Attendance & Payment to Excel",
            defaultextension=".xlsx",
            initialfile=f"Attendance_Payment_{event_id}.xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not out_path:
            return
        try:
            wb = self._build_attendance_workbook()
            wb.save(out_path)
        except Exception as e:
            messagebox.showerror("Error", f"Couldn't export the report:\n{e}")
            return
        messagebox.showinfo("Exported", f"Saved the attendance & payment report to:\n{out_path}")

    def _calendar_body_args(self):
        """Lấy đúng (event_name, event_date, location, budget) hiện tại từ
        Tab 1 — dùng để build nội dung Appointment body mặc định (Tab 5),
        tránh lặp lại 4 dòng self.var_xxx.get() ở nhiều nơi gọi
        build_calendar_body()."""
        return (
            self.var_event_name.get(),
            get_date_str(self.date_event),
            self.var_location.get(),
            self.var_budget.get(),
        )

    def _refresh_calendar_datetime_display(self):
        """Tự làm mới nội dung Appointment body mặc định theo Event Name/
        Date/Location/Budget mới nhất từ Tab 1 — NHƯNG CHỈ khi người dùng
        CHƯA tự tay sửa nó (tức nội dung đang hiển thị vẫn khớp y hệt bản
        mặc định đã sinh lần gần nhất — self.var_appt_body_default). Nếu
        khác (đã hand-edit), GIỮ NGUYÊN, không ghi đè mất công sửa tay.
        (Trước đây hàm này còn cập nhật 1 dòng chữ hiển thị Event date/
        time riêng ở Tab 5 — dòng đó đã được BỎ theo yêu cầu, xem
        _build_tab_calendar(); phần refresh Appointment body vẫn giữ lại
        vì vẫn cần thiết.)"""
        if hasattr(self, "txt_appt_body") and hasattr(self, "combo_calendar_lang"):
            current_text = self.txt_appt_body.get("1.0", "end").strip()
            if current_text == (getattr(self, "var_appt_body_default", "") or "").strip():
                self._apply_calendar_body_lang()

    def _apply_calendar_body_lang(self):
        lang_label = self.combo_calendar_lang.get() or LANG_LABELS["en"]
        lang_code = LANG_LABEL_TO_CODE.get(lang_label, "en")
        new_body = build_calendar_body(lang_code, *self._calendar_body_args())
        self.var_appt_body_default = new_body
        self.txt_appt_body.delete("1.0", "end")
        self.txt_appt_body.insert("1.0", new_body)

    def _thankyou_body_args(self):
        """Tab 1's event and the table's figures for build_thankyou_body():
        one row per round (label, attendees, collected, paid, remaining) and
        the totals over all rounds - from the same figures as the screen."""
        figures = self._attendance_figures()
        rounds_info = [(f.label, str(f.attendees), format_amount(f.collected),
                        format_amount(f.paid), format_amount(f.remaining)) for f in figures]
        collected, paid, remaining = round_totals(figures)
        return (
            self.var_event_name.get(),
            get_date_str(self.date_event),
            self.var_location.get(),
            str(figures[0].attendees),
            rounds_info,
            format_amount(collected),
            format_amount(paid),
            format_amount(remaining),
        )

    def _refresh_thankyou_body_display(self):
        """Tự làm mới nội dung email cảm ơn mặc định theo Tab 1/Attendance
        mới nhất — CÙNG quy tắc "chỉ ghi đè nếu chưa hand-edit" như
        _refresh_calendar_datetime_display() ở trên, để không mất công sửa
        tay khi chuyển qua chuyển lại giữa các tab."""
        if hasattr(self, "txt_thankyou_body") and hasattr(self, "combo_thankyou_lang"):
            current_text = self.txt_thankyou_body.get("1.0", "end").strip()
            if current_text == (getattr(self, "var_thankyou_body_default", "") or "").strip():
                self._apply_thankyou_body_lang()

    def _update_thankyou_body_from_table(self):
        """"🔄 Update from table": rebuilds the text from the current figures,
        asking first when it was edited by hand."""
        current = self.txt_thankyou_body.get("1.0", "end").strip()
        if current and current != (self.var_thankyou_body_default or "").strip():
            if not messagebox.askyesno(
                    "Update Thank You email",
                    "The email content has been edited by hand.\n\n"
                    "Rebuild it from the current Attendance & Payment table? "
                    "Your manual edits will be lost."):
                return
        self._apply_thankyou_body_lang()

    def _apply_thankyou_body_lang(self):
        lang_label = self.combo_thankyou_lang.get() or LANG_LABELS["en"]
        lang_code = LANG_LABEL_TO_CODE.get(lang_label, "en")
        new_body = build_thankyou_body(lang_code, *self._thankyou_body_args())
        self.var_thankyou_body_default = new_body
        self.txt_thankyou_body.delete("1.0", "end")
        self.txt_thankyou_body.insert("1.0", new_body)

    def _refresh_calendar_yes_list(self):
        # MỚI: lấy CẢ "Yes" LẪN "Maybe" (trước đây CHỈ lấy "Yes") — biến
        # self._yes_emails vẫn giữ nguyên TÊN cũ (để không phải sửa lại các
        # chỗ khác đang dùng, vd _send_calendar()) nhưng giờ nội dung thực
        # sự là "Yes + Maybe".
        self.list_yes.delete(0, "end")
        self._yes_emails = []
        for iid in self.tree_responses.get_children():
            _manual, name, email, vote, _ = self.tree_responses.item(iid, "values")
            if vote in ("Yes", "Maybe"):
                self.list_yes.insert("end", f"{name} <{email}> [{vote}]")
                self._yes_emails.append(email)

    def _send_calendar(self):
        if not self._yes_emails:
            messagebox.showwarning("No Yes/Maybe votes", "No one has voted Yes or Maybe yet, or responses haven't been scanned on Tab 4.")
            return
        if not self._table_belongs_to_tab1_event("calendar invite"):
            return
        try:
            hh, mm = self.var_start_time.get().split(":")
            eh, em = self.var_end_time.get().split(":")
            base = get_date_obj(self.date_event)
            start_dt = base.replace(hour=int(hh), minute=int(mm))
            end_dt = base.replace(hour=int(eh), minute=int(em))
        except Exception:
            messagebox.showerror("Invalid time format", "Enter Start/End time on Tab 1 as HH:MM, e.g. 18:00")
            return
        if end_dt <= start_dt:
            messagebox.showerror("Invalid times", "End time on Tab 1 must be later than Start time.")
            return
        subject = self.var_event_name.get()
        location = self.var_location.get()
        # Use custom appointment body from Tab 5
        body = self.txt_appt_body.get("1.0", "end").strip()
        if not body:
            body = self.var_appt_body_default

        event_id = self.var_event_id.get().strip()
        attendees = list(self._yes_emails)
        history_path = self.history_path.get()

        # Chỉ tìm email UPDATE INVITE để đính kèm nếu History CÓ ghi nhận đã
        # từng gửi update invite cho đúng Event ID này (cột UpdateInviteDate).
        include_update = False
        try:
            for rec in db.load_history(history_path):
                if rec.get("EventID") == event_id:
                    include_update = bool(rec.get("UpdateInviteDate"))
                    break
        except Exception:
            pass  # best-effort — không đọc được History thì coi như chưa có update invite

        def worker():
            try:
                # send_calendar_invite() CHỈ tìm + đính kèm email mời GỐC (luôn
                # tìm) và email UPDATE INVITE (nếu include_update=True), trong
                # CÙNG 1 phiên COM — xem docstring trong outlook_com.py.
                # attached_count = SỐ email đính kèm thành công.
                appt, attached_count = self.outlook.send_calendar_invite(
                    attendees, subject, location, start_dt, end_dt, body,
                    attach_event_id=event_id,
                    include_update=include_update,
                )
                if attached_count:
                    attach_note = (f"\n\n📎 Found and attached {attached_count} email(s) "
                                    f"(original invite" + (" + update invite" if include_update else "")
                                    + f") related to this event.")
                else:
                    attach_note = (
                        "\n\n⚠️ Could not find/attach any confirmation email in Sent "
                        "Items (it may have been moved/deleted, sent from a different account, "
                        "or this EventID has no matching email) — the invite was still created "
                        "without an attachment."
                    )
                # Outlook only OPENED the invite; whether it is sent is up to
                # the user, so History records the opening, not a send.
                history_note = ""
                try:
                    if db.update_event(
                            event_id,
                            {"CalendarSent": "Opened " + datetime.now().strftime("%Y-%m-%d %H:%M")},
                            history_path):
                        self.after(0, self._refresh_history_tree)
                except Exception:
                    history_note = "\n\n⚠️ Couldn't record this in History."
                self.after(0, lambda: messagebox.showinfo(
                    "Done",
                    f"Meeting invite opened for {len(attendees)} people. "
                    "Review it and click Send in Outlook." + attach_note + history_note))
            except Exception as e:
                err_msg = str(e)
                self.after(0, lambda: messagebox.showerror("Error", f"Could not send the meeting invite:\n{err_msg}"))

        threading.Thread(target=worker, daemon=True).start()

    def _send_thank_you_email(self):
        """MỚI: gửi email cảm ơn sau sự kiện tới những ai "Actual Attend" =
        Yes trong bảng Attendance & Payment. Tự đính kèm: (1) báo cáo Excel
        Attendance & Payment (đầy đủ tên/vote/actual attend/free/amount +
        Amount paid/Remaining amount, xem _build_attendance_workbook()),
        và (2) Calendar Invite của sự kiện, tìm best-effort trong folder
        Calendar theo đúng Subject = Event Name (xem
        self.outlook.send_thankyou_email() / outlook_com._find_calendar_
        invite_appointment() — CÙNG file/convention với các nút gửi khác
        của app, không còn tách riêng như bản trước khi có outlook_com.py)."""
        actual_attendees = [
            (info.get("name") or email, email)
            for email, info in getattr(self, "_attendance_roster", {}).items()
            if (info.get("actual_attend") or "").strip().lower() == "yes"
        ]
        if not actual_attendees:
            messagebox.showwarning(
                "No confirmed attendees",
                "No one is marked \"Actual Attend\" = Yes in the table above yet. "
                "Fill in Actual Attend for the people who showed up first."
            )
            return

        body = self.txt_thankyou_body.get("1.0", "end").strip()
        # Text the app generated (not edited by hand) is rebuilt from the
        # table first: it is generated when the page opens, and the table may
        # have changed since. Hand-edited text is sent as it is ("🔄 Update
        # from table" rebuilds it on request).
        if not body or body == (self.var_thankyou_body_default or "").strip():
            self._apply_thankyou_body_lang()
            body = self.txt_thankyou_body.get("1.0", "end").strip()
        if not body:
            messagebox.showwarning(
                "Empty content",
                "Change the language dropdown to regenerate the text, or type the content by hand "
                "before sending.")
            return

        event_id = self.var_event_id.get().strip()
        event_name = self.var_event_name.get()
        lang_label = self.combo_thankyou_lang.get() or LANG_LABELS["en"]
        lang_code = LANG_LABEL_TO_CODE.get(lang_label, "en")
        subject = build_thankyou_subject(lang_code, event_id, event_name)

        # Chuẩn bị file Excel đính kèm TRƯỚC (trong main thread, vì cần
        # đọc self._attendance_roster/các StringVar UI — an toàn hơn đọc
        # chúng từ worker thread), lưu vào 1 file tạm — không phải nơi lưu
        # trữ chính (dữ liệu thật vẫn ở database), chỉ để đính kèm email.
        try:
            wb = self._build_attendance_workbook()
            excel_path = os.path.join(
                tempfile.gettempdir(), f"Attendance_Payment_{event_id or 'event'}.xlsx")
            wb.save(excel_path)
        except Exception as e:
            messagebox.showerror("Error", f"Couldn't prepare the attendance Excel attachment:\n{e}")
            return

        attendee_count = len(actual_attendees)
        # Sent as HTML: Outlook's proportional font breaks the space-aligned
        # table of rounds, so every block of "│" lines becomes a real table.
        html_body = text_body_to_html(body)

        def worker():
            try:
                mail, calendar_attached = self.outlook.send_thankyou_email(
                    actual_attendees, subject, body,
                    excel_path=excel_path, event_name=event_name, html_body=html_body,
                )
                if calendar_attached:
                    attach_note = "\n\n📎 Attached the attendance report and the Calendar Invite."
                else:
                    attach_note = (
                        "\n\n📎 Attached the attendance report.\n\n"
                        "⚠️ Couldn't find the Calendar Invite in your Calendar folder to attach "
                        "(it may not have been sent yet, or the meeting Subject no longer matches "
                        "the Event Name on Tab 1) — the email was still created without it."
                    )
                self.after(0, lambda: messagebox.showinfo(
                    "Done",
                    f"Thank-you email opened for {attendee_count} confirmed attendee(s). "
                    "Review it, then click Send in Outlook." + attach_note))
            except Exception as e:
                err_msg = str(e)
                self.after(0, lambda: messagebox.showerror(
                    "Error", f"Could not create the thank-you email:\n{err_msg}"))

        threading.Thread(target=worker, daemon=True).start()

    def _build_tab_history(self):
        f = self.tab_history.body

        source = Card(f, "Database", "Every page reads and writes this file. Back it up now and then.")
        source.pack(fill="x", pady=(0, 16))
        row = ttk.Frame(source.body)
        row.pack(fill="x")
        ttk.Entry(row, textvariable=self.history_path).pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="📂 Browse...", command=self._browse_history_file).pack(side="left", padx=(8, 0))

        events = Card(f, "Saved events",
                      "Double-click a cell to edit it (Ctrl+C / Ctrl+V work with Excel too), then save: "
                      "only the cells you edited are written. Renaming an Event ID moves the whole "
                      "event — recipients, votes, attendance, rounds and gift list — to the new ID. "
                      "Drag a column heading sideways to reorder the columns; the order is remembered.")
        events.pack(fill="both", expand=True)
        # The database is the working copy; Export writes a snapshot file with
        # every stored column. Nothing is written until "Save changes".
        ttk.Button(events.actions, text="📊 Export to Excel", command=self._export_history_to_excel)\
            .pack(side="right")
        ttk.Button(events.actions, text="🔄 Reload", command=self._refresh_history_tree)\
            .pack(side="right", padx=(0, 8))
        ttk.Button(events.actions, text="💾 Save changes", style="Primary.TButton",
                   command=self._save_history_edits).pack(side="right", padx=(0, 8))

        WrapLabel(events.body, style="Hint.TLabel",
                  text="The money columns are calculated from Attendance & payment and Gift contribution "
                       "and cannot be typed here. Actual Att. (main) counts round 1 only; Income, Expense "
                       "and Balance cover every round and the gift. Dept. Fund Left is the running total "
                       "of Balance from the first event down to that row.").pack(fill="x")
        layout = ttk.Frame(events.body, style="Card.TFrame")
        layout.pack(fill="x", pady=(8, 10))
        ttk.Button(layout, text="↺ Reset order", style="Small.TButton",
                   command=self._reset_history_column_order).pack(side="right")
        ttk.Button(layout, text="↔ Fit columns", style="Small.TButton",
                   command=self._autofit_history_columns).pack(side="right", padx=(0, 6))
        tree_container, self.tree_history = make_scrollable_treeview(
            events.body, columns=HISTORY_TABLE_KEYS, height=16, horizontal=True)
        for c, h, w in HISTORY_TABLE_COLUMNS:
            self.tree_history.heading(c, text=h, anchor="w")
            self.tree_history.column(c, width=w, stretch=False)
        tree_container.pack(fill="both", expand=True)
        self.lbl_history_note = WrapLabel(events.body, style="Hint.TLabel", text="")
        self.lbl_history_note.pack(fill="x", pady=(8, 0))
        # Double-click or paste edits only what is shown; nothing is written
        # until '💾 Save changes'. The calculated columns are left out.
        editable = set(HISTORY_TABLE_KEYS) - HISTORY_READ_ONLY
        self.tree_history.bind(
            "<Double-1>",
            lambda e: self._on_editable_tree_double_click(
                self.tree_history, e, editable, self._commit_history_edit))
        self._enable_treeview_copy_paste(self.tree_history, on_commit=self._commit_history_edit,
                                         editable_cols=editable)
        self._enable_column_drag_reorder(self.tree_history, self._on_history_columns_reordered)
        self._apply_history_column_order()

        self._history_edits = {}    # row iid -> {column: edited value}, not saved yet
        self._history_row_ids = {}  # row iid -> the EventID that row has in the database
        self._history_fund = {}     # row iid -> Dept. Fund Left as shown
        self._refresh_history_tree()

    def _browse_history_file(self):
        path = filedialog.askopenfilename(filetypes=[("SQLite database", "*.db"), ("All files", "*.*")])
        if path:
            self.history_path.set(path)
            self._history_edits = {}
            # The column order is stored in the database file itself.
            self._apply_history_column_order()
            self._refresh_history_tree()

    def _refresh_history_tree(self):
        """Redraws Tab 7 from the database. Edits not saved yet are laid back
        on top, so an automatic write elsewhere (a scan, a reminder) that
        refreshes this table never throws them away. Dept. Fund Left is the
        running total of Balance in History order, computed here and never
        written back: reading the table must not change the database."""
        self.tree_history.delete(*self.tree_history.get_children())
        try:
            records = db.load_history(self.history_path.get())
        except Exception:
            records = []
        fund = running_fund(r.get("Balance") for r in records)
        cols = self.tree_history["columns"]
        self._history_row_ids = {}
        self._history_fund = {}
        unreadable = None
        for rec, total in zip(records, fund):
            # The row's iid is its ORIGINAL Event ID, so an edited Event ID
            # cell can still be traced back to its row in the database.
            base_iid = (rec.get("EventID") or "").strip() or "(blank)"
            iid = base_iid
            suffix = 2
            while self.tree_history.exists(iid):
                iid = f"{base_iid}__{suffix}"
                suffix += 1
            self._history_row_ids[iid] = rec.get("EventID")
            self._history_fund[iid] = "?" if total is None else format_amount(total)
            if total is None and unreadable is None:
                unreadable = rec
            pending = self._history_edits.get(iid, {})
            shown = dict(rec, DeptFundRemaining=self._history_fund[iid])
            values = [pending.get(c, "" if shown.get(c) is None else shown.get(c)) for c in cols]
            self.tree_history.insert("", "end", iid=iid, values=values)
        # Edits of rows that no longer exist cannot be saved anywhere.
        self._history_edits = {iid: e for iid, e in self._history_edits.items()
                               if iid in self._history_row_ids}
        self.lbl_history_note.configure(text=(
            f"Dept. Fund Left shows ? from '{unreadable.get('EventID')}' on: its Balance "
            f"'{unreadable.get('Balance')}' is not a plain amount. Correct it in that copy of the app "
            "or change its money on Attendance & payment." if unreadable else ""))
        self._autofit_history_columns()

    # ── Tab 7: column order and widths ──

    def _load_history_column_order(self):
        """The order the user dragged the columns into, saved in the database
        (app_settings). Healed on the way in: unknown names are dropped (a Tk
        error would leave the table empty) and columns added since are
        appended, so an order saved by the other copy of the app, or an older
        one, still works."""
        try:
            raw = db.get_setting(HISTORY_COLUMN_ORDER_KEY, "", self.history_path.get())
        except Exception:
            raw = ""
        saved = [c.strip() for c in (raw or "").split(",") if c.strip()]
        valid = list(dict.fromkeys(c for c in saved if c in HISTORY_TABLE_KEYS))
        return valid + [c for c in HISTORY_TABLE_KEYS if c not in valid] if valid else list(HISTORY_TABLE_KEYS)

    def _apply_history_column_order(self, order=None, save=False):
        order = list(order) if order else self._load_history_column_order()
        try:
            self.tree_history.configure(displaycolumns=order)
        except tk.TclError:
            self.tree_history.configure(displaycolumns=list(HISTORY_TABLE_KEYS))
            return
        if save:
            try:
                db.set_setting(HISTORY_COLUMN_ORDER_KEY, ",".join(order), self.history_path.get())
            except Exception:
                pass  # a layout preference; the table still works

    def _on_history_columns_reordered(self, order):
        self._apply_history_column_order(order, save=True)

    def _reset_history_column_order(self):
        self._apply_history_column_order(list(HISTORY_TABLE_KEYS), save=True)
        self._autofit_history_columns()

    def _autofit_history_columns(self):
        """Every heading readable, cells as wide as their longest line up to a
        cap - measured with the table's fonts, since Japanese text and emoji
        are wider than their character count."""
        tree = self.tree_history
        try:
            heading_font = tkfont.nametofont(self.fonts.small_medium)
            cell_font = tkfont.nametofont(self.fonts.body)
        except tk.TclError:
            return
        rows = [tree.item(iid, "values") for iid in tree.get_children()[:400]]
        for i, col in enumerate(tree["columns"]):
            need_head = heading_font.measure(HISTORY_TABLE_HEADERS.get(col, col)) + 30
            widest = max((cell_font.measure(line) for values in rows if i < len(values)
                          for line in str(values[i] or "").splitlines()), default=0)
            width = max(need_head, min(widest + 24, 320) if widest else 0, 44)
            tree.column(col, width=width, minwidth=need_head, stretch=False)

    def _enable_column_drag_reorder(self, tree, on_reorder):
        """Drag a heading sideways to move its column (ttk.Treeview cannot).
        Presses on the thin line between headings are left alone - that is
        where Tk resizes a column - and on cells, so selection still works.
        Dropped right of where it started, the column lands after the target;
        dropped left, before it - as in Excel."""
        state = {"col": None}

        def press(event):
            state["col"] = (self._tree_column_name_at(tree, tree.identify_column(event.x))
                            if tree.identify_region(event.x, event.y) == "heading" else None)

        def motion(event):
            if state["col"] is not None:
                tree.configure(cursor="exchange")

        def release(event):
            source, state["col"] = state["col"], None
            tree.configure(cursor="")
            if source is None or tree.identify_region(event.x, event.y) != "heading":
                return
            target = self._tree_column_name_at(tree, tree.identify_column(event.x))
            if not target or target == source:
                return
            order = list(self._tree_display_columns(tree))
            moving_right = order.index(source) < order.index(target)
            order.remove(source)
            order.insert(order.index(target) + (1 if moving_right else 0), source)
            on_reorder(order)

        tree.bind("<ButtonPress-1>", press, add="+")
        tree.bind("<B1-Motion>", motion, add="+")
        tree.bind("<ButtonRelease-1>", release, add="+")

    def _commit_history_edit(self, row_id, col_name, new_value):
        """Called after double-click editing (or pasting into) a cell on Tab
        7 — only updates the table and remembers the edit. Nothing touches
        the database until _save_history_edits() runs."""
        if not self.tree_history.exists(row_id) or col_name in HISTORY_READ_ONLY:
            return
        self.tree_history.set(row_id, col_name, new_value)
        self._history_edits.setdefault(row_id, {})[col_name] = new_value

    def _save_history_edits(self):
        """Writes the cells edited on Tab 7 — only those, so a value written
        meanwhile by the app (vote counts, reminder and calendar times) is
        never reverted by a stale copy of the table. An edited Event ID
        moves the whole event, every column and every per-event table, to
        the new ID (db.rename_event()); an ID that already holds data is
        refused rather than merged."""
        if not self._history_edits:
            messagebox.showinfo("Nothing to save", "No cell has been edited since the last save.")
            return
        path = self.history_path.get()
        saved, renamed, errors = 0, 0, []
        for iid, edits in list(self._history_edits.items()):
            original_id = self._history_row_ids.get(iid)
            target_id = original_id
            try:
                if not original_id:
                    raise ValueError("this row has no Event ID to save it under")
                new_id = str(edits.get("EventID", original_id)).strip()
                fields = {c: v for c, v in edits.items() if c != "EventID"}
                if new_id != original_id:
                    if not db.rename_event(original_id, new_id, path):
                        raise ValueError("the row is no longer in the database")
                    target_id = new_id
                    renamed += 1
                    self._follow_event_rename(original_id, new_id)
                    # The row is drawn under its new ID from now on; keep the
                    # other edits with it in case writing them fails below.
                    del self._history_edits[iid]
                    iid = new_id
                    self._history_edits[iid] = fields
                if fields and not db.update_event(target_id, fields, path):
                    raise ValueError("the row is no longer in the database")
                saved += 1
                self._history_edits.pop(iid, None)
            except Exception as e:
                errors.append(f"{original_id or iid}: {e}")

        self._refresh_history_tree()
        self._refresh_history_combo_values()
        if errors:
            messagebox.showerror(
                "Some rows failed to save",
                f"Saved {saved} row(s), but {len(errors)} failed (their edits are still on screen):\n\n"
                + "\n".join(errors[:10])
            )
        else:
            note = f"\n\n({renamed} Event ID rename(s) applied.)" if renamed else ""
            messagebox.showinfo("Saved", f"Saved the edits of {saved} row(s) to the database.{note}")

    def _follow_event_rename(self, old_id, new_id):
        """After an Event ID is renamed on Tab 7, everything the app holds for
        that event follows it, so later auto-saves land on the renamed event
        instead of re-creating data under the old ID."""
        if self.var_event_id.get().strip() == old_id:
            self.var_event_id.set(new_id)
        for attr in ("_last_scanned_event_id", "_attendance_event", "_gift_event", "_gift_item_event",
                     "_amount_paid_event"):
            if getattr(self, attr) == old_id:
                setattr(self, attr, new_id)
        self._update_scan_status_banner()

    def _export_history_to_excel(self):
        """Exports every row of the Tab 7 table to a new RSVP_History.xlsx-
        style Excel file, with EVERY stored column: the ones shown come from
        the table (edits not saved yet included), the hidden ones from the
        database row. Prompts for a save location so it doesn't silently
        overwrite an old Excel file left over from before the SQLite
        migration."""
        if not self.tree_history.get_children():
            messagebox.showwarning("Nothing to export", "There's no data in the table to export yet.")
            return
        out_path = filedialog.asksaveasfilename(
            title="Export History to Excel",
            defaultextension=".xlsx",
            initialfile="RSVP_History_export.xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not out_path:
            return
        try:
            stored = {r.get("EventID"): r for r in db.load_history(self.history_path.get())}
            shown_cols = self.tree_history["columns"]
            cols = db.EVENT_COLUMNS
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "History"
            reports.write_header_row(ws, cols)
            for r, row_id in enumerate(self.tree_history.get_children(), start=2):
                shown = dict(zip(shown_cols, self.tree_history.item(row_id, "values")))
                record = stored.get(self._history_row_ids.get(row_id), {})
                # Dept. Fund Left as computed for the table, not a stale copy.
                record = dict(record, DeptFundRemaining=self._history_fund.get(row_id, ""))
                for c_idx, c in enumerate(cols, start=1):
                    ws.cell(row=r, column=c_idx, value=shown[c] if c in shown else record.get(c))
            for i in range(1, len(cols) + 1):
                ws.column_dimensions[get_column_letter(i)].width = 16
            wb.save(out_path)
        except Exception as e:
            messagebox.showerror("Error", f"Couldn't export to Excel:\n{e}")
            return
        messagebox.showinfo("Exported", f"Saved a snapshot of the History table to:\n{out_path}")


if __name__ == "__main__":
    app = RSVPApp()
    app.mainloop()

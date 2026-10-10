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
from tkinter import ttk, filedialog, messagebox

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

APP_TITLE = "Outlook RSVP Tool"

# Tab 7 (Event History): (column, header, width), in db.EVENT_COLUMNS order.
# Not shown, because nothing writes them any more: the cost columns of the
# removed "Actual cost tracking" (ActualAttendees, CostPerPerson,
# TotalIncome, TotalExpense, Balance), ReminderSent (only ever set for a
# scheduled-reminder script that does not exist) and ReportFile (only set by
# a report export that had no button). Their stored values are kept, and
# "Export to Excel" still writes them. AmountPaid is edited on Tab 5.
HISTORY_TABLE_COLUMNS = [
    ("EventID", "Event ID", 110), ("EventName", "Event Name", 170),
    ("EventDate", "Date", 85), ("Deadline", "Deadline", 85),
    ("Location", "Location", 110), ("Budget", "Budget", 90),
    ("EmailLanguage", "Language", 140), ("OrganizerNote", "Organizer Note", 220),
    ("RecipientFile", "Recipient File", 220), ("SentDate", "Sent Date", 120),
    ("UpdateInviteDate", "Update Invite Date", 130),
    ("TotalInvited", "Invited", 65), ("Yes", "Yes", 45), ("No", "No", 45),
    ("Maybe", "Maybe", 55), ("NoResponse", "No Resp.", 70),
    ("CalendarSent", "Calendar Invite", 140),
    ("LastReminderSentDate", "Last Reminder Sent", 140),
    ("EventMode", "Event Mode", 90), ("Organizer", "Organizer", 140),
    ("GuestOfHonor", "Guest of Honor", 140), ("GiftBudget", "Gift Budget", 100),
    ("GiftDeadline", "Gift Deadline", 90), ("StartTime", "Start", 65), ("EndTime", "End", 65),
]

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
    count_actual_attendees,
    merge_expanded_roster,
    format_amount,
    parse_amount_from_text,
    remaining_amount,
    sum_contributions,
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
    cleanup_pasted_translation,
    dedupe_pasted_translation,
    detect_possible_duplicate_paste,
)


# ══════════════════════════════════════════════════════════════════════════
# Date-picker helper widget
# ══════════════════════════════════════════════════════════════════════════
def make_date_picker(parent, initial=None):
    if HAS_TKCALENDAR:
        w = DateEntry(parent, date_pattern="dd/mm/yyyy", width=14)
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
# Scrollable tab container — every tab's real content goes inside `.body`,
# so long tabs get a vertical (and horizontal, if needed) scrollbar instead
# of being cut off on smaller screens / windows.
# ══════════════════════════════════════════════════════════════════════════
class ScrollableFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        # BUG ĐÃ SỬA: tk.Canvas KHÔNG tự lấy màu nền theo theme ttk ("clam")
        # — mặc định nó dùng màu trắng thuần của hệ thống, khác hẳn màu nền
        # xám/xanh nhạt mà mọi ttk.Frame/ttk.Label khác trong app đang dùng.
        # Kết quả: bất cứ khoảng trống nào trong vùng cuộn (canvas rộng hơn
        # nội dung thật bên trong self.body) đều lộ ra 1 mảng trắng lệch
        # tông so với phần nền chứa nút bấm/nhãn thông báo xung quanh. Lấy
        # đúng màu nền mà style ttk "TFrame" đang dùng (fallback về màu xám
        # nhạt tiêu chuẩn của theme "clam" nếu vì lý do gì đó chưa tra được)
        # rồi gán thẳng cho canvas để 2 vùng luôn cùng 1 màu.
        style = ttk.Style(self)
        bg_color = style.lookup("TFrame", "background") or "#dcdad5"
        canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0, background=bg_color)
        vscroll = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        hscroll = ttk.Scrollbar(self, orient="horizontal", command=canvas.xview)
        self.body = ttk.Frame(canvas)

        self.body.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas_window = canvas.create_window((0, 0), window=self.body, anchor="nw")

        # BUG ĐÃ SỬA: self.body trước đây được đặt vào canvas với kích
        # thước CỐ ĐỊNH bằng đúng nội dung của nó lúc tạo — khi cửa sổ app
        # được kéo RỘNG hơn, canvas nới rộng ra theo nhưng self.body (và do
        # đó mọi nhãn thông báo/nút bấm bên trong) KHÔNG nới theo, để lại 1
        # dải trống bên phải/dưới, đúng như hiện tượng "phần thông báo
        # không tự động fix theo" khi resize. Giờ mỗi khi canvas đổi kích
        # thước, ép item cửa sổ bên trong (self.body) rộng bằng đúng canvas
        # — nội dung sẽ luôn lấp đầy chiều ngang, và các nhãn word-wrap
        # (_make_wrapping_label) tính lại đúng bề rộng thật để xuống dòng.
        def _on_canvas_configure(event):
            canvas.itemconfig(canvas_window, width=event.width)
            # MỚI: cũng tính lại wraplength cho MỌI label đã đăng ký vào
            # ĐÚNG canvas này (xem RSVPApp._make_wrapping_label() bên dưới)
            # — dùng TRỰC TIẾP chiều rộng THẬT của canvas, thay vì suy luận
            # gián tiếp từ chiều rộng cửa sổ gốc (self.winfo_width()) như
            # cách cũ. 2 con số đó lệch nhau khá nhiều (thanh cuộn dọc,
            # padding của Notebook, indent riêng của từng frame con...),
            # khiến 1 số label tính sai độ rộng cần xuống dòng và không co
            # lại đúng khi cửa sổ bị thu nhỏ — đây là NGUYÊN NHÂN gốc của
            # lỗi "nội dung hướng dẫn chưa fit theo khi resize". Giờ mỗi
            # canvas tự quản lý danh sách label của riêng nó, luôn khớp
            # đúng chiều rộng thực tế đang hiển thị của ĐÚNG tab đó.
            for lbl, margin in list(getattr(canvas, "wrap_labels", [])):
                try:
                    lbl.configure(wraplength=max(event.width - margin, 250))
                except tk.TclError:
                    pass  # label đã bị huỷ — bỏ qua an toàn
        canvas.bind("<Configure>", _on_canvas_configure)

        canvas.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)

        canvas.grid(row=0, column=0, sticky="nsew")
        vscroll.grid(row=0, column=1, sticky="ns")
        hscroll.grid(row=1, column=0, sticky="ew")
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


def make_scrollable_treeview(parent, columns, height=10, show="headings"):
    """Creates a Treeview WITH its scrollbar, both inside a small container frame.
    Returns (container, tree). Grid/pack the returned CONTAINER — never the tree itself —
    since the tree's actual parent is the container, not the outer `parent` passed in."""
    container = ttk.Frame(parent)
    tree = ttk.Treeview(container, columns=columns, show=show, height=height)
    vscroll = ttk.Scrollbar(container, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vscroll.set)
    tree.pack(side="left", fill="both", expand=True)
    vscroll.pack(side="right", fill="y")
    return container, tree


def make_scrollable_text(parent, **text_kwargs):
    """Creates a Text widget WITH its scrollbar, both inside a small container frame.
    Returns (container, text_widget). Grid/pack the returned CONTAINER — never the text
    widget itself — since the text widget's actual parent is the container."""
    container = ttk.Frame(parent)
    text_widget = tk.Text(container, **text_kwargs)
    vscroll = ttk.Scrollbar(container, orient="vertical", command=text_widget.yview)
    text_widget.configure(yscrollcommand=vscroll.set, wrap="word")
    text_widget.pack(side="left", fill="both", expand=True)
    vscroll.pack(side="right", fill="y")
    return container, text_widget


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
        self.geometry("1020x720")

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

        # ── Notebook tab styling ──
        # By default, ttk uses the OS theme's tab colors — on Windows this
        # is usually "vista"/"xpnative", which draws the Notebook tab
        # background NATIVELY and ignores our custom "background" color,
        # while still respecting our custom "foreground" (text) color. That
        # mismatch is exactly what caused the bug report: the selected
        # tab's background stayed the native white/light color while its
        # text was set to white, making the label invisible. Switching to
        # the "clam" theme (a pure-Tk theme, not OS-native) makes ttk fully
        # respect our style customization, so both background AND
        # foreground actually render as configured below. This does change
        # the look of other widgets slightly too (buttons, checkboxes,
        # etc.) to the "clam" style, but keeps the app fully usable and
        # fixes the readability bug for good, instead of just tweaking
        # colors that the old theme might ignore again.
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except Exception:
            pass  # "clam" should always be available, but fall back safely if not
        style.configure(
            "TNotebook.Tab",
            padding=(14, 8),
            font=("Arial", 10, "bold"),
            background="#D9E2F3",
            foreground="#003366",
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", "#003366"), ("active", "#B7C9EB")],
            foreground=[("selected", "#FFFFFF"), ("active", "#003366")],
        )

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=8, pady=8)
        self.nb = nb

        self.tab_config = ScrollableFrame(nb)
        self.tab_recipients = ScrollableFrame(nb)
        self.tab_compose = ScrollableFrame(nb)
        self.tab_collect = ScrollableFrame(nb)
        self.tab_calendar = ScrollableFrame(nb)
        self.tab_gift = ScrollableFrame(nb)  # MỚI — theo dõi quyên góp quà tặng
        self.tab_history = ScrollableFrame(nb)

        nb.add(self.tab_config, text="  1. Event Setup  ")
        nb.add(self.tab_recipients, text="  2. Recipients  ")
        nb.add(self.tab_compose, text="  3. Compose & Send  ")
        nb.add(self.tab_collect, text="  4. Collect Responses  ")
        nb.add(self.tab_calendar, text="  5. Attendance & Payment  ")
        nb.add(self.tab_gift, text="  6. Gift Contribution  ")
        nb.add(self.tab_history, text="  7. Event History  ")

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

    def _on_tab_changed(self, event):
        try:
            selected = event.widget.nametowidget(event.widget.select())
        except Exception:
            return
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
        col = tree.identify_column(event.x)  # e.g. '#1', '#2', ...
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
        if editable_cols is not None and col_name not in editable_cols:
            return
        self._begin_cell_edit(tree, row_id, col_name, on_commit)

    def _make_wrapping_label(self, parent, margin=60, **kwargs):
        """Creates a ttk.Label that automatically re-wraps its text (via
        wraplength) whenever ITS OWN TAB is resized — used for long
        instructional/help text that would otherwise get clipped, or
        trigger the horizontal scrollbar, when the window is narrowed.
        `margin` is roughly how much horizontal space (in pixels) to
        reserve for padding/indentation around the label — bump it up a
        bit for labels nested further to the right. Usage is a drop-in
        replacement for ttk.Label(...) — chain .grid(...)/.pack(...) on
        the result exactly the same way:
            self._make_wrapping_label(f, text="...", font=(...)).grid(row=r, ...)

        BUG ĐÃ SỬA: bản trước tính wraplength dựa theo self.winfo_width()
        (chiều rộng của CỬA SỔ GỐC toàn app) — con số này luôn RỘNG HƠN
        vùng nội dung THẬT SỰ nhìn thấy được bên trong 1 tab cụ thể (mất đi
        vì thanh cuộn dọc, padding của Notebook, indent riêng của từng
        frame con lồng nhau...), nên nhiều label không co lại đúng khi cửa
        sổ bị thu nhỏ. Giờ dò ngược lên cây widget cha để tìm đúng
        tk.Canvas của ScrollableFrame đang bọc tab chứa label này, và đăng
        ký vào DANH SÁCH RIÊNG của canvas đó (canvas.wrap_labels, xem
        ScrollableFrame.__init__) — mỗi khi CHÍNH canvas đó đổi kích thước
        (khớp chính xác vùng hiển thị thật), wraplength được tính lại theo
        đúng con số đó, không qua trung gian nào khác."""
        lbl = ttk.Label(parent, **kwargs)
        # Every tab is a ScrollableFrame, so every label has a Canvas ancestor.
        w = parent
        while w is not None and not isinstance(w, tk.Canvas):
            w = w.master
        if w is not None:
            if not hasattr(w, "wrap_labels"):
                w.wrap_labels = []
            w.wrap_labels.append((lbl, margin))
            try:
                current_width = w.winfo_width()
                if current_width > 1:
                    lbl.configure(wraplength=max(current_width - margin, 250))
            except Exception:
                pass
        return lbl

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
            lines = []
            for row_id in rows:
                values = tree.item(row_id, "values")
                lines.append("\t".join("" if v is None else str(v) for v in values))
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
            columns = tree["columns"]
            pasted_lines = [ln for ln in text.replace("\r\n", "\n").replace("\r", "\n").split("\n") if ln != ""]
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
        pad = {"padx": 10, "pady": 6}

        header_bar = ttk.Frame(f)
        header_bar.grid(row=0, column=0, columnspan=3, sticky="we", **pad)
        ttk.Label(header_bar, text="Event details (these fields auto-fill into the invite email)",
                  font=("Arial", 11, "bold")).pack(side="left")
        # MỚI: "Event mode" — chọn sự kiện này thuộc dạng nào:
        #  - "Event": chỉ RSVP bình thường (mặc định, hành vi y hệt trước giờ)
        #  - "Gift": CHỈ thông báo kêu gọi đóng góp mua quà tặng (Tab 3 có
        #    thêm mode "Send Gift Contribution Notice"), không cần RSVP
        #  - "Event + Gift": vừa RSVP vừa có kêu gọi đóng góp quà — dùng khi
        #    cùng lúc tổ chức tiệc (cần biết số người tham dự) VÀ muốn quyên
        #    góp mua quà tặng (vd tiệc farewell tặng quà người sắp nghỉ việc)
        # Bản thân "Event mode" chỉ là NHÃN LƯU Ý — không tự động khoá/ẩn
        # tính năng nào cả, bạn vẫn luôn có thể dùng bất kỳ Send mode nào ở
        # Tab 3 bất kể chọn gì ở đây; nó giúp bạn nhớ lại mục đích sự kiện
        # khi xem lại History sau này (cột "EventMode").
        ttk.Label(header_bar, text="   Event mode:", font=("Arial", 10, "bold")).pack(side="left")
        self.var_event_mode = tk.StringVar(value="Event")
        self.combo_event_mode = ttk.Combobox(
            header_bar, width=16, state="readonly", textvariable=self.var_event_mode,
            values=["Event", "Gift", "Event + Gift"],
        )
        self.combo_event_mode.pack(side="left", padx=(4, 0))

        self.var_event_name = tk.StringVar(value="Team Building Q3 2026")
        self.var_event_id = tk.StringVar(value="TB2026-Q3")
        # Mỗi khi Event ID đổi (gõ tay hoặc do Load setup/Register event...),
        # tự cập nhật banner trạng thái ở đầu Tab 4 — báo ngay nếu bảng kết
        # quả đang hiện là của 1 Event ID KHÁC (còn sót từ lần Scan trước).
        self.var_event_id.trace_add(
            "write", lambda *a: self._update_scan_status_banner() if hasattr(self, "lbl_scan_status") else None)
        self.var_location = tk.StringVar(value="")
        self.var_budget = tk.StringVar(value="")
        self.var_gift_budget = tk.StringVar(value="")
        self.var_organizer = tk.StringVar(value="")
        self.var_guest_of_honor = tk.StringVar(value="")

        r = 1
        ttk.Label(f, text="Event Name:").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(f, textvariable=self.var_event_name, width=50).grid(row=r, column=1, sticky="w", **pad)
        r += 1
        ttk.Label(f, text="Event ID (no spaces/accents, must be unique):").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(f, textvariable=self.var_event_id, width=50).grid(row=r, column=1, sticky="w", **pad)
        r += 1
        ttk.Label(f, text="Location:").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(f, textvariable=self.var_location, width=50).grid(row=r, column=1, sticky="w", **pad)
        r += 1
        # Đổi tên "Expected Budget" -> "Expected EVENT Budget" để phân biệt rõ
        # với "Expected GIFT Budget" mới thêm ngay bên dưới (2 khoản chi khác
        # nhau: ngân sách tổ chức sự kiện vs ngân sách mua quà tặng).
        ttk.Label(f, text="Expected event budget (e.g. \"5,000 JPY / person\"):").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(f, textvariable=self.var_budget, width=50).grid(row=r, column=1, sticky="w", **pad)
        r += 1
        ttk.Label(f, text="Expected gift budget (e.g. \"3,000 JPY / person\"):").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(f, textvariable=self.var_gift_budget, width=50).grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(f, text="Event Date:").grid(row=r, column=0, sticky="w", **pad)
        self.date_event = make_date_picker(f, datetime.now() + timedelta(days=21))
        self.date_event.grid(row=r, column=1, sticky="w", **pad)
        r += 1

        # MỚI: Start time / End time CHUYỂN TỪ TAB 5 LÊN ĐÂY (ngay dưới Event
        # Date) — trước đây 2 ô này chỉ tồn tại ở Tab 5 (Attendance & Payment) và
        # KHÔNG được lưu vào History, nên "Load setup from selected event"
        # không nạp lại được, mỗi lần mở Tab 5 lại phải gõ tay lại giờ. Giờ
        # thuộc Tab 1 nên được lưu/nạp cùng các thông tin sự kiện khác, và
        # Tab 5 dùng lại TRỰC TIẾP 2 biến này (self.var_start_time/
        # self.var_end_time) khi tạo Calendar Invite — không còn ô riêng ở
        # Tab 5 nữa.
        ttk.Label(f, text="Start time (HH:MM):").grid(row=r, column=0, sticky="w", **pad)
        self.var_start_time = tk.StringVar(value="18:00")
        ttk.Entry(f, textvariable=self.var_start_time, width=10).grid(row=r, column=1, sticky="w", **pad)
        r += 1
        ttk.Label(f, text="End time (HH:MM):").grid(row=r, column=0, sticky="w", **pad)
        self.var_end_time = tk.StringVar(value="21:00")
        ttk.Entry(f, textvariable=self.var_end_time, width=10).grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(f, text="Event response deadline:").grid(row=r, column=0, sticky="w", **pad)
        self.date_deadline = make_date_picker(f, datetime.now() + timedelta(days=10))
        self.date_deadline.grid(row=r, column=1, sticky="w", **pad)
        r += 1

        # MỚI: "Gift contribution deadline" — hạn RIÊNG để mọi người đóng
        # góp tiền mua quà, khác với "Event response deadline" (hạn RSVP có
        # tham dự hay không) ở trên — 2 việc này thường không cùng ngày (vd
        # cần chốt RSVP sớm hơn để chuẩn bị tiệc, nhưng có thể gia hạn thời
        # gian đóng góp quà lâu hơn 1 chút). Dùng cho email "Send Gift
        # Contribution Notice" ở Tab 3 (xem build_gift_fixed_block()).
        ttk.Label(f, text="Gift contribution deadline:").grid(row=r, column=0, sticky="w", **pad)
        self.date_gift_deadline = make_date_picker(f, datetime.now() + timedelta(days=10))
        self.date_gift_deadline.grid(row=r, column=1, sticky="w", **pad)
        r += 1

        # MỚI: Organizer / Guest of Honor — dùng cho email "Send Gift
        # Contribution Notice" ở Tab 3 (Organizer = người nhận đóng góp,
        # Guest of Honor = người được tặng quà, vd người sắp nghỉ việc) —
        # nhưng vẫn hiển thị luôn ở Event mode "Event" (không bị ẩn) vì đôi
        # khi hữu ích để ghi chú ai là người tổ chức/nhân vật chính dù
        # không quyên góp quà, không bắt buộc phải điền.
        ttk.Label(f, text="Organizer (contact person for gift contributions):").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(f, textvariable=self.var_organizer, width=50).grid(row=r, column=1, sticky="w", **pad)
        r += 1
        ttk.Label(f, text="Guest of Honor (who the gift is for):").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(f, textvariable=self.var_guest_of_honor, width=50).grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Label(f, text="Organizer note (free text — you can write in any language;\n"
                          "translate it per-language on the Compose tab):").grid(row=r, column=0, sticky="nw", **pad)
        self.entry_note = tk.Text(f, width=55, height=3)
        self.entry_note.insert("1.0", "Please respond before the deadline so we can prepare accurate numbers.")
        self.entry_note.grid(row=r, column=1, sticky="w", **pad)
        r += 1

        ttk.Button(f, text="📝 Register event (save to History + reset form for a NEW event)",
                   command=self._register_event).grid(row=r, column=0, columnspan=2, sticky="w", **pad)
        r += 1
        ttk.Label(f, text="→ Saves as a NEW event in History, then clears this form for the next one.",
                  font=("Arial", 8, "italic")).grid(row=r, column=0, columnspan=2, sticky="w", padx=4)
        r += 1

        ttk.Separator(f).grid(row=r, column=0, columnspan=3, sticky="ew", pady=10)
        r += 1

        ttk.Label(f, text="Reload setup from a past event saved in history:",
                  font=("Arial", 10, "italic")).grid(row=r, column=0, columnspan=2, sticky="w", **pad)
        r += 1
        # postcommand re-reads the database each time the list opens, so an
        # event saved from any tab shows up without a manual refresh.
        self.combo_load_history = ttk.Combobox(f, width=50, state="readonly",
                                               postcommand=self._refresh_history_combo_values)
        self.combo_load_history.grid(row=r, column=1, sticky="w", **pad)
        ttk.Button(f, text="⬅ Load setup from selected event", command=self._load_from_history)\
            .grid(row=r, column=0, sticky="w", **pad)
        r += 1

        ttk.Separator(f).grid(row=r, column=0, columnspan=3, sticky="ew", pady=10)
        r += 1
        self._make_wrapping_label(f, text="Already sent invites for this Event ID, but the details above changed since "
                          "(location/date/deadline/budget/note)? Push the update into History directly —\n"
                          "no email needs to be sent for this. (To also NOTIFY people by email, use "
                          "'Send update invite' mode on Tab 3 instead — that sends AND updates History.)",
                  font=("Arial", 9, "italic")).grid(row=r, column=0, columnspan=3, sticky="w", **pad)
        r += 1
        ttk.Button(f, text="💾 Update this event in History (no email sent)", command=self._update_history_from_tab1)\
            .grid(row=r, column=0, columnspan=2, sticky="w", **pad)

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
        try:
            db.save_event_record(record, self.history_path.get())
        except Exception as e:
            messagebox.showerror("Error", f"Couldn't save the event:\n{e}")
            return None
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
            messagebox.showwarning("Missing Event ID", "Enter an Event ID before registering the event.")
            return

        if not messagebox.askyesno(
            "Register new event",
            f"Register event '{event_id}' to the database?\n\n"
            "After registering, the WHOLE form (Tab 1 → 6: event info, recipient list, "
            "composed content and translations, scanned results...) will be CLEARED so you can "
            "start entering a NEW event. The event you just registered is still safely stored in "
            "History, and can be reloaded any time via '⬅ Load setup from selected event'."
        ):
            return

        if not self._save_event_from_tab1():
            return
        self._reset_for_new_event()
        messagebox.showinfo(
            "Registered",
            f"Event '{event_id}' registered to History.\n\n"
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
        self.lbl_deadline_banner.config(text="")
        self.txt_reminder_body.delete("1.0", "end")

        # Tab 5 — the two bodies are regenerated from Tab 1 when the tab opens
        self._yes_emails = []
        self.list_yes.delete(0, "end")
        self._attendance_event = None
        self._attendance_roster = {}
        self._render_attendance_tree()
        self._amount_paid_event = None
        self._set_amount_paid_quietly("0")
        self.combo_calendar_lang.current(0)
        self.var_appt_body_default = ""
        self.txt_appt_body.delete("1.0", "end")
        self.combo_thankyou_lang.current(0)
        self.var_thankyou_body_default = ""
        self.txt_thankyou_body.delete("1.0", "end")

        # Tab 6
        self._gift_event = None
        self._gift_roster = {}
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
        """Re-reads the past-event list from the database, keeping the
        current selection. Runs every time the dropdown opens."""
        try:
            records = db.load_history(self.history_path.get())
        except Exception:
            records = []
        selected = self.combo_load_history.get()
        self._history_records = records
        labels = [f'{r["EventID"]} — {r["EventName"]}' for r in records]
        self.combo_load_history["values"] = labels
        if selected in labels:
            self.combo_load_history.current(labels.index(selected))
        else:
            self.combo_load_history.set("")
        return labels

    def _refresh_history_combo(self):
        """Like _refresh_history_combo_values(), then selects the newest event."""
        labels = self._refresh_history_combo_values()
        if labels:
            self.combo_load_history.current(len(labels) - 1)

    def _unsaved_work_if_switching(self):
        """What would be lost by replacing the current event in memory — work
        that exists in no database table, or in one the past-event list can
        never reach again."""
        lost = []
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

    def _event_in_history(self, event_id):
        try:
            return any(r.get("EventID") == event_id for r in db.load_history(self.history_path.get()))
        except Exception:
            return False

    def _load_from_history(self):
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
        top = ttk.Frame(f)
        top.pack(fill="x", padx=10, pady=8)

        # The list is saved to the database automatically after every change
        # (Import/Add/Delete row/Expand group); Excel is only an import source
        # and an export format.
        ttk.Button(top, text="📥 Import from Excel...", command=self._load_recipients).pack(side="left")
        ttk.Button(top, text="📊 Export to Excel", command=self._export_recipients_to_excel).pack(side="left", padx=6)
        ttk.Label(top, textvariable=self.recipient_file, font=("Arial", 8, "italic")).pack(side="left", padx=6)

        expand_bar = ttk.Frame(f)
        expand_bar.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Button(expand_bar, text="🔎 Expand group emails in list (incl. sub-groups)",
                   command=self._expand_group_recipients).pack(side="left")
        # margin: the button to its left takes ~380px of the row.
        self._make_wrapping_label(expand_bar, margin=420, text="  → Any row that is a company Distribution List (group email, "
                                   "e.g. 'EET Employees All') gets replaced with its REAL individual "
                                   "members, so Tab 4 can track exactly who hasn't responded yet.",
                  font=("Arial", 8, "italic")).pack(side="left")

        # MỚI: ô tìm kiếm nhanh theo Name/Email — lọc TRỰC TIẾP bảng bên dưới
        # khi gõ (không cần bấm nút), khớp theo dạng "chứa chuỗi" (không phân
        # biệt hoa/thường), tìm trên CẢ Name lẫn Email cùng lúc.
        search_bar = ttk.Frame(f)
        search_bar.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(search_bar, text="🔎 Search (name or email):").pack(side="left")
        self.var_recipient_search = tk.StringVar(value="")
        search_entry = ttk.Entry(search_bar, textvariable=self.var_recipient_search, width=40)
        search_entry.pack(side="left", padx=6)
        self.var_recipient_search.trace_add("write", lambda *a: self._apply_recipient_filter())
        ttk.Button(search_bar, text="✕ Clear", command=lambda: self.var_recipient_search.set(""))\
            .pack(side="left", padx=4)

        cols = ("name", "email")
        tree_container, self.tree_recipients = make_scrollable_treeview(f, columns=cols, height=18)
        self.tree_recipients.heading("name", text="Name")
        self.tree_recipients.heading("email", text="Email")
        self.tree_recipients.column("name", width=280)
        self.tree_recipients.column("email", width=320)
        tree_container.pack(fill="both", expand=True, padx=10, pady=6)

        edit_bar = ttk.Frame(f)
        edit_bar.pack(fill="x", padx=10, pady=6)
        self.var_new_name = tk.StringVar()
        self.var_new_email = tk.StringVar()
        ttk.Label(edit_bar, text="Name:").pack(side="left")
        ttk.Entry(edit_bar, textvariable=self.var_new_name, width=28).pack(side="left", padx=3)
        ttk.Label(edit_bar, text="Email:").pack(side="left", padx=(6, 0))
        ttk.Entry(edit_bar, textvariable=self.var_new_email, width=28).pack(side="left", padx=3)
        ttk.Button(edit_bar, text="➕ Add person", command=self._add_recipient_row).pack(side="left", padx=6)
        ttk.Button(edit_bar, text="🗑 Delete selected row", command=self._delete_recipient_row).pack(side="left", padx=6)

        self.lbl_recipient_count = ttk.Label(f, text="No list loaded yet.")
        self.lbl_recipient_count.pack(anchor="w", padx=10, pady=4)

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
            errors = []
            for name, email in original:
                try:
                    members = self.outlook.expand_group_members(email)
                except Exception as e:
                    members = None
                    errors.append(f"{email}: {e}")
                if members is None:
                    # Không phải group (hoặc không resolve được) -> giữ nguyên dòng gốc
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
        pad = {"padx": 10, "pady": 5}

        # ── send mode selector (TOP-MOST — important choice, made first) ──
        send_mode_bar = ttk.Frame(f, relief="ridge", borderwidth=1)
        send_mode_bar.grid(row=0, column=0, columnspan=2, sticky="we", padx=10, pady=(8, 4))
        ttk.Label(send_mode_bar, text="📤 Send mode:", font=("Arial", 11, "bold"))\
            .pack(side="left", padx=(8, 6), pady=6)
        self.combo_send_mode = ttk.Combobox(
            send_mode_bar, width=28, state="readonly",
            values=["Send first Invite", "Send update invite", "Send Gift Contribution Notice"],
        )
        self.combo_send_mode.current(0)  # default: first invite
        self.combo_send_mode.pack(side="left", padx=4, pady=6)
        self.combo_send_mode.bind("<<ComboboxSelected>>", lambda e: (self._refresh_send_button_label(), self._refresh_compose_preview()))
        self._make_wrapping_label(send_mode_bar,
                  text="  ↳ 'Update invite' adds a change-notice banner + different subject prefix, but\n"
                       "still lets people re-vote. 'Gift Contribution Notice' sends a Guest-of-Honor/\n"
                       "Organizer/Budget notice from Tab 1 — NO voting buttons (see 'Event mode' on Tab 1).",
                  font=("Arial", 8, "italic")).pack(side="left", padx=4, pady=6)

        # ── language selector ──
        ttk.Label(f, text="Email language:", font=("Arial", 10, "bold"))\
            .grid(row=1, column=0, sticky="w", **pad)
        self.combo_email_lang = ttk.Combobox(
            f, width=32, state="readonly",
            values=[LANG_LABELS["en"], LANG_LABELS["ja"], LANG_LABELS["vi"], LANG_LABELS["bilingual"]],
        )
        self.combo_email_lang.current(0)  # default: English
        self.combo_email_lang.grid(row=1, column=1, sticky="w", **pad)
        self.combo_email_lang.bind("<<ComboboxSelected>>", lambda e: self._refresh_compose_preview())

        self.var_subject_preview = tk.StringVar(value="")
        ttk.Label(f, textvariable=self.var_subject_preview, font=("Arial", 9, "italic"))\
            .grid(row=2, column=0, columnspan=2, sticky="w", padx=10)

        self.var_greeting_preview = tk.StringVar(value="")
        ttk.Label(f, textvariable=self.var_greeting_preview, font=("Arial", 9, "italic"))\
            .grid(row=3, column=0, columnspan=2, sticky="w", padx=10, pady=(2, 8))

        # Email is assembled in THIS order: Greeting (above) → EDITABLE → FIXED.
        # ── editable block preview (organizer's free-text background note, shown FIRST) ──
        ttk.Label(f, text="🟩 EDITABLE part — YOUR background notes for this specific event, shown\n"
                          "right after the greeting (write in any language; translate below if needed):",
                  font=("Arial", 10, "bold"), foreground="#00703C").grid(row=4, column=0, columnspan=2, sticky="w", **pad)
        editable_container, self.txt_editable_preview = make_scrollable_text(f, width=100, height=5, bg="#e8f5e9")
        editable_container.grid(row=5, column=0, columnspan=2, sticky="w", padx=10)

        # ── fixed block preview (event details from Tab 1 + voting instructions, shown AFTER editable) ──
        ttk.Label(f, text="🟧 FIXED part — auto-filled from Tab 1 (event name/date/location/deadline/budget)\n"
                          "plus the standard voting wording, shown AFTER your note. You CAN edit this box —\n"
                          "click 'Save as default' below to reuse your edited wording for future events too:",
                  font=("Arial", 10, "bold"), foreground="#8a4b00").grid(row=6, column=0, columnspan=2, sticky="w", **pad)
        fixed_container, self.txt_fixed_preview = make_scrollable_text(f, width=100, height=9, bg="#fff3e0")
        fixed_container.grid(row=7, column=0, columnspan=2, sticky="w", padx=10)

        fixed_btns = ttk.Frame(f)
        fixed_btns.grid(row=8, column=0, columnspan=2, sticky="w", padx=10, pady=4)
        ttk.Button(fixed_btns, text="🔄 Refresh preview from Tab 1 / Tab 2", command=self._refresh_compose_preview)\
            .pack(side="left", padx=(0, 6))
        ttk.Button(fixed_btns, text="💾 Save FIXED wording as default for this language", command=self._save_fixed_default)\
            .pack(side="left", padx=6)
        ttk.Button(fixed_btns, text="↺ Reset FIXED wording to system default", command=self._reset_fixed_default)\
            .pack(side="left", padx=6)

        # ── prompt customization (NEW) ──
        ttk.Separator(f).grid(row=9, column=0, columnspan=2, sticky="ew", pady=8)
        # Set once and rarely touched, so it starts collapsed: the send
        # controls below are a long scroll away already.
        prompt_header = ttk.Frame(f)
        prompt_header.grid(row=10, column=0, columnspan=2, sticky="w", **pad)
        ttk.Label(prompt_header, text="🎨 Copilot prompt (for single-language targets — icons already built in)",
                  font=("Arial", 10, "bold")).pack(side="left")
        self.var_prompt_toggle = tk.StringVar(value="▸ Show / edit")
        ttk.Button(prompt_header, textvariable=self.var_prompt_toggle,
                   command=self._toggle_prompt_editor).pack(side="left", padx=8)

        self.prompt_editor = ttk.Frame(f)
        self.prompt_editor.grid(row=11, column=0, columnspan=2, sticky="w")
        self._make_wrapping_label(self.prompt_editor, text="This box shows the exact prompt that will be sent to Copilot (system default,\n"
                          "already includes emoji-icon instructions ⏰📍💰📋👥). Edit it freely and click\n"
                          "'Save' to keep your version — it persists until you Reset. The bilingual target\n"
                          "uses its own built-in default (see 'Show system default prompt' below).",
                  font=("Arial", 9, "italic")).pack(anchor="w", padx=10, pady=(2, 4))

        prompt_container, self.txt_custom_prompt = make_scrollable_text(
            self.prompt_editor, width=100, height=11, bg="#fffce0")
        # Always show SOMETHING — the saved override if present, otherwise the
        # system default (with icon instructions) — so the user can see exactly
        # what will be sent, and edit it directly instead of starting from blank.
        current_prompt = self.prompt_overrides.get("single", "").strip() or DEFAULT_PROMPT_SINGLE
        self.txt_custom_prompt.insert("1.0", current_prompt)
        prompt_container.pack(anchor="w", padx=10, pady=4)

        prompt_btns = ttk.Frame(self.prompt_editor)
        prompt_btns.pack(anchor="w", padx=10, pady=4)
        ttk.Button(prompt_btns, text="💾 Save custom prompt as default", command=self._save_custom_prompt)\
            .pack(side="left", padx=(0, 6))
        ttk.Button(prompt_btns, text="↺ Reset prompt to system default", command=self._reset_custom_prompt)\
            .pack(side="left", padx=6)
        ttk.Button(prompt_btns, text="📘 Show system default prompt (single + bilingual)", command=self._show_system_prompt)\
            .pack(side="left", padx=6)
        self.prompt_editor.grid_remove()  # remembers its grid options for _toggle_prompt_editor()

        # ── translation helper (Copilot bridge) ──
        ttk.Separator(f).grid(row=14, column=0, columnspan=2, sticky="ew", pady=8)
        ttk.Label(f, text="Translate the FULL email (greeting + note + event details) via Copilot copy/paste bridge:",
                  font=("Arial", 10, "bold")).grid(row=15, column=0, columnspan=2, sticky="w", **pad)
        ttk.Label(f, text="(Source is the language currently shown above — switch off Bilingual to pick a source)",
                  font=("Arial", 9, "italic")).grid(row=16, column=0, columnspan=2, sticky="w", padx=10)

        helper = ttk.Frame(f)
        helper.grid(row=17, column=0, columnspan=2, sticky="w", padx=10)

        ttk.Label(helper, text="Translate into:").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        self.combo_translate_target = ttk.Combobox(helper, width=26, state="readonly", values=TRANSLATE_TARGETS)
        self.combo_translate_target.current(1)  # default Japanese
        self.combo_translate_target.grid(row=0, column=1, sticky="w", padx=4, pady=4)
        ttk.Button(helper, text="📋 Copy full email + prompt to clipboard (for Copilot)",
                   command=self._copy_email_for_translation).grid(row=0, column=2, sticky="w", padx=8, pady=4)

        ttk.Label(helper, text="Paste the translated result from Copilot here — this will be used\n"
                              "as the complete, ready-to-send email (no section labels).\n"
                              "⚠️ Pasting a NEW result? Click '🗑 Clear' FIRST — pasting into a\n"
                              "non-empty box appends instead of replacing, combining old + new text.")\
            .grid(row=1, column=0, columnspan=3, sticky="w", padx=4, pady=(8, 2))
        paste_container, self.txt_translation_paste = make_scrollable_text(helper, width=90, height=6)
        paste_container.grid(row=2, column=0, columnspan=3, sticky="w", padx=4)
        self._make_wrapping_label(helper, text="⚠️ If words look glued together with no spaces (a known Copilot-copy quirk,\n"
                              "outside this tool's control), click 'Clean up' below, then click 'Save' again\n"
                              "to apply the cleaned-up version — Clean up alone does not change what is sent.",
                  font=("Arial", 8, "italic"), foreground="#8a4b00")\
            .grid(row=3, column=0, columnspan=3, sticky="w", padx=4, pady=(0, 2))
        ttk.Button(helper, text="💾 Save as translated version for this language", command=self._save_translated_email)\
            .grid(row=4, column=0, sticky="w", padx=4, pady=6)
        ttk.Button(helper, text="🗑 Clear saved translation for this language", command=self._clear_translated_email)\
            .grid(row=4, column=1, sticky="w", padx=4, pady=6)
        ttk.Button(helper, text="🧹 Clean up (fix lost spaces/line breaks)", command=self._cleanup_pasted_text)\
            .grid(row=4, column=2, sticky="w", padx=4, pady=6)

        self.lbl_translation_status = ttk.Label(
            helper, text="Full translated email ready: EN ❌  |  JA ❌  |  VI ❌  |  Bilingual ❌")
        self.lbl_translation_status.grid(row=5, column=0, columnspan=3, sticky="w", padx=4, pady=2)

        # ── send controls ──
        ttk.Separator(f).grid(row=18, column=0, columnspan=2, sticky="ew", pady=8)

        self._make_wrapping_label(f, text='Send to (optional override — e.g. a department group email).\n'
                          "Leave blank to send individually to everyone on Tab 2.\n"
                          "Tab 2's roster is always used for vote tracking either way.",
                  font=("Arial", 9, "italic")).grid(row=19, column=0, columnspan=2, sticky="w", **pad)
        self.var_send_to_override = tk.StringVar(value="")
        ttk.Entry(f, textvariable=self.var_send_to_override, width=50).grid(row=20, column=0, columnspan=2, sticky="w", padx=10)

        self.var_auto_send = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text="Send immediately without review (unchecked = open Outlook for you to click Send)",
                         variable=self.var_auto_send).grid(row=22, column=0, columnspan=2, sticky="w", **pad)

        self.var_send_btn_label = tk.StringVar(value="✉ Send Invite via Outlook (Voting Buttons)")
        ttk.Button(f, textvariable=self.var_send_btn_label, command=self._send_invite)\
            .grid(row=23, column=0, sticky="w", **pad)

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
        if self.prompt_editor.grid_info():
            self.prompt_editor.grid_remove()
            self.var_prompt_toggle.set("▸ Show / edit")
        else:
            self.prompt_editor.grid()
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

    def _single_lang_full_body(self, lang_code, event_name, event_date, location, deadline, budget):
        """Full assembled body (greeting + note + fixed) for ONE language — used both for
        sending and as a building block for the bilingual (JA+EN) combination."""
        override = self.full_translations.get(lang_code, "").strip()
        if override:
            return override
        greeting = build_greeting(lang_code)
        fixed = self.fixed_overrides.get(lang_code, "").strip() or \
            build_fixed_block(lang_code, event_name, event_date, location, deadline, budget)
        editable = build_editable_block(lang_code, self._source_note_text(), False)
        return f"{greeting}\n\n{editable}\n\n{fixed}".strip()

    # ── MỚI: các hàm build nội dung riêng cho mode "Send Gift Contribution
    # Notice" — song song với _lang_content()/_single_lang_full_body() ở
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

    def _gift_single_lang_full_body(self, lang_code, guest_of_honor, start_time, event_date, location,
                                     organizer, deadline, gift_budget):
        override = self.gift_full_translations.get(lang_code, "").strip()
        if override:
            return override
        greeting = build_greeting(lang_code)
        fixed = build_gift_fixed_block(lang_code, guest_of_honor, start_time, event_date, location,
                                        organizer, deadline, gift_budget)
        editable = build_editable_block(lang_code, self._source_note_text(), False)
        return f"{greeting}\n\n{editable}\n\n{fixed}".strip()

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

        # MỚI: nhánh RIÊNG hoàn toàn cho mode "Send Gift Contribution Notice"
        # — cùng cấu trúc UI (Subject/Greeting/Editable/Fixed/Bilingual) như
        # nhánh Invite bên dưới, nhưng dùng nội dung + subject + dict lưu bản
        # dịch Copilot RIÊNG cho Gift (xem _gift_lang_content()/
        # _gift_single_lang_full_body() ở trên) — return sớm, không chạy tiếp
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
                self.var_greeting_preview.set("Greeting (auto): shown inside the combined bilingual draft below (Japanese first, English second).")
                override = self.gift_full_translations.get("bilingual", "").strip()
                if override:
                    combined = override
                else:
                    ja_full = self._gift_single_lang_full_body(
                        "ja", guest_of_honor, start_time, event_date, location, organizer, gift_deadline, gift_budget)
                    en_full = self._gift_single_lang_full_body(
                        "en", guest_of_honor, start_time, event_date, location, organizer, gift_deadline, gift_budget)
                    combined = "[English below]\n\n" + ja_full + BILINGUAL_SEPARATOR + en_full
                fixed_display = ("→ Bilingual mode: the FIXED box is not used here. The complete bilingual "
                                  "draft (Japanese first, English second) is shown in the EDITABLE box below — "
                                  "you can hand-edit it there before sending.")
                editable_display = combined
                self.txt_fixed_preview.config(state="normal")
                self.txt_fixed_preview.delete("1.0", "end")
                self.txt_fixed_preview.insert("1.0", fixed_display)
                self.txt_fixed_preview.config(state="disabled")
            else:
                self.var_greeting_preview.set(f"Greeting (auto, appears first): {build_greeting(lang_code)}")
                fixed_text, editable_text, override_used = self._gift_lang_content(
                    lang_code, guest_of_honor, start_time, event_date, location, organizer, gift_deadline, gift_budget)
                fixed_display = fixed_text if fixed_text is not None else \
                    "→ Using a full Copilot-translated email (see EDITABLE box below — this FIXED box is unused for this language)."
                editable_display = editable_text
                self.txt_fixed_preview.config(state="normal")
                self.txt_fixed_preview.delete("1.0", "end")
                self.txt_fixed_preview.insert("1.0", fixed_display)
                if override_used:
                    self.txt_fixed_preview.config(state="disabled")

            self.txt_editable_preview.delete("1.0", "end")
            self.txt_editable_preview.insert("1.0", editable_display)
            self._refresh_translation_status()
            return

        subject = build_subject(lang_code, event_id, event_name, is_update=self._is_update_mode())
        self._compose_subject = subject
        self.var_subject_preview.set(f"Subject preview: {subject}")

        if lang_code == "bilingual":
            self.var_greeting_preview.set("Greeting (auto): shown inside the combined bilingual draft below (Japanese first, English second).")
            override = self.full_translations.get("bilingual", "").strip()
            if override:
                combined = override
            else:
                ja_full = self._single_lang_full_body("ja", event_name, event_date, location, deadline, budget)
                en_full = self._single_lang_full_body("en", event_name, event_date, location, deadline, budget)
                combined = "[English below]\n\n" + ja_full + BILINGUAL_SEPARATOR + en_full
            if self._is_update_mode():
                combined = build_update_notice("bilingual") + "\n\n" + combined
            fixed_display = ("→ Bilingual mode: the FIXED box is not used here. The complete bilingual "
                              "draft (Japanese first, English second) is shown in the EDITABLE box below — "
                              "you can hand-edit it there before sending.")
            editable_display = combined
            self.txt_fixed_preview.config(state="normal")
            self.txt_fixed_preview.delete("1.0", "end")
            self.txt_fixed_preview.insert("1.0", fixed_display)
            self.txt_fixed_preview.config(state="disabled")
        else:
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

    def _refresh_translation_status(self):
        translations = self._active_full_translations()
        def mark(code):
            return "✅" if translations.get(code, "").strip() else "❌"
        self.lbl_translation_status.config(
            text=f"Full translated email ready: EN {mark('en')}  |  JA {mark('ja')}  |  "
                 f"VI {mark('vi')}  |  Bilingual {mark('bilingual')}"
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
            "Used when 'Translate into' = Bilingual (Japanese + English). This one is "
            "fixed/built-in (no separate customization box for it yet):\n\n"
            + DEFAULT_PROMPT_BILINGUAL
        )

    def _copy_email_for_translation(self):
        lang_code = self._current_lang_code()
        if lang_code == "bilingual":
            messagebox.showinfo(
                "Switch language first",
                "Pick the SOURCE language you're drafting in first (English/Japanese/Vietnamese).\n\n"
                "You can still choose 'Bilingual (Japanese + English)' as the TRANSLATE-INTO target "
                "below — Copilot will produce both languages from that single source."
            )
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
        event_name = self.var_event_name.get()
        event_date = get_date_str(self.date_event)
        location = self.var_location.get()
        deadline = get_date_str(self.date_deadline)
        budget = self.var_budget.get()
        if self._is_gift_mode():
            fixed_text = build_gift_fixed_block(
                lang_code, self.var_guest_of_honor.get(), self.var_start_time.get(), event_date,
                location, self.var_organizer.get(), get_date_str(self.date_gift_deadline), self.var_gift_budget.get())
        else:
            fixed_text = self.fixed_overrides.get(lang_code, "").strip() or \
                build_fixed_block(lang_code, event_name, event_date, location, deadline, budget)
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
        target_code = self._target_lang_code()

        # Email order is: Greeting -> organizer's note -> event details/voting instructions.
        note_part = editable_text if editable_text else "(no note was written for this event)"
        combined = f"{greeting}\n\n{note_part}\n\n{fixed_text}"

        # Build prompt.
        # IMPORTANT: for the single-language box, read directly from what's shown/edited
        # in txt_custom_prompt — that box is ALWAYS pre-filled with either the saved
        # override or the system default (which already includes icon instructions), so
        # "what you see in the box is exactly what gets sent" — no more silently falling
        # back to an old icon-less hardcoded prompt just because nothing was Saved yet.
        if target_code == "bilingual":
            custom_prompt_bilingual = self.prompt_overrides.get("bilingual", "").strip()
            prompt_template = custom_prompt_bilingual or DEFAULT_PROMPT_BILINGUAL
            prompt = prompt_template + "\n\n" + combined
        else:
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

        # ✅ ALWAYS check for full translation override FIRST (from Copilot)
        # MỚI: dùng đúng dict theo mode hiện tại (Gift dùng self.gift_full_translations
        # riêng, không lẫn với self.full_translations của Invite/Update invite).
        full_override = self._active_full_translations().get(lang_code, "").strip()
        if full_override:
            # Use the COMPLETE translated email (greeting + note + details already in it)
            return full_override
        
        # If no override, build manually from components
        if lang_code == "bilingual":
            # BUG ĐÃ SỬA: UI ghi rõ ở ô FIXED (bilingual mode) "...shown in the
            # EDITABLE box below — you can hand-edit it there before sending",
            # nhưng trước đây hàm này KHÔNG hề đọc ô Editable — luôn build lại
            # từ đầu bằng self._source_note_text() (ghi chú gốc Tab 1), nên
            # MỌI chỉnh sửa tay trực tiếp trên Tab 3 đều bị bỏ qua lúc gửi
            # thật. Giờ ưu tiên đọc TRỰC TIẾP từ ô Editable đang hiển thị —
            # đúng bản nháp song ngữ bạn đã xem/sửa trước khi gửi.
            box_text = self.txt_editable_preview.get("1.0", "end").strip()
            if box_text:
                return box_text
            # Ô đang trống (vd chưa từng bấm sang Bilingual/chưa Refresh
            # preview lần nào) -> build mới như cũ, làm fallback an toàn.
            # MỚI: fallback cũng phải phân biệt Gift mode (dùng thông tin
            # Guest of Honor/Organizer/Gift budget) vs Invite thường.
            if self._is_gift_mode():
                ja_full = self._gift_single_lang_full_body(
                    "ja", self.var_guest_of_honor.get(), self.var_start_time.get(),
                    get_date_str(self.date_event), self.var_location.get(),
                    self.var_organizer.get(), get_date_str(self.date_gift_deadline), self.var_gift_budget.get())
                en_full = self._gift_single_lang_full_body(
                    "en", self.var_guest_of_honor.get(), self.var_start_time.get(),
                    get_date_str(self.date_event), self.var_location.get(),
                    self.var_organizer.get(), get_date_str(self.date_gift_deadline), self.var_gift_budget.get())
            else:
                ja_full = self._single_lang_full_body("ja", self.var_event_name.get(),
                                                       get_date_str(self.date_event),
                                                       self.var_location.get(),
                                                       get_date_str(self.date_deadline),
                                                       self.var_budget.get())
                en_full = self._single_lang_full_body("en", self.var_event_name.get(),
                                                       get_date_str(self.date_event),
                                                       self.var_location.get(),
                                                       get_date_str(self.date_deadline),
                                                       self.var_budget.get())
            return "[English below]\n\n" + ja_full + BILINGUAL_SEPARATOR + en_full
        else:
            # Build single language: greeting + editable + fixed
            # (đọc TRỰC TIẾP từ 2 ô đang hiển thị trên Tab 3 — đã đúng nội
            # dung Gift hay Invite tuỳ mode, vì _refresh_compose_preview() đã
            # điền đúng nội dung cho từng mode; không cần branch thêm ở đây)
            greeting = build_greeting(lang_code)
            fixed_text = self.txt_fixed_preview.get("1.0", "end").strip()
            editable_text = self.txt_editable_preview.get("1.0", "end").strip()
            return f"{greeting}\n\n{editable_text}\n\n{fixed_text}".strip()

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

        info_frame = ttk.LabelFrame(f, text="🔍 How 'Scan Inbox' works — read before using")
        info_frame.pack(fill="x", padx=10, pady=(8, 4))
        self._make_wrapping_label(
            info_frame, justify="left",
            text=(
                "• Scans EVERY folder of your mailbox (Inbox, its subfolders, and folders your rules "
                "move mail into) for emails whose Subject contains the EXACT current Event ID from Tab 1 — it doesn't matter whether "
                "the email is the original invite, an update invite, or a reminder, as long as the "
                "Subject contains the Event ID it counts.\n"
                "• For each sender, only the MOST RECENT vote is kept if they clicked more than once "
                "(votes are not added up).\n"
                "• Does NOT re-run automatically when you change the Event ID on Tab 1 — the table "
                "below only updates when you click '📨 Scan Inbox for Vote results'. If you switch to "
                "a different event and forget to click it again, the table will still show the results "
                "of the OLD Event ID (check the status line right below to be sure).\n"
                "• Results, manual corrections and the Yes/No/Maybe counts in History are saved "
                "automatically after every scan or edit — there is no separate Save button."
            ),
            font=("Arial", 8)).pack(anchor="w", padx=8, pady=(4, 2))
        self.lbl_scan_status = ttk.Label(info_frame, text="", font=("Arial", 9, "bold"))
        self.lbl_scan_status.pack(anchor="w", padx=8, pady=(2, 6))

        top = ttk.Frame(f)
        top.pack(fill="x", padx=10, pady=8)
        ttk.Button(top, text="📨 Scan Inbox for Vote results", command=self._collect_responses)\
            .pack(side="left", padx=4)

        ttk.Label(f, text="✅ Responded / not yet responded (based on Tab 2 + scanned votes):",
                  font=("Arial", 9, "bold")).pack(anchor="w", padx=10, pady=(4, 0))
        # MỚI: double-click ô "Vote" để sửa tay — dùng cho những người chỉ
        # đổi ý/báo miệng thay vì bấm lại nút Vote trong email (xem
        # _on_response_tree_double_click()/_commit_response_vote_edit()).
        # MỚI: thêm hẳn 1 cột checkbox "Manual edit" (✅/⬜) bên trái cột
        # "Name" — tick vào đó để mở dropdown Yes/No/Maybe sửa ngay (giống
        # double-click ô Vote); bỏ tick lại (với dòng ĐANG sửa tay) để trả
        # phiếu vote đó về cho lần Scan Inbox sau tự do cập nhật (xem
        # _on_response_manual_check_click()).
        self._make_wrapping_label(
            f, text="💡 Tick the \"Manual edit\" checkbox (or double-click the \"Vote\" cell) to correct "
                    "someone's vote by hand (Yes/No/Maybe) — useful when they just told you verbally or "
                    "changed their mind instead of clicking the button in the email. Manually-edited rows "
                    "are shown with a light blue background and stay checked. Un-ticking the checkbox for "
                    "an already manually-edited row lets the NEXT \"📨 Scan Inbox for Vote results\" freely "
                    "update it again from real emails; leaving it ticked KEEPS your manual value even after "
                    "re-scanning, unless a NEWER email vote is found for that same person.",
            font=("Arial", 8, "italic")).pack(anchor="w", padx=10)
        cols = ("manual_edit", "name", "email", "vote", "received")
        tree_container, self.tree_responses = make_scrollable_treeview(f, columns=cols, height=10)
        headers = ["✏️", "Name", "Email", "Vote", "Received At"]
        widths = [70, 220, 260, 120, 160]
        for c, label, w in zip(cols, headers, widths):
            self.tree_responses.heading(c, text=label)
            anchor = "center" if c == "manual_edit" else "w"
            self.tree_responses.column(c, width=w, anchor=anchor)
        self.tree_responses.bind("<Double-1>", self._on_response_tree_double_click)
        self.tree_responses.bind("<Button-1>", self._on_response_manual_check_click)
        tree_container.pack(fill="both", expand=True, padx=10, pady=6)

        self.lbl_summary = ttk.Label(f, text="No responses scanned yet.", font=("Arial", 10, "bold"))
        self.lbl_summary.pack(anchor="w", padx=10, pady=4)

        # ══════════════════════════════════════════════════════════════
        # KHUNG PHÍA DƯỚI — CHƯA PHẢN HỒI (pending), + soạn/gửi email nhắc nhở
        # ══════════════════════════════════════════════════════════════
        ttk.Separator(f).pack(fill="x", padx=10, pady=8)
        self.lbl_deadline_banner = ttk.Label(f, text="", font=("Arial", 10, "bold"))
        self.lbl_deadline_banner.pack(anchor="w", padx=10, pady=(0, 4))

        ttk.Label(f, text="🕓 NOT yet responded (no vote) — auto-updates every time you Scan Inbox:",
                  font=("Arial", 9, "bold"), foreground="#8a4b00").pack(anchor="w", padx=10, pady=(2, 0))
        pending_cols = ("name", "email")
        pending_container, self.tree_pending = make_scrollable_treeview(f, columns=pending_cols, height=6)
        for c, label, w in zip(pending_cols, ["Name", "Email"], [280, 320]):
            self.tree_pending.heading(c, text=label)
            self.tree_pending.column(c, width=w)
        pending_container.pack(fill="both", expand=False, padx=10, pady=6)

        reminder_frame = ttk.LabelFrame(f, text="📨 Compose & send a reminder email to people who haven't responded")
        reminder_frame.pack(fill="x", padx=10, pady=6)

        rbtn = ttk.Frame(reminder_frame)
        rbtn.pack(fill="x", padx=6, pady=4)
        ttk.Label(rbtn, text="Language:").pack(side="left", padx=(0, 4))
        self.combo_reminder_lang = ttk.Combobox(
            rbtn, width=26, state="readonly",
            values=[LANG_LABELS["en"], LANG_LABELS["ja"], LANG_LABELS["vi"], LANG_LABELS["bilingual"]],
        )
        self.combo_reminder_lang.current(0)  # default: English
        self.combo_reminder_lang.pack(side="left", padx=(0, 6))
        # Changing the language auto-regenerates the draft right away (same
        # behavior as combo_email_lang on Tab 3), so the user sees the new
        # language's content immediately instead of needing an extra click
        # on "Regenerate".
        self.combo_reminder_lang.bind("<<ComboboxSelected>>", lambda e: self._generate_reminder_draft())
        ttk.Button(rbtn, text="🔄 Regenerate reminder text (from Tab 1)",
                   command=self._generate_reminder_draft).pack(side="left", padx=(0, 6))
        self.var_reminder_attach = tk.BooleanVar(value=True)
        ttk.Checkbutton(rbtn, text="📎 Attach original invite email (looked up in Sent Items)",
                        variable=self.var_reminder_attach).pack(side="left", padx=6)
        # Every reminder (RSVP here, Gift on Tab 6) only opens Outlook for
        # review; there is deliberately no auto-send option.

        ttk.Label(reminder_frame, text="Reminder email content (review/edit before sending):",
                  font=("Arial", 9, "italic")).pack(anchor="w", padx=6, pady=(4, 0))
        reminder_text_container, self.txt_reminder_body = make_scrollable_text(
            reminder_frame, width=100, height=7)
        reminder_text_container.pack(fill="x", padx=6, pady=4)

        self.lbl_reminder_send = ttk.Button(
            reminder_frame, text="📨 Send reminder email to 0 people who haven't responded",
            command=self._send_reminder)
        self.lbl_reminder_send.pack(anchor="w", padx=6, pady=(2, 8))

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
            color = "#8a4b00"
        elif last_id is None:
            text = (f"🔍 Current Event ID: '{current_id}' — NOT scanned yet this session. "
                    f"Click '📨 Scan Inbox for Vote results' to fetch results.")
            color = "#8a4b00"
        elif last_id != current_id:
            text = (f"⚠️ The table below is showing results for the OLD Event ID ('{last_id}', scanned at "
                    f"{last_time_str}) — NOT the current Event ID ('{current_id}'). Click Scan Inbox "
                    f"again to refresh!")
            color = "#c42b1c"
        elif not getattr(self, "_scan_in_history", True):
            text = (f"✅ Showing results for Event ID '{current_id}' — scanned at {last_time_str}. "
                    f"⚠️ This event is not in History yet, so its Yes/No/Maybe counts are not recorded "
                    f"there: save it on Tab 1 ('💾 Update this event in History').")
            color = "#8a4b00"
        else:
            text = f"✅ Showing results for Event ID '{current_id}' — scanned at {last_time_str}."
            color = "#0e700e"
        self.lbl_scan_status.config(text=text, foreground=color)

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
        '🔎 Expand group emails in list' ở Tab 2.

        Kết quả mỗi group được CACHE lại (self._group_expansion_cache) để
        không phải hỏi lại Exchange GAL mỗi lần bấm Scan Inbox trong cùng
        phiên làm việc — chỉ query lần đầu tiên gặp mỗi group email."""
        # The merge and de-duplication rules live in rsvp/domain/roster.py,
        # which is testable without Outlook. Expansion itself needs the
        # address book, so it is passed in rather than imported there.
        return merge_expanded_roster(
            self.recipients if recipients is None else recipients,
            self.outlook.expand_group_members,
            cache=self._group_expansion_cache,
        )

    def _refresh_response_tree(self, skipped=0, roster=None):
        roster = roster if roster is not None else self.recipients
        # MỚI: nhớ lại roster vừa dùng để vẽ bảng — dùng khi sửa tay 1 ô Vote
        # (_commit_response_vote_edit()) cần vẽ lại toàn bộ bảng mà KHÔNG phải
        # tính lại _build_effective_roster() (có thể gọi COM để mở rộng group
        # email, không nên chạy lại chỉ vì sửa 1 ô).
        self._last_responses_roster = roster
        self.tree_responses.delete(*self.tree_responses.get_children())
        self.tree_responses.tag_configure("extra", background="#fff3cd")
        # MỚI: tô nền xanh nhạt cho các dòng có phiếu vote đã được SỬA TAY
        # (double-click ô Vote — xem _commit_response_vote_edit()), để phân
        # biệt trực quan với phiếu quét được thật sự từ email.
        self.tree_responses.tag_configure("manual", background="#dceeff")
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
            banner = f"✅ Response deadline: {deadline_str}  —  everyone in Tab 2 has responded."
        elif deadline_date is None:
            banner = f"🕓 {len(pending)} people still haven't responded. (Couldn't read deadline date: '{deadline_str}')"
        elif overdue:
            banner = (f"🔴 Response deadline REACHED/PASSED ({deadline_str}) — {len(pending)} people still haven't "
                       f"replied. Review the reminder content below, then click send.")
        else:
            days_left = (deadline_date - datetime.now().date()).days
            banner = (f"🕓 Response deadline: {deadline_str} ({days_left} days left) — "
                      f"{len(pending)} people still haven't replied.")
        self.lbl_deadline_banner.config(text=banner)

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
                "Click '🔄 Regenerate reminder text' or type the content by hand before sending.")
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
        pad = {"padx": 10, "pady": 6}

        ttk.Label(f, text="Gift Contribution tracking", font=("Arial", 11, "bold"))\
            .grid(row=0, column=0, columnspan=3, sticky="w", **pad)
        self._make_wrapping_label(f, text="Recipient list pulled from Tab 2 (Recipients). Check off everyone who has\n"
                          "ALREADY contributed money for the gift — the checked state is saved to the\n"
                          "database automatically as soon as you click, and reloaded automatically if\n"
                          "there's saved data (nothing is lost when you close and reopen the app).",
                  font=("Arial", 9, "italic")).grid(row=1, column=0, columnspan=3, sticky="w", padx=10)

        btns = ttk.Frame(f)
        btns.grid(row=2, column=0, columnspan=3, sticky="w", **pad)
        # The list follows Tab 2 every time this tab opens. Importing a file
        # (e.g. one a colleague edited in Excel) instead OVERWRITES the
        # checked state and amount of EVERYONE found in it.
        ttk.Button(btns, text="📂 Load contribution list from file", command=self._load_gift_list_from_file)\
            .pack(side="left")
        # MỚI: đã BỎ 2 nút "☑ Check all"/"☐ Uncheck all" — thay bằng 1 dấu
        # ✅/⬜ NGAY TRÊN ĐẦU CỘT "Send email" và cột "Contributed" (bấm
        # trực tiếp vào tiêu đề cột, xem tree.heading(..., command=...) bên
        # dưới và _toggle_all_gift_column()) — mỗi cột tự chọn/bỏ chọn tất
        # cả ĐỘC LẬP với nhau, và vẫn chỉ áp dụng cho các dòng ĐANG HIỂN THỊ
        # (tôn trọng ô tìm kiếm, giống hành vi 2 nút cũ).

        # Quick Name/Email search box — filters the table below directly as you
        # type, case-insensitive. The checked state stays stored in
        # self._gift_roster (the source of truth), so filtering/searching
        # never loses anyone's checked state, even while they're hidden.
        search_bar = ttk.Frame(f)
        search_bar.grid(row=3, column=0, columnspan=3, sticky="w", **pad)
        ttk.Label(search_bar, text="🔎 Search (name or email):").pack(side="left")
        self.var_gift_search = tk.StringVar(value="")
        ttk.Entry(search_bar, textvariable=self.var_gift_search, width=40).pack(side="left", padx=6)
        self.var_gift_search.trace_add("write", lambda *a: self._apply_gift_filter())
        ttk.Button(search_bar, text="✕ Clear", command=lambda: self.var_gift_search.set(""))\
            .pack(side="left", padx=4)

        # Column order: send_email, check, No. (sequence number), Name, Email, Amount.
        # MỚI: "send_email" — cột checkbox THÊM MỚI, đặt bên trái "check"
        # (Contributed) — người dùng tick chọn AI sẽ nhận email báo cáo
        # quyên góp (xem khối "📧 Send contribution report" ở cuối tab, và
        # _send_gift_report_email()) — HOÀN TOÀN ĐỘC LẬP với "Contributed"
        # (1 người có thể được chọn nhận báo cáo dù họ CHƯA đóng góp, hoặc
        # ngược lại).
        cols = ("send_email", "check", "no", "name", "email", "amount")
        tree_container, self.tree_gift = make_scrollable_treeview(f, columns=cols, height=16)
        # Tiêu đề 2 cột checkbox ("send_email"/"check") tự hiển thị ✅/⬜
        # phản ánh "tất cả dòng ĐANG HIỂN THỊ có đang được tick hết không"
        # — bấm vào TIÊU ĐỀ (không phải 1 ô cụ thể) để chọn/bỏ chọn tất cả
        # cùng lúc cho đúng cột đó (xem _toggle_all_gift_column()).
        self.tree_gift.heading("send_email", text="⬜ Send email",
                                command=lambda: self._toggle_all_gift_column("send_email"))
        self.tree_gift.heading("check", text="⬜ Contributed",
                                command=lambda: self._toggle_all_gift_column("check"))
        self.tree_gift.heading("no", text="No.")
        self.tree_gift.heading("name", text="Name")
        self.tree_gift.heading("email", text="Email")
        self.tree_gift.heading("amount", text="Amount")
        self.tree_gift.column("send_email", width=100, anchor="center")
        self.tree_gift.column("check", width=100, anchor="center")
        self.tree_gift.column("no", width=40, anchor="center")
        self.tree_gift.column("name", width=240)
        self.tree_gift.column("email", width=280)
        self.tree_gift.column("amount", width=100, anchor="e")
        tree_container.grid(row=4, column=0, columnspan=3, sticky="w", padx=10)
        # Clicking the "send_email" or "check" (Contributed) column toggles
        # that row's checked state (tkinter's Treeview has no real checkbox
        # widget, so ✅/⬜ characters simulate one — a common, simple
        # approach that avoids pulling in an extra third-party library).
        self.tree_gift.bind("<Button-1>", self._on_gift_tree_click)
        # Double-click Amount to type what someone actually gave.
        self.tree_gift.bind("<Double-1>", self._on_gift_tree_double_click)

        ttk.Label(f, text="Total contributors:").grid(row=5, column=0, sticky="w", **pad)
        self.var_gift_contributed_count = tk.StringVar(value="0 / 0")
        ttk.Label(f, textvariable=self.var_gift_contributed_count, font=("Arial", 10, "bold"))\
            .grid(row=5, column=1, sticky="w", **pad)

        # Total amount collected so far — sum of "Amount" for everyone
        # currently checked as "Contributed" (see _update_gift_contributed_count()).
        ttk.Label(f, text="Total amount collected:").grid(row=6, column=0, sticky="w", **pad)
        self.var_gift_total_amount = tk.StringVar(value="0")
        ttk.Label(f, textvariable=self.var_gift_total_amount, font=("Arial", 10, "bold"))\
            .grid(row=6, column=1, sticky="w", **pad)

        ttk.Button(f, text="📊 Export to Excel", command=self._export_gift_contribution_list)\
            .grid(row=7, column=0, columnspan=2, sticky="w", **pad)
        self._make_wrapping_label(f, text="→ Ticking \"Contributed\" fills Amount from Tab 1's expected gift budget; "
                          "double-click an Amount to type what that person actually gave. Every change is "
                          "saved to the database automatically — this button only creates an Excel file "
                          "when you need one to share or archive.",
                  font=("Arial", 8, "italic")).grid(row=8, column=0, columnspan=3, sticky="w", padx=14)

        # ── Compose & send a reminder email to people who haven't contributed yet ──
        # Always ONLY opens Outlook for review (mail.Display()) — there is no
        # "Send immediately without review" option here (unlike Tab 4's RSVP
        # reminder), per the requirement that every reminder email (both RSVP
        # and Gift) must be checked by the user and sent by hand from
        # Outlook, never sent automatically from within the app. See
        # _send_gift_reminder().
        gift_reminder_frame = ttk.LabelFrame(
            f, text="📨 Compose & send a reminder email to people who haven't contributed yet")
        gift_reminder_frame.grid(row=9, column=0, columnspan=3, sticky="we", padx=10, pady=6)

        grbtn = ttk.Frame(gift_reminder_frame)
        grbtn.pack(fill="x", padx=6, pady=4)
        ttk.Label(grbtn, text="Language:").pack(side="left", padx=(0, 4))
        self.combo_gift_reminder_lang = ttk.Combobox(
            grbtn, width=26, state="readonly",
            values=[LANG_LABELS["en"], LANG_LABELS["ja"], LANG_LABELS["vi"], LANG_LABELS["bilingual"]],
        )
        self.combo_gift_reminder_lang.current(0)  # default: English
        self.combo_gift_reminder_lang.pack(side="left", padx=(0, 6))
        self.combo_gift_reminder_lang.bind(
            "<<ComboboxSelected>>", lambda e: self._generate_gift_reminder_draft())
        ttk.Button(grbtn, text="🔄 Regenerate reminder text (from Tab 1)",
                   command=self._generate_gift_reminder_draft).pack(side="left", padx=(0, 6))
        self.var_gift_reminder_attach = tk.BooleanVar(value=True)
        ttk.Checkbutton(grbtn, text="📎 Attach original Gift email (looked up from Sent Items)",
                        variable=self.var_gift_reminder_attach).pack(side="left", padx=6)

        ttk.Label(gift_reminder_frame, text="Reminder email content (review/edit before sending):",
                  font=("Arial", 9, "italic")).pack(anchor="w", padx=6, pady=(4, 0))
        gift_reminder_text_container, self.txt_gift_reminder_body = make_scrollable_text(
            gift_reminder_frame, width=100, height=7)
        gift_reminder_text_container.pack(fill="x", padx=6, pady=4)

        self.lbl_gift_reminder_send = ttk.Button(
            gift_reminder_frame, text="📨 Open reminder email for 0 people who haven't contributed",
            command=self._send_gift_reminder)
        self.lbl_gift_reminder_send.pack(anchor="w", padx=6, pady=(2, 8))
        ttk.Label(gift_reminder_frame,
                  text="→ Always opens the email compose window in Outlook so you can review it and "
                       "click Send YOURSELF — there is no auto-send option here.",
                  font=("Arial", 8, "italic")).pack(anchor="w", padx=6, pady=(0, 6))

        # ── MỚI: gửi email báo cáo số tiền đã thu được, tới NHỮNG NGƯỜI ĐÃ
        # TICK CHỌN ở cột "Send email" (hoàn toàn tách biệt với người CHƯA
        # đóng góp ở khối nhắc nhở phía trên) — nội dung email liệt kê danh
        # sách những ai ĐÃ ĐÓNG GÓP (cột "Contributed" = Yes), đánh số lại
        # từ 1, KHÔNG đưa 2 cột checkbox ("Send email"/"Contributed") vào
        # danh sách đó — xem reports.gift_report_workbook()/build_gift_report_body().
        # Nội dung mail CHỈ là 1 thông báo TỔNG QUAN (đã thu bao nhiêu
        # người/bao nhiêu tiền), KHÔNG liệt kê từng người trong nội dung —
        # danh sách chi tiết nằm trong file Excel TỰ ĐỘNG ĐÍNH KÈM.
        report_frame = ttk.LabelFrame(f, text="📧 Send contribution report to selected people")
        report_frame.grid(row=10, column=0, columnspan=3, sticky="we", padx=10, pady=6)
        self._make_wrapping_label(
            report_frame,
            text="Sent to everyone ticked in the \"Send email\" column above. The message body is a "
                 "short summary (how many people contributed, total amount) — the detailed list "
                 "(everyone who has ALREADY contributed, renumbered, without the \"Send email\"/"
                 "\"Contributed\" checkbox columns) is automatically attached as an Excel file.",
            font=("Arial", 8, "italic")).pack(anchor="w", padx=6, pady=(4, 0))

        rbtn = ttk.Frame(report_frame)
        rbtn.pack(fill="x", padx=6, pady=4)
        ttk.Label(rbtn, text="Language:").pack(side="left", padx=(0, 4))
        self.combo_gift_report_lang = ttk.Combobox(
            rbtn, width=26, state="readonly",
            values=[LANG_LABELS["en"], LANG_LABELS["ja"], LANG_LABELS["vi"], LANG_LABELS["bilingual"]],
        )
        self.combo_gift_report_lang.current(0)  # default: English
        self.combo_gift_report_lang.pack(side="left", padx=(0, 6))
        self.combo_gift_report_lang.bind("<<ComboboxSelected>>", lambda e: self._generate_gift_report_draft())
        ttk.Button(rbtn, text="🔄 Regenerate report text", command=self._generate_gift_report_draft)\
            .pack(side="left", padx=(0, 6))

        ttk.Label(report_frame, text="Report email content (review/edit before sending):",
                  font=("Arial", 9, "italic")).pack(anchor="w", padx=6, pady=(4, 0))
        report_text_container, self.txt_gift_report_body = make_scrollable_text(
            report_frame, width=100, height=10)
        report_text_container.pack(fill="x", padx=6, pady=4)

        self.lbl_gift_report_send = ttk.Button(
            report_frame, text="📧 Send report to 0 selected people", command=self._send_gift_report_email)
        self.lbl_gift_report_send.pack(anchor="w", padx=6, pady=(2, 8))
        ttk.Label(report_frame,
                  text="→ Always opens the email compose window in Outlook so you can review it and "
                       "click Send YOURSELF — there is no auto-send option here.",
                  font=("Arial", 8, "italic")).pack(anchor="w", padx=6, pady=(0, 6))

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
                "Click '🔄 Regenerate reminder text' or type the content by hand before sending.")
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

    def _gift_report_body_args(self):
        """Lấy đúng (guest_of_honor, event_name, contributor_count,
        total_amount) hiện tại từ Tab 1 + bảng Gift Contribution (Tab 6) —
        dùng để build nội dung email báo cáo mặc định
        (build_gift_report_body()) — nội dung chỉ là 1 THÔNG BÁO TỔNG QUAN,
        danh sách chi tiết nằm trong file Excel đính kèm (xem
        reports.gift_report_workbook())."""
        contributor_count = sum(
            1 for info in getattr(self, "_gift_roster", {}).values() if info.get("checked"))
        return (
            self.var_guest_of_honor.get(),
            self.var_event_name.get(),
            str(contributor_count),
            getattr(self, "var_gift_total_amount", tk.StringVar(value="0")).get(),
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

        body = self.txt_gift_report_body.get("1.0", "end").strip()
        if not body:
            messagebox.showwarning(
                "Empty content",
                "Click '🔄 Regenerate report text' or type the content by hand before sending.")
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

        def worker():
            try:
                mail, attached = self.outlook.send_gift_report_email(
                    recipients, subject, body, excel_path=excel_path)
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
        if self._gift_event != event_id:
            # The roster in memory belongs to another event: start from what
            # is saved for this one, never from the other event's ticks.
            self._gift_roster = {}  # email -> {"name":..., "checked": bool, "amount": float, "send_email": bool}
            self._gift_event = event_id or None

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
            if email in self._gift_roster:
                checked = self._gift_roster[email]["checked"]
                amount = self._gift_roster[email].get("amount", 0.0)
                send_email = self._gift_roster[email].get("send_email", False)
            else:
                prev = saved_state.get(email, {})
                checked = prev.get("checked", False)
                amount = prev.get("amount", 0.0)
                send_email = prev.get("send_email", False)
            new_roster[email] = {"name": name, "checked": checked, "amount": amount, "send_email": send_email}
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
            if new_checked:
                self._gift_roster[email]["amount"] = parse_amount_from_text(self.var_gift_budget.get())
            else:
                self._gift_roster[email]["amount"] = 0.0
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
        """A typed amount replaces the expected budget for that person. A
        positive amount also marks them as Contributed, since the totals and
        the report count only contributors."""
        info = self._gift_roster.get(row_id)
        if info is None:
            return
        amount = parse_amount_from_text(new_value)
        info["amount"] = amount
        if amount > 0:
            info["checked"] = True
        self._apply_gift_filter()
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
                self._gift_roster[iid]["checked"] = new_state
                self._gift_roster[iid]["amount"] = (
                    parse_amount_from_text(self.var_gift_budget.get()) if new_state else 0.0
                )
            else:
                self._gift_roster[iid]["send_email"] = new_state
        self._apply_gift_filter()
        self._save_gift_roster_to_db(silent=True)

    def _update_gift_contributed_count(self):
        total = len(self._gift_roster) if hasattr(self, "_gift_roster") else 0
        contributed = sum(1 for info in self._gift_roster.values() if info["checked"]) if hasattr(self, "_gift_roster") else 0
        total_amount = sum(info.get("amount", 0.0) for info in self._gift_roster.values()) if hasattr(self, "_gift_roster") else 0.0
        shown = len(self.tree_gift.get_children())
        if shown != total:
            self.var_gift_contributed_count.set(f"{contributed} / {total}  (showing {shown}/{total} due to search)")
        else:
            self.var_gift_contributed_count.set(f"{contributed} / {total}")
        if hasattr(self, "var_gift_total_amount"):
            self.var_gift_total_amount.set(f"{total_amount:,.0f}")
        self._update_gift_reminder_button_label()
        self._update_gift_report_button_label()

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
        try:
            db.save_gift_roster(event_id, self._gift_roster, self.history_path.get())
        except Exception:
            if not silent:
                raise
            pass  # best-effort silent auto-save

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
        pad = {"padx": 10, "pady": 6}

        # MỚI: giờ lấy CẢ "Yes" LẪN "Maybe" (trước đây chỉ lấy "Yes") — vì
        # người trả lời "Maybe" (chưa chắc chắn) vẫn nên được mời Calendar
        # Invite chính thức, để họ có lịch trên Outlook phòng khi sau đó họ
        # quyết định tham dự — không nên loại họ ra chỉ vì chưa chốt.
        ttk.Label(f, text='People who voted "Yes" or "Maybe" (auto-pulled from Tab 4):',
                  font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", **pad)
        self.list_yes = tk.Listbox(f, width=60, height=10)
        yes_vscroll = ttk.Scrollbar(f, orient="vertical", command=self.list_yes.yview)
        self.list_yes.configure(yscrollcommand=yes_vscroll.set)
        self.list_yes.grid(row=1, column=0, sticky="w", padx=10)
        yes_vscroll.grid(row=1, column=1, sticky="nsw")
        # Kept in step with Tab 4 by _refresh_response_tree(); no refresh needed.

        # MỚI: đã BỎ dòng "Event date / time (from Tab 1)" ở đây theo yêu
        # cầu — thông tin đó vẫn được dùng NGẦM khi gửi Calendar Invite
        # (xem _send_calendar(), luôn lấy trực tiếp từ Tab 1), chỉ không
        # còn hiển thị lặp lại thành 1 dòng riêng ở Tab 5 nữa. Muốn xem/đổi
        # giờ, quay lại Tab 1.

        # ── customizable appointment body ──
        appt_header = ttk.Frame(f)
        appt_header.grid(row=6, column=0, columnspan=2, sticky="w", **pad)
        ttk.Label(appt_header, text="Appointment body (customize if needed):", font=("Arial", 9, "bold"))\
            .pack(side="left")
        ttk.Label(appt_header, text="   Language:").pack(side="left")
        self.combo_calendar_lang = ttk.Combobox(
            appt_header, width=26, state="readonly",
            values=[LANG_LABELS["en"], LANG_LABELS["ja"], LANG_LABELS["vi"], LANG_LABELS["bilingual"]],
        )
        self.combo_calendar_lang.current(0)  # default: English
        self.combo_calendar_lang.pack(side="left", padx=(4, 0))
        # Đổi ngôn ngữ -> tự thay lại nội dung mặc định ngay (giống combo_reminder_lang
        # ở Tab 4) — nếu bạn đã tự gõ tay nội dung riêng, đổi ngôn ngữ sẽ THAY THẾ
        # bằng bản mặc định của ngôn ngữ mới (không cộng dồn/giữ lại bản cũ).
        self.combo_calendar_lang.bind("<<ComboboxSelected>>", lambda e: self._apply_calendar_body_lang())

        appt_container, self.txt_appt_body = make_scrollable_text(f, width=100, height=4)
        default_body = build_calendar_body("en", *self._calendar_body_args())
        self.txt_appt_body.insert("1.0", default_body)
        self.var_appt_body_default = default_body  # Keep track of default for reset
        appt_container.grid(row=7, column=0, columnspan=2, sticky="w", padx=10, pady=4)

        ttk.Button(f, text="📅 Send Calendar Invite to the list above", command=self._send_calendar)\
            .grid(row=8, column=0, sticky="w", **pad)

        ttk.Separator(f).grid(row=9, column=0, columnspan=3, sticky="ew", padx=10, pady=10)

        # ── Attendance & Payment tracking (post-event) ──
        # Lists everyone who voted "Yes" or "Maybe" (same population as the
        # Calendar Invite list above), so you can mark who ACTUALLY showed
        # up, who's exempt from paying, and how much they paid, after the
        # event happens. Double-click "Name"/"Vote"/"Amount" to type a new
        # value; "Actual Attend" opens a Yes/No dropdown instead of free
        # text; single-click the "Free" column to toggle it. Every edit is
        # auto-saved straight to Attendance_Payment_{EventID}.xlsx in the
        # Input folder — no manual export step needed (see
        # _save_attendance_sheet_to_file()).
        ttk.Label(f, text="Attendance & Payment tracking (fill in after the event):",
                  font=("Arial", 10, "bold")).grid(row=10, column=0, columnspan=2, sticky="w", **pad)
        self._make_wrapping_label(
            f, text="Double-click \"Name\"/\"Vote\"/\"Amount\" to edit; \"Actual Attend\" opens a Yes/No "
                    "dropdown; click a row's \"Free\" cell to toggle exemption. Setting \"Actual Attend\" to "
                    "\"Yes\" auto-fills Amount from Tab 1's \"Expected event budget\" (0 if marked Free) — "
                    "you can still overwrite Amount by hand afterward. Select rows and use Ctrl+C / Ctrl+V "
                    "to copy/paste to or from Excel, or Delete to clear a row's tracking fields. Every "
                    "change is saved to the database automatically; the list follows Tab 4's "
                    "Yes/Maybe votes each time this tab opens.",
            font=("Arial", 8, "italic")).grid(row=11, column=0, columnspan=3, sticky="w", padx=10)

        attend_btns = ttk.Frame(f)
        attend_btns.grid(row=12, column=0, columnspan=3, sticky="w", **pad)
        # Every edit is saved to the database already; this only writes a
        # file to share or archive.
        ttk.Button(attend_btns, text="📊 Export to Excel", command=self._export_attendance_to_excel)\
            .pack(side="left")

        attend_cols = ("no", "name", "vote", "actual_attend", "free", "amount")
        attend_container, self.tree_attendance = make_scrollable_treeview(f, columns=attend_cols, height=14)
        attend_headers = ["No.", "Name", "Vote", "Actual Attend", "Free", "Amount"]
        attend_widths = [40, 240, 70, 110, 60, 100]
        for c, h, w in zip(attend_cols, attend_headers, attend_widths):
            self.tree_attendance.heading(c, text=h)
            anchor = "center" if c in ("no", "vote", "actual_attend", "free") else ("e" if c == "amount" else "w")
            self.tree_attendance.column(c, width=w, anchor=anchor)
        attend_container.grid(row=13, column=0, columnspan=3, sticky="w", padx=10, pady=6)
        # "Actual Attend" opens a readonly Yes/No dropdown (see
        # _begin_cell_edit_combobox()); Name/Vote/Amount get a normal
        # type-in-place Entry; "No." and "Free" are not double-click
        # editable ("Free" toggles on single-click instead — see
        # _on_attendance_free_click(), bound separately below so the two
        # bindings don't conflict).
        self.tree_attendance.bind("<Double-1>", self._on_attendance_tree_double_click)
        self.tree_attendance.bind("<Button-1>", self._on_attendance_free_click)
        self._enable_treeview_copy_paste(
            self.tree_attendance, on_commit=self._commit_attendance_edit,
            editable_cols={"name", "vote", "actual_attend", "free", "amount"})
        # Select rows and press Delete/Backspace to clear that row's
        # tracking fields (Actual Attend/Free/Amount) back to blank — Name
        # and Vote are left untouched since they're identity data pulled
        # from Tab 4, not something you'd normally want to blank out by
        # accident. See _on_attendance_delete_key().
        self.tree_attendance.bind("<Delete>", self._on_attendance_delete_key)
        self.tree_attendance.bind("<BackSpace>", self._on_attendance_delete_key)

        totals_frame = ttk.Frame(f)
        totals_frame.grid(row=14, column=0, columnspan=3, sticky="w", **pad)
        ttk.Label(totals_frame, text="Total actual attend:").pack(side="left")
        self.var_total_actual_attend = tk.StringVar(value="0")
        ttk.Label(totals_frame, textvariable=self.var_total_actual_attend, font=("Arial", 10, "bold"))\
            .pack(side="left", padx=(4, 24))
        ttk.Label(totals_frame, text="Total collected amount:").pack(side="left")
        self.var_total_collected_amount = tk.StringVar(value="0")
        ttk.Label(totals_frame, textvariable=self.var_total_collected_amount, font=("Arial", 10, "bold"))\
            .pack(side="left", padx=(4, 0))

        # MỚI: "Amount paid" — số tiền THỰC TẾ đã trả (vd đặt cọc/thanh
        # toán trước cho nhà hàng, chuyển khoản cho Guest of Honor...),
        # nhập tay vì con số này KHÔNG thể tự suy ra được từ bảng
        # Attendance. "Remaining amount" tự tính = Total collected amount −
        # Amount paid, cập nhật ngay khi gõ (self.var_amount_paid.trace_add
        # bên dưới) hoặc khi bảng Attendance đổi (_update_attendance_totals
        # gọi lại cùng công thức). Giá trị Amount paid được lưu theo từng
        # Event ID trong database (cột "AmountPaid" ở bảng events, xem
        # db.py) — nạp lại tự động bởi "Load setup from selected event".
        paid_frame = ttk.Frame(f)
        paid_frame.grid(row=15, column=0, columnspan=3, sticky="w", **pad)
        ttk.Label(paid_frame, text="Amount paid:").pack(side="left")
        self.var_amount_paid = tk.StringVar(value="0")
        ttk.Entry(paid_frame, textvariable=self.var_amount_paid, width=14)\
            .pack(side="left", padx=(4, 24))
        ttk.Label(paid_frame, text="Remaining amount:").pack(side="left")
        self.var_remaining_amount = tk.StringVar(value="0")
        ttk.Label(paid_frame, textvariable=self.var_remaining_amount, font=("Arial", 10, "bold"))\
            .pack(side="left", padx=(4, 0))
        # Ghi chú nhỏ để không ai nhầm đây là số tự tính từ bảng.
        ttk.Label(paid_frame, text="  (type the amount actually paid out; "
                                    "remaining = Total collected − Amount paid)",
                  font=("Arial", 8, "italic")).pack(side="left")
        self.var_amount_paid.trace_add("write", lambda *a: self._on_amount_paid_changed())

        # ── Thank You email (post-event, MỚI) ──
        # Soạn sẵn nội dung email cảm ơn mọi người đã tham gia, cùng cách
        # làm với Appointment body ở trên: đa ngôn ngữ (English/Japanese/
        # Vietnamese/Bilingual), có thể sửa tay trước khi gửi, tự refresh
        # theo Tab 1 khi CHƯA hand-edit (xem _refresh_thankyou_body_display()
        # / _apply_thankyou_body_lang(), cùng cơ chế với
        # _refresh_calendar_datetime_display()). Người nhận = những ai
        # "Actual Attend" = Yes trong bảng ở trên (xem _send_thank_you_email()).
        ttk.Separator(f).grid(row=16, column=0, columnspan=3, sticky="ew", padx=10, pady=10)

        ttk.Label(f, text="Thank You email (send after the event):",
                  font=("Arial", 10, "bold")).grid(row=17, column=0, columnspan=2, sticky="w", **pad)
        self._make_wrapping_label(
            f, text="Sent to everyone marked \"Actual Attend\" = Yes above. Automatically attaches "
                    "the Attendance & Payment Excel report (with everyone's attendance/cost info), "
                    "and the Calendar Invite for this event if it can be found in your Calendar "
                    "(best-effort — see the confirmation message after sending).",
            font=("Arial", 8, "italic")).grid(row=18, column=0, columnspan=3, sticky="w", padx=10)

        thankyou_header = ttk.Frame(f)
        thankyou_header.grid(row=19, column=0, columnspan=2, sticky="w", **pad)
        ttk.Label(thankyou_header, text="   Language:").pack(side="left")
        self.combo_thankyou_lang = ttk.Combobox(
            thankyou_header, width=26, state="readonly",
            values=[LANG_LABELS["en"], LANG_LABELS["ja"], LANG_LABELS["vi"], LANG_LABELS["bilingual"]],
        )
        self.combo_thankyou_lang.current(0)  # default: English
        self.combo_thankyou_lang.pack(side="left", padx=(4, 0))
        self.combo_thankyou_lang.bind("<<ComboboxSelected>>", lambda e: self._apply_thankyou_body_lang())

        thankyou_container, self.txt_thankyou_body = make_scrollable_text(f, width=100, height=6)
        default_thankyou_body = build_thankyou_body("en", *self._thankyou_body_args())
        self.txt_thankyou_body.insert("1.0", default_thankyou_body)
        self.var_thankyou_body_default = default_thankyou_body  # Keep track of default for reset
        thankyou_container.grid(row=20, column=0, columnspan=2, sticky="w", padx=10, pady=4)

        ttk.Button(f, text="📧 Send Thank You email to confirmed attendees",
                   command=self._send_thank_you_email).grid(row=21, column=0, sticky="w", **pad)

    def _refresh_attendance_list(self):
        """Rebuilds self._attendance_roster (the source of truth for the
        Attendance & Payment table) from the Yes/Maybe voters scanned on
        Tab 4 (self.tree_responses) — same population as the Calendar
        Invite list above. Existing "Actual Attend"/"Free"/"Amount" edits
        for people already in the roster are KEPT (so re-scanning Tab 4 or
        switching back to this tab doesn't wipe out attendance already
        marked); brand-new voters default to Actual Attend = "Yes" (with
        Amount auto-filled to match), per the requirement that Yes is the
        default rather than blank.

        Runs when Tab 5 opens. The table belongs to one event
        (self._attendance_event): when Tab 1 now shows another, the roster is
        first replaced by what is saved for that event, so marks never move
        between events; votes are merged only from a Tab 4 table that was
        scanned for that same event."""
        current_id = self.var_event_id.get().strip()
        if self._attendance_event != current_id:
            try:
                saved = db.load_attendance_roster(current_id, self.history_path.get()) if current_id else {}
            except Exception:
                saved = {}
            self._attendance_roster = saved
            self._attendance_event = current_id or None
        if not current_id or current_id != self._last_scanned_event_id:
            self._render_attendance_tree()
            return
        new_roster = {}
        newly_added = []
        for iid in self.tree_responses.get_children():
            _manual, name, email, vote, _received = self.tree_responses.item(iid, "values")
            if vote not in ("Yes", "Maybe"):
                continue
            is_new = email not in self._attendance_roster
            prev = self._attendance_roster.get(email, {})
            new_roster[email] = {
                "name": name,
                "vote": vote,
                "actual_attend": prev.get("actual_attend", "Yes"),
                "free": prev.get("free", False),
                "amount": prev.get("amount", 0.0),
            }
            if is_new:
                newly_added.append(email)
        self._attendance_roster = new_roster
        # Only auto-sync Amount for brand-new people (matching their
        # default "Yes") — anyone already in the roster keeps whatever
        # Amount they already had, even if it's 0 by manual choice.
        for email in newly_added:
            self._sync_attendance_amount(email)
        self._render_attendance_tree()

    def _render_attendance_tree(self):
        self.tree_attendance.delete(*self.tree_attendance.get_children())
        for i, (email, info) in enumerate(self._attendance_roster.items(), start=1):
            amount = info.get("amount", 0.0)
            amount_display = f"{amount:,.0f}" if amount else ""
            free_display = "✅" if info.get("free") else "⬜"
            self.tree_attendance.insert(
                "", "end", iid=email,
                values=(i, info["name"], info["vote"], info.get("actual_attend", ""),
                        free_display, amount_display))
        self._update_attendance_totals()

    def _sync_attendance_amount(self, email):
        """Recomputes one person's Amount from their current Actual
        Attend / Free state: Free always wins (Amount forced to 0,
        regardless of attendance); otherwise Amount is auto-pulled from
        Tab 1's "Expected event budget" if Actual Attend is "Yes", or 0
        otherwise. Called after editing/pasting/toggling either field —
        any Amount typed in by hand afterward stays until Actual Attend or
        Free changes again."""
        info = self._attendance_roster[email]
        if info.get("free"):
            info["amount"] = 0.0
        elif (info.get("actual_attend") or "").strip().lower() == "yes":
            info["amount"] = parse_amount_from_text(self.var_budget.get())
        else:
            info["amount"] = 0.0

    def _commit_attendance_edit(self, row_id, col_name, new_value):
        email = row_id
        if email not in self._attendance_roster:
            return
        info = self._attendance_roster[email]
        new_value = new_value.strip()
        if col_name == "name":
            info["name"] = new_value
        elif col_name == "vote":
            info["vote"] = new_value
        elif col_name == "actual_attend":
            info["actual_attend"] = new_value
            self._sync_attendance_amount(email)
        elif col_name == "free":
            # Reached via paste (Ctrl+V), not the single-click toggle —
            # accepts a few common truthy spellings so pasting a column
            # copied from Excel (Yes/TRUE/1/✅) works as expected.
            info["free"] = new_value.lower() in ("yes", "true", "1", "✅", "x")
            self._sync_attendance_amount(email)
        elif col_name == "amount":
            info["amount"] = parse_amount_from_text(new_value)
        self._render_attendance_tree()
        # MỚI: auto-save ngay xuống Attendance_Payment_{EventID}.xlsx sau
        # MỌI thay đổi (double-click edit lẫn paste), không cần bấm nút
        # riêng — best-effort, lỗi (nếu có, vd file đang mở ở chỗ khác) bị
        # bỏ qua âm thầm để không làm gián đoạn thao tác chỉnh sửa bình
        # thường (xem docstring _save_attendance_sheet_to_file()).
        self._save_attendance_sheet_to_file(silent=True)

    def _on_attendance_tree_double_click(self, event):
        """Double-click dispatcher for the Attendance & Payment table:
        "Actual Attend" opens a readonly Yes/No dropdown (see
        _begin_cell_edit_combobox()); "Name"/"Vote"/"Amount" open a normal
        type-in-place Entry (see _begin_cell_edit()); "No." and "Free" are
        not editable via double-click ("Free" toggles on single-click
        instead — see _on_attendance_free_click())."""
        tree = self.tree_attendance
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
        if col_name == "actual_attend":
            self._begin_cell_edit_combobox(tree, row_id, col_name, ["Yes", "No"], self._commit_attendance_edit)
        elif col_name in ("name", "vote", "amount"):
            self._begin_cell_edit(tree, row_id, col_name, self._commit_attendance_edit)
        # "no" and "free" columns: not double-click editable, ignored here

    def _on_attendance_free_click(self, event):
        """Single-click toggle for the "Free" column (simulated checkbox,
        same ✅/⬜ pattern as the Gift Contribution tab's "Contributed"
        column) — marking someone Free forces their Amount to 0 regardless
        of Actual Attend, for people exempt from the contribution."""
        region = self.tree_attendance.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.tree_attendance.identify_column(event.x)
        row_id = self.tree_attendance.identify_row(event.y)
        if not row_id or col != "#5":  # "#5" = the "free" column (no,name,vote,actual_attend,free,amount)
            return
        email = row_id
        if email not in self._attendance_roster:
            return
        info = self._attendance_roster[email]
        info["free"] = not info.get("free", False)
        self._sync_attendance_amount(email)
        self._render_attendance_tree()
        self._save_attendance_sheet_to_file(silent=True)

    def _on_attendance_delete_key(self, event):
        """Delete/Backspace on selected row(s) — clears that row's
        tracking fields (Actual Attend → blank, Free → off, Amount → 0)
        back to an untouched state. Name and Vote are left as-is (they're
        identity data pulled from Tab 4, not meant to be blanked by a
        stray Delete press); to change those, double-click and type a new
        value, or leave them blank by hand via that same edit box.
        Treeview only supports selecting whole ROWS (not individual
        cells), so this clears the row's editable tracking fields as a
        group rather than a single cell — the closest practical match to
        "delete a cell's content" the widget allows."""
        changed = False
        for row_id in self.tree_attendance.selection():
            if row_id not in self._attendance_roster:
                continue
            info = self._attendance_roster[row_id]
            info["actual_attend"] = ""
            info["free"] = False
            info["amount"] = 0.0
            changed = True
        if changed:
            self._render_attendance_tree()
            self._save_attendance_sheet_to_file(silent=True)
        return "break"

    def _update_attendance_totals(self):
        rows = self._attendance_roster.values()
        self.var_total_actual_attend.set(str(count_actual_attendees(rows)))
        self.var_total_collected_amount.set(format_amount(sum_contributions(rows)))
        self._refresh_remaining_amount()

    def _refresh_remaining_amount(self):
        """Remaining amount = Total collected amount − Amount paid. Called
        both when the Attendance table changes (Total collected amount
        moves) and when the user types into "Amount paid" directly (see
        the trace_add on self.var_amount_paid) — either side changing
        should immediately update this."""
        if not hasattr(self, "var_amount_paid") or not hasattr(self, "var_remaining_amount"):
            return
        total_collected = parse_amount_from_text(self.var_total_collected_amount.get())
        amount_paid = parse_amount_from_text(self.var_amount_paid.get())
        self.var_remaining_amount.set(
            format_amount(remaining_amount(total_collected, amount_paid)))

    def _on_amount_paid_changed(self):
        """Recomputes Remaining amount and auto-saves "Amount paid" as the user
        types, to the event it belongs to (self._amount_paid_event) — never
        to whatever Tab 1 shows. It is money, so a save that fails is
        reported, once per run rather than on every keystroke: it used to
        swallow every exception, losing the figure while the UI still showed
        it as entered. It only UPDATEs: an event that is not in History yet
        gets no row of its own, and the user is told to save it first."""
        self._refresh_remaining_amount()
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
                      f"('💾 Update this event in History'), then re-enter Amount paid.")
        if saved:
            self._amount_paid_save_failed = False
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
        self._refresh_remaining_amount()

    def _refresh_amount_paid_for_current_event(self):
        """Tab 5 opened: when the Amount paid on screen belongs to another
        event than Tab 1's, show the one saved for Tab 1's event instead."""
        current_id = self.var_event_id.get().strip()
        if self._amount_paid_event == (current_id or None):
            return
        saved = "0"
        if current_id:
            try:
                rec = next((r for r in db.load_history(self.history_path.get())
                            if r.get("EventID") == current_id), None)
            except Exception:
                rec = None
            if rec and rec.get("AmountPaid"):
                saved = rec["AmountPaid"]
        self._amount_paid_event = current_id or None
        self._set_amount_paid_quietly(saved)

    # ── Attendance & Payment / Responded result — lưu vào database ──
    # MỚI: đã đổi từ file Excel Attendance_Payment_{EventID}.xlsx (2 sheet)
    # sang lưu THẲNG vào database (bảng attendance + responses trong
    # db.py) — auto-save mỗi khi có thay đổi, không tự ghi Excel liên tục
    # nữa. Muốn có file Excel để báo cáo/gửi người khác, dùng nút
    # "📊 Export to Excel" (xem _export_attendance_to_excel()). Cả 2 đều
    # được đọc lại tự động bởi '⬅ Load setup from selected event' trên
    # Tab 1, không cần quét lại Outlook để xem lại dữ liệu cũ.

    def _save_attendance_sheet_to_file(self, silent=True):
        """Lưu bảng Attendance & Payment hiện tại (self._attendance_roster)
        vào database — gọi TỰ ĐỘNG sau mọi thay đổi (double-click, toggle
        Free, paste, Delete). Tên hàm giữ nguyên như cũ (dù giờ không còn
        ghi "file" Excel nữa) để không phải sửa lại các nơi đang gọi nó."""
        event_id = self._attendance_event  # the event this table belongs to
        if not event_id or not self._attendance_roster:
            return
        try:
            db.save_attendance_roster(event_id, self._attendance_roster, self.history_path.get())
        except Exception:
            if not silent:
                raise
            pass  # best-effort silent auto-save

    def _load_attendance_sheet_from_file(self, event_id):
        """Đọc bảng Attendance & Payment đã lưu trong database cho
        event_id, nạp vào self._attendance_roster và vẽ lại — dùng bởi
        '⬅ Load setup from selected event'. Trả về True nếu có dữ liệu để
        nạp, False nếu chưa từng lưu gì cho sự kiện này (không phải lỗi)."""
        try:
            roster = db.load_attendance_roster(event_id, self.history_path.get())
        except Exception:
            roster = {}
        self._attendance_roster = roster
        self._attendance_event = event_id
        self._render_attendance_tree()
        return bool(roster)

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
        return reports.attendance_workbook(
            self._attendance_roster, parse_amount_from_text(self.var_amount_paid.get()))

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
        """Lấy đúng (event_name, event_date, location, total_attend,
        total_collected, amount_paid, remaining_amount) hiện tại từ Tab 1
        + bảng Attendance & Payment (Tab 5) — dùng để build nội dung email
        cảm ơn mặc định (build_thankyou_body()). Các số liệu Attendance
        luôn đọc TRỰC TIẾP từ 3 StringVar hiển thị trên UI (đã được
        _update_attendance_totals()/_refresh_remaining_amount() giữ luôn
        cập nhật), tránh phải tính lại từ self._attendance_roster ở đây."""
        return (
            self.var_event_name.get(),
            get_date_str(self.date_event),
            self.var_location.get(),
            getattr(self, "var_total_actual_attend", tk.StringVar(value="0")).get(),
            getattr(self, "var_total_collected_amount", tk.StringVar(value="0")).get(),
            getattr(self, "var_amount_paid", tk.StringVar(value="0")).get(),
            getattr(self, "var_remaining_amount", tk.StringVar(value="0")).get(),
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

        def worker():
            try:
                mail, calendar_attached = self.outlook.send_thankyou_email(
                    actual_attendees, subject, body,
                    excel_path=excel_path, event_name=event_name,
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
        top = ttk.Frame(f)
        top.pack(fill="x", padx=10, pady=8)
        ttk.Label(top, text="Database file:").pack(side="left")
        ttk.Entry(top, textvariable=self.history_path, width=40).pack(side="left", padx=6)
        ttk.Button(top, text="📂 Browse...", command=self._browse_history_file).pack(side="left", padx=4)
        ttk.Button(top, text="🔄 Reload", command=self._refresh_history_tree).pack(side="left", padx=4)
        # Double-click any cell below to edit it directly, then click this
        # button to write your changes back to the database (see
        # _save_history_edits()). Nothing is written until you click Save —
        # double-clicking only edits what's shown on screen.
        ttk.Button(top, text="💾 Save changes", command=self._save_history_edits)\
            .pack(side="left", padx=4)
        # The database is the working copy; this writes a snapshot file.
        ttk.Button(top, text="📊 Export to Excel", command=self._export_history_to_excel)\
            .pack(side="left", padx=4)

        self._make_wrapping_label(f, text="💡 Double-click any cell below to edit it in place (or select rows and use "
                          "Ctrl+C / Ctrl+V to copy/paste to or from Excel), then click "
                          "'💾 Save changes' to write the cells you edited back to the database. "
                          "(Renaming an Event ID moves the whole event — its recipients, votes, attendance "
                          "and gift list — to the new ID.) Use '📊 Export to Excel' any time you need a "
                          "report file to share or archive; it includes every stored column.",
                  font=("Arial", 8, "italic")).pack(anchor="w", padx=10, pady=(0, 4))

        cols = tuple(c for c, _h, _w in HISTORY_TABLE_COLUMNS)
        tree_container, self.tree_history = make_scrollable_treeview(f, columns=cols, height=18)
        for c, h, w in HISTORY_TABLE_COLUMNS:
            self.tree_history.heading(c, text=h)
            self.tree_history.column(c, width=w)
        tree_container.pack(fill="both", expand=True, padx=10, pady=6)
        # Every column is editable via double-click (editable_cols=None ->
        # no restriction) — see _commit_history_edit(). This only updates
        # what's shown in the table; nothing is saved to disk until
        # '💾 Save changes' is clicked.
        self.tree_history.bind(
            "<Double-1>",
            lambda e: self._on_editable_tree_double_click(
                self.tree_history, e, None, self._commit_history_edit))
        # Select rows and use Ctrl+C / Ctrl+V to copy/paste to or from
        # Excel — pasted cells go through the same _commit_history_edit()
        # used by double-click, so it only edits what's on screen; nothing
        # is written to the database until '💾 Save changes' is clicked.
        self._enable_treeview_copy_paste(self.tree_history, on_commit=self._commit_history_edit)

        self._history_edits = {}    # row iid -> {column: edited value}, not saved yet
        self._history_row_ids = {}  # row iid -> the EventID that row has in the database
        self._refresh_history_tree()

    def _browse_history_file(self):
        path = filedialog.askopenfilename(filetypes=[("SQLite database", "*.db"), ("All files", "*.*")])
        if path:
            self.history_path.set(path)
            self._history_edits = {}
            self._refresh_history_tree()

    def _refresh_history_tree(self):
        """Redraws Tab 7 from the database. Edits not saved yet are laid back
        on top, so an automatic write elsewhere (a scan, a reminder) that
        refreshes this table never throws them away."""
        self.tree_history.delete(*self.tree_history.get_children())
        try:
            records = db.load_history(self.history_path.get())
        except Exception:
            records = []
        cols = self.tree_history["columns"]
        self._history_row_ids = {}
        for rec in records:
            # The row's iid is its ORIGINAL Event ID, so an edited Event ID
            # cell can still be traced back to its row in the database.
            base_iid = (rec.get("EventID") or "").strip() or "(blank)"
            iid = base_iid
            suffix = 2
            while self.tree_history.exists(iid):
                iid = f"{base_iid}__{suffix}"
                suffix += 1
            self._history_row_ids[iid] = rec.get("EventID")
            pending = self._history_edits.get(iid, {})
            values = [pending.get(c, "" if rec.get(c) is None else rec.get(c)) for c in cols]
            self.tree_history.insert("", "end", iid=iid, values=values)
        # Edits of rows that no longer exist cannot be saved anywhere.
        self._history_edits = {iid: e for iid, e in self._history_edits.items()
                               if iid in self._history_row_ids}

    def _commit_history_edit(self, row_id, col_name, new_value):
        """Called after double-click editing (or pasting into) a cell on Tab
        7 — only updates the table and remembers the edit. Nothing touches
        the database until _save_history_edits() runs."""
        if not self.tree_history.exists(row_id):
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
        for attr in ("_last_scanned_event_id", "_attendance_event", "_gift_event", "_amount_paid_event"):
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

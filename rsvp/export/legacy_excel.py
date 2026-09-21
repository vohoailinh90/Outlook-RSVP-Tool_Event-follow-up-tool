"""One-time import of the pre-SQLite Excel history.

Moved verbatim out of db.py by the phase 2 extraction in
docs/agentic/ARCHITECTURE.md. It lived there because it writes to the
database, but it READS Excel, which made the storage layer depend on
openpyxl - the one dependency db.py otherwise does not have. db.py is
now stdlib-only.

Runs only when the database is empty and an old RSVP_History.xlsx is
present, so for anyone already using the database it does nothing.
"""

import os
from datetime import datetime

from db import (  # the storage layer this writes into
    DB_FILE_DEFAULT,
    EVENT_COLUMNS,
    get_connection,
    save_attendance_roster,
    save_event_record,
    save_gift_roster,
    save_recipients,
    save_responses,
)


def migrate_from_excel_if_needed(db_path=DB_FILE_DEFAULT, history_xlsx="RSVP_History.xlsx", search_dir=None):
    """Chỉ chạy MỘT LẦN DUY NHẤT: nếu DB đã có ít nhất 1 sự kiện rồi (tức
    đã từng migrate hoặc đã dùng DB từ đầu), KHÔNG làm gì cả — an toàn khi
    gọi lại nhiều lần (vd mỗi lần mở app). Nếu DB đang trống VÀ tìm thấy
    RSVP_History.xlsx cũ, đọc toàn bộ sự kiện từ đó, rồi với MỖI Event ID,
    tìm thêm Participant_List_{EventID}.xlsx / Gift_Contribution_List_
    {EventID}.xlsx / Attendance_Payment_{EventID}.xlsx (sheet "Attendance &
    Payment" + "Responded result") CÙNG THƯ MỤC để nhập nốt Recipients/
    Gift/Attendance/Responses tương ứng — best-effort, thiếu file nào thì
    bỏ qua file đó, không chặn việc nhập các phần còn lại.

    Trả về (migrated: bool, event_count: int, notes: list[str]) để app
    hiển thị cho user biết đã nhập được những gì."""
    notes = []
    conn = get_connection(db_path)
    try:
        existing_count = conn.execute("SELECT COUNT(*) FROM events").fetchone()[0]
    finally:
        conn.close()
    if existing_count > 0:
        return False, 0, notes  # DB đã có dữ liệu -> không migrate lại, tránh nhân đôi

    if not os.path.exists(history_xlsx):
        return False, 0, notes  # không có gì để nhập — coi như bắt đầu mới hoàn toàn

    try:
        import openpyxl
    except ImportError:
        notes.append("openpyxl chưa cài — không migrate được.")
        return False, 0, notes

    search_dir = search_dir or os.path.dirname(os.path.abspath(history_xlsx)) or "."

    try:
        wb = openpyxl.load_workbook(history_xlsx, data_only=True)
        ws = wb["History"] if "History" in wb.sheetnames else wb.active
        # File cũ có thể chỉ 24 cột (bản trước khi thêm Event mode/Gift) —
        # zip với EVENT_COLUMNS và lấy None cho cột thiếu, giống hệt cách
        # history.py cũ tự "migrate" file thiếu cột.
        header_row = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
        rows = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]:
                continue
            rec = dict(zip(header_row, row))
            rows.append(rec)
    except Exception as e:
        notes.append(f"Không đọc được {history_xlsx}: {e}")
        return False, 0, notes

    imported = 0
    for rec in rows:
        event_id = rec.get("EventID")
        if not event_id:
            continue
        clean_rec = {c: rec.get(c, "") for c in EVENT_COLUMNS}
        clean_rec["EventID"] = event_id
        save_event_record(clean_rec, db_path)
        imported += 1

        # Recipients — từ Participant_List_{EventID}.xlsx (đường dẫn ghi
        # trong cột RecipientFile, hoặc suy ra theo tên chuẩn nếu cột trống).
        # BUG ĐÃ SỬA: file này dùng ĐÚNG format của _load_recipients() /
        # _export_recipients_to_excel() trong rsvp_app.py — sheet "DanhSach"
        # (hoặc active nếu không có), HEADER Ở DÒNG 3 ("Họ tên"/"Email"), dữ
        # liệu bắt đầu DÒNG 4 — KHÔNG PHẢI header ở dòng 1 như bản migration
        # trước đó lỡ giả định (khiến không tìm thấy cột nào, bỏ qua hết
        # người nhận — đúng như lỗi "Recipients trống sau khi Load setup").
        # Cũng bỏ qua đúng dòng "(ví dụ)"/"(example)" của file mẫu TEMPLATE,
        # y hệt _load_recipients().
        recipient_path = rec.get("RecipientFile") or os.path.join(
            search_dir, f"Participant_List_{event_id}.xlsx")
        if recipient_path and os.path.exists(recipient_path):
            try:
                rwb = openpyxl.load_workbook(recipient_path, data_only=True)
                rws = rwb["DanhSach"] if "DanhSach" in rwb.sheetnames else rwb.active
                people = []
                for r in rws.iter_rows(min_row=4, values_only=True):
                    name, email = (r + (None, None))[:2] if r else (None, None)
                    if not email or "@" not in str(email):
                        continue
                    if "(ví dụ)" in str(name or "") or "(example)" in str(name or "").lower():
                        continue
                    people.append((str(name or "").strip(), str(email).strip()))
                if people:
                    save_recipients(event_id, people, db_path)
            except Exception:
                notes.append(f"[{event_id}] không nhập được danh sách người nhận.")

        # Gift Contribution — từ Gift_Contribution_List_{EventID}.xlsx
        gift_path = os.path.join(search_dir, f"Gift_Contribution_List_{event_id}.xlsx")
        if os.path.exists(gift_path):
            try:
                gwb = openpyxl.load_workbook(gift_path, data_only=True)
                gws = gwb.active
                gheader_row = next(gws.iter_rows(min_row=1, max_row=1), ())
                gheader = {(str(c.value).strip() if c.value else ""): i for i, c in enumerate(gheader_row)}
                i_name, i_email = gheader.get("Name"), gheader.get("Email")
                i_amount, i_contrib = gheader.get("Amount"), gheader.get("Contributed")
                roster = {}
                if i_email is not None:
                    for r in gws.iter_rows(min_row=2, values_only=True):
                        if not r or i_email >= len(r) or not r[i_email]:
                            continue
                        email = str(r[i_email]).strip()
                        roster[email] = {
                            "name": str(r[i_name]).strip() if i_name is not None and r[i_name] else email,
                            "checked": (str(r[i_contrib]).strip().lower() == "yes") if i_contrib is not None and i_contrib < len(r) and r[i_contrib] is not None else False,
                            "amount": float(r[i_amount]) if i_amount is not None and i_amount < len(r) and isinstance(r[i_amount], (int, float)) else 0.0,
                        }
                if roster:
                    save_gift_roster(event_id, roster, db_path)
            except Exception:
                notes.append(f"[{event_id}] không nhập được danh sách quyên góp quà.")

        # Attendance & Payment + Responded result — từ
        # Attendance_Payment_{EventID}.xlsx (2 sheet)
        attn_path = os.path.join(search_dir, f"Attendance_Payment_{event_id}.xlsx")
        if os.path.exists(attn_path):
            try:
                awb = openpyxl.load_workbook(attn_path, data_only=True)
                if "Attendance & Payment" in awb.sheetnames:
                    aws = awb["Attendance & Payment"]
                    aheader_row = next(aws.iter_rows(min_row=1, max_row=1), ())
                    aheader = {(str(c.value).strip() if c.value else ""): i for i, c in enumerate(aheader_row)}
                    i_email = aheader.get("Email")
                    roster = {}
                    if i_email is not None:
                        for r in aws.iter_rows(min_row=2, values_only=True):
                            if not r or i_email >= len(r) or not r[i_email]:
                                continue
                            email = str(r[i_email]).strip()
                            g = lambda key: (r[aheader[key]] if key in aheader and aheader[key] < len(r) else None)
                            amount = g("Amount")
                            roster[email] = {
                                "name": str(g("Name")).strip() if g("Name") else email,
                                "vote": str(g("Vote")).strip() if g("Vote") else "",
                                "actual_attend": str(g("Actual Attend")).strip() if g("Actual Attend") else "",
                                "free": str(g("Free")).strip().lower() == "yes" if g("Free") else False,
                                "amount": float(amount) if isinstance(amount, (int, float)) else 0.0,
                            }
                    if roster:
                        save_attendance_roster(event_id, roster, db_path)
                if "Responded result" in awb.sheetnames:
                    rws2 = awb["Responded result"]
                    rheader_row = next(rws2.iter_rows(min_row=1, max_row=1), ())
                    rheader = {(str(c.value).strip() if c.value else ""): i for i, c in enumerate(rheader_row)}
                    i_email = rheader.get("Email")
                    responses = {}
                    if i_email is not None:
                        for r in rws2.iter_rows(min_row=2, values_only=True):
                            if not r or i_email >= len(r) or not r[i_email]:
                                continue
                            email = str(r[i_email]).strip()
                            g = lambda key: (r[rheader[key]] if key in rheader and rheader[key] < len(r) else None)
                            received_raw = g("Received At")
                            received_dt = None
                            if received_raw:
                                try:
                                    received_dt = datetime.strptime(str(received_raw), "%Y-%m-%d %H:%M")
                                except Exception:
                                    received_dt = None
                            responses[email.lower()] = {
                                "name": str(g("Name")).strip() if g("Name") else "",
                                "vote": str(g("Vote")).strip() if g("Vote") else "",
                                "received": received_dt,
                            }
                    if responses:
                        save_responses(event_id, responses, None, db_path)
            except Exception:
                notes.append(f"[{event_id}] không nhập được dữ liệu Attendance/Responded.")

    return True, imported, notes

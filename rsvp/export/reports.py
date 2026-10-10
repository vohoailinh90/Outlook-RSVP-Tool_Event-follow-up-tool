"""The Excel reports the app writes and the gift list it reads back.

Moved out of RSVPApp so they can be tested: these files are exported and
attached to emails sent to colleagues (the attendance report rides on the
Thank-you email, the contributor list on the gift report), and a wrong
figure in them is a wrong figure in someone's inbox.

Rosters are the dicts the app keeps per email address:
    attendance: {"name", "vote", "actual_attend", "free", "amount"}
    gift:       {"name", "checked", "amount", "send_email"}
"""
from __future__ import annotations

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

from rsvp.domain import parse_amount_from_text, remaining_amount


def write_header_row(ws, headers):
    """Row 1 in the app's report style: white bold text on dark blue."""
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="003366")
        c.alignment = Alignment(horizontal="center")


def attendance_workbook(roster, amount_paid):
    """The Attendance & Payment report (Tab 5): one row per Yes/Maybe voter,
    then Total actual attend, Total collected amount, Amount paid and
    Remaining amount. Used by "📊 Export to Excel" and as the Thank-you
    email's attachment.

    Remaining is computed from the total and `amount_paid`, never re-read from
    the label on screen: parse_amount_from_text drops the sign, so an overpaid
    "-5,000" used to be written into this report as 5,000."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Attendance & Payment"
    write_header_row(ws, ["No.", "Name", "Email", "Vote", "Actual Attend", "Free", "Amount"])
    row_idx = 2
    total_attend = 0
    total_amount = 0.0
    for i, (email, info) in enumerate(roster.items(), start=1):
        attend = (info.get("actual_attend") or "").strip()
        amount = info.get("amount", 0.0)
        if attend.lower() == "yes":
            total_attend += 1
        total_amount += amount
        ws.cell(row=row_idx, column=1, value=i)
        ws.cell(row=row_idx, column=2, value=info["name"])
        ws.cell(row=row_idx, column=3, value=email)
        ws.cell(row=row_idx, column=4, value=info.get("vote", ""))
        ws.cell(row=row_idx, column=5, value=attend)
        ws.cell(row=row_idx, column=6, value="Yes" if info.get("free") else "No")
        ws.cell(row=row_idx, column=7, value=amount or None)
        row_idx += 1
    total_row = row_idx + 1
    ws.cell(row=total_row, column=2, value="Total actual attend:").font = Font(bold=True)
    ws.cell(row=total_row, column=5, value=total_attend).font = Font(bold=True)
    amount_row = total_row + 1
    ws.cell(row=amount_row, column=2, value="Total collected amount:").font = Font(bold=True)
    c_amt = ws.cell(row=amount_row, column=7, value=total_amount)
    c_amt.font = Font(bold=True)
    c_amt.number_format = "#,##0"
    paid_row = amount_row + 1
    ws.cell(row=paid_row, column=2, value="Amount paid:").font = Font(bold=True)
    c_paid = ws.cell(row=paid_row, column=7, value=amount_paid)
    c_paid.font = Font(bold=True)
    c_paid.number_format = "#,##0"
    remaining_row = paid_row + 1
    ws.cell(row=remaining_row, column=2, value="Remaining amount:").font = Font(bold=True)
    c_rem = ws.cell(row=remaining_row, column=7, value=remaining_amount(total_amount, amount_paid))
    c_rem.font = Font(bold=True)
    c_rem.number_format = "#,##0"
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 10
    ws.column_dimensions["G"].width = 14
    return wb


def gift_contribution_workbook(roster):
    """The full gift list (Tab 6 "📊 Export to Excel"): everyone, with a
    Contributed Yes/No column, and a TOTAL COLLECTED row of the contributors'
    amounts. read_gift_contribution_rows() reads it back."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Gift Contribution"
    write_header_row(ws, ["No.", "Name", "Email", "Amount", "Contributed"])
    row_idx = 2
    total_amount = 0.0
    for i, (email, info) in enumerate(roster.items(), start=1):
        amount = info.get("amount", 0.0)
        if info["checked"]:
            total_amount += amount
        ws.cell(row=row_idx, column=1, value=i)
        ws.cell(row=row_idx, column=2, value=info["name"])
        ws.cell(row=row_idx, column=3, value=email)
        ws.cell(row=row_idx, column=4, value=amount or None)
        ws.cell(row=row_idx, column=5, value="Yes" if info["checked"] else "No")
        row_idx += 1

    # Grand total row, one blank row below the table. The label goes in the
    # NAME column (not Email) so read_gift_contribution_rows() skips this row
    # on reload: it keys off the Email column being non-empty.
    total_row = row_idx + 1
    label_cell = ws.cell(row=total_row, column=2, value="TOTAL COLLECTED:")
    label_cell.font = Font(bold=True)
    label_cell.alignment = Alignment(horizontal="right")
    total_cell = ws.cell(row=total_row, column=4, value=total_amount)
    total_cell.font = Font(bold=True)
    total_cell.number_format = "#,##0"

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 32
    ws.column_dimensions["D"].width = 14
    ws.column_dimensions["E"].width = 14
    return wb


def gift_report_workbook(roster):
    """The attachment of the gift report email: only the people who HAVE
    contributed, renumbered from 1, as No./Name/Email/Amount - the "Send
    email"/"Contributed" checkboxes are UI state, not report data."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Contributors"
    write_header_row(ws, ["No.", "Name", "Email", "Amount"])
    row_idx = 2
    seq = 0
    total_amount = 0.0
    for email, info in roster.items():
        if not info.get("checked"):
            continue
        seq += 1
        amount = info.get("amount", 0.0)
        total_amount += amount
        ws.cell(row=row_idx, column=1, value=seq)
        ws.cell(row=row_idx, column=2, value=info.get("name") or email)
        ws.cell(row=row_idx, column=3, value=email)
        ws.cell(row=row_idx, column=4, value=amount or None)
        row_idx += 1
    total_row = row_idx + 1
    ws.cell(row=total_row, column=2, value="TOTAL COLLECTED:").font = Font(bold=True)
    c_amt = ws.cell(row=total_row, column=4, value=total_amount)
    c_amt.font = Font(bold=True)
    c_amt.number_format = "#,##0"
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 32
    ws.column_dimensions["D"].width = 14
    return wb


def read_gift_contribution_rows(path):
    """Reads a Gift_Contribution_List_*.xlsx file and returns
    list[dict(name, email, checked, amount)], matching columns by their
    HEADER text (row 1) instead of a fixed position — so both older files
    (Name/Email/Contributed only) and newer files (No./Name/Email/Amount/
    Contributed) can be read back correctly."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    header_row = next(ws.iter_rows(min_row=1, max_row=1), ())
    header = {(str(c.value).strip() if c.value else ""): i for i, c in enumerate(header_row)}
    i_name = header.get("Name")
    i_email = header.get("Email")
    i_amount = header.get("Amount")
    i_contrib = header.get("Contributed")
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        email = row[i_email] if i_email is not None and i_email < len(row) else None
        if not email:
            continue  # skips blank rows and the "TOTAL COLLECTED" summary row
        name = row[i_name] if i_name is not None and i_name < len(row) else None
        contributed = row[i_contrib] if i_contrib is not None and i_contrib < len(row) else None
        amount = row[i_amount] if i_amount is not None and i_amount < len(row) else None
        rows.append({
            "name": str(name).strip() if name else str(email).strip(),
            "email": str(email).strip(),
            "checked": (str(contributed).strip().lower() == "yes") if contributed is not None else False,
            "amount": parse_amount_from_text(amount) if amount is not None else 0.0,
        })
    return rows

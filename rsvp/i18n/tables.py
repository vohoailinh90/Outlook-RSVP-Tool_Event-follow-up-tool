"""Aligned text tables for email bodies, and their conversion to HTML.

Ported unchanged from the separately developed copy of this app (the user's
live one): the thank-you email's table of payment rounds and the gift
report's Event + Gift table are text the user can edit before sending, and
text_body_to_html() turns every block of lines containing "│" into a real
<table> when the mail is created, because Outlook's proportional font breaks
space-aligned columns.
"""
from __future__ import annotations

import re
import unicodedata


def _display_width(text):
    """Bề rộng HIỂN THỊ thật của 1 chuỗi, tính theo số ô ký tự trên màn
    hình — KHÁC len(): emoji (👥💰💸📊) và chữ Nhật (集金額, 参加人数...) mỗi
    ký tự chiếm 2 ô chứ không phải 1. Nếu canh cột bằng len() thì bảng tổng
    kết trong email cảm ơn sẽ lệch hàng ngay khi có emoji hoặc tiếng Nhật —
    đúng lỗi đã gặp. east_asian_width trả 'W' (Wide) / 'F' (Fullwidth) cho
    các ký tự loại này; riêng emoji ở một số dải Unicode mới bị xếp là 'N',
    nên bắt thêm bằng khoảng mã."""
    width = 0
    for ch in text:
        if unicodedata.east_asian_width(ch) in ("W", "F"):
            width += 2
        elif 0x1F300 <= ord(ch) <= 0x1FAFF or 0x2600 <= ord(ch) <= 0x27BF:
            width += 2  # emoji / pictographs (📊 💰 💸 👥 ✅ ...)
        elif unicodedata.combining(ch):
            width += 0  # dấu thanh tiếng Việt tổ hợp — không chiếm ô riêng
        else:
            width += 1
    return width


def _pad_display(text, width, align="left"):
    """ljust/rjust nhưng đệm theo _display_width() thay vì len()."""
    padding = " " * max(0, width - _display_width(text))
    return (text + padding) if align == "left" else (padding + text)


def _build_rounds_table(L, rounds_info):
    """MỚI: dựng BẢNG tổng kết các đợt thu tiền dưới dạng text canh cột đều
    nhau (Round | Attendees | Collected | Paid | Remaining), thay cho danh
    sách gạch đầu dòng khó đọc trước đây. Bề rộng mỗi cột tự co giãn theo
    nội dung DÀI NHẤT của chính cột đó (kể cả tiêu đề), đo bằng
    _display_width() để emoji/tiếng Nhật không làm lệch hàng.

    Dùng text thuần (không HTML) vì email cảm ơn được gửi qua Body của
    Outlook chứ không phải HTMLBody — các cột tách bạch rõ nhờ dấu │ ngăn
    giữa, và thẳng hàng tuyệt đối khi người nhận đọc bằng font monospace."""
    headers = [L["col_round"], L["col_attend"], L["col_collected"], L["col_paid"], L["col_remaining"]]
    rows = [[str(x) for x in r] for r in rounds_info]
    widths = [
        max([_display_width(headers[i])] + [_display_width(row[i]) for row in rows])
        for i in range(5)
    ]

    def fmt(cells):
        # Cột đầu (tên đợt) canh TRÁI, 4 cột số canh PHẢI cho dễ so sánh.
        out = [_pad_display(cells[0], widths[0], "left")]
        out += [_pad_display(cells[i], widths[i], "right") for i in range(1, 5)]
        return "  " + " │ ".join(out)

    sep = "  " + "─┼─".join("─" * w for w in widths)
    lines = [fmt(headers), sep] + [fmt(row) for row in rows]
    return "\n".join(lines)


def _build_aligned_table(headers, rows):
    """MỚI: dựng 1 bảng text canh cột đều nhau với SỐ CỘT BẤT KỲ — bản tổng
    quát của _build_rounds_table() (vốn cố định 5 cột cho email cảm ơn).
    Dùng cho bảng tổng kết Event + Gift trong email báo cáo quyên góp
    (Tab 6), vốn chỉ có 4 cột.

    Vẫn giữ đúng 2 quy ước cũ để text_body_to_html() nhận diện được và dựng
    lại thành <table> HTML thật lúc gửi:
      • các ô ngăn nhau bằng " │ ",
      • dòng kẻ ngang dùng "─┼─" (bắt buộc có "┼", nếu không sẽ bị nhầm với
        đường phân cách song ngữ "――――" và bị xoá oan).
    Bề rộng cột đo bằng _display_width() chứ KHÔNG phải len(), vì tiêu đề
    có emoji + tiếng Nhật."""
    n = len(headers)
    rows = [[str(c) for c in row] for row in rows]
    widths = [
        max([_display_width(str(headers[i]))] + [_display_width(row[i]) for row in rows])
        for i in range(n)
    ]

    def fmt(cells):
        # Cột đầu (tên dòng: Party/Gift) canh TRÁI, các cột số canh PHẢI.
        out = [_pad_display(str(cells[0]), widths[0], "left")]
        out += [_pad_display(str(cells[i]), widths[i], "right") for i in range(1, n)]
        return "  " + " │ ".join(out)

    sep = "  " + "─┼─".join("─" * w for w in widths)
    return "\n".join([fmt(headers), sep] + [fmt(row) for row in rows])


# Dò URL trong CHUỖI ĐÃ ESCAPE (xem esc() trong text_body_to_html()).
# Dừng ở khoảng trắng và ở "<>\"" — sau khi escape thì 3 ký tự đó không thể
# còn nằm giữa 1 URL nữa, nên chúng là mốc kết thúc an toàn. Dấu "&" trong
# URL lúc này đã là "&amp;" (không có khoảng trắng) nên vẫn được gộp trọn
# vào link — quan trọng với các link nhiều tham số như của Rakuten.
_URL_IN_TEXT_RE = re.compile(r'https?://[^\s<>"]+')
# Dấu câu dính ở CUỐI câu chứ không thuộc URL — vd "xem tại https://a.b/c."
# thì dấu chấm cuối là của câu văn. Không cắt ra thì link sẽ hỏng.
_URL_TRAILING_PUNCT = ".,;:!?)]}、。」）"


def _linkify_url_match(m):
    """Biến 1 URL đã dò được thành thẻ <a> mở ở tab mới. Trả lại phần dấu
    câu ở cuối ra NGOÀI thẻ <a>, để dấu chấm cuối câu không bị nuốt vào
    link (làm link 404)."""
    url = m.group(0)
    trail = ""
    while url and url[-1] in _URL_TRAILING_PUNCT:
        trail = url[-1] + trail
        url = url[:-1]
    if not url:
        return m.group(0)
    # word-break: URL đặt hàng thường rất dài (link Rakuten trong ví dụ dài
    # hơn 100 ký tự) — không cho xuống dòng thì nó đẩy toác bố cục email.
    return (f'<a href="{url}" target="_blank" '
            f'style="word-break:break-all;">{url}</a>{trail}')


def text_body_to_html(body):
    """MỚI: chuyển nội dung email dạng TEXT (thứ user nhìn thấy và sửa được
    trong khung soạn thảo) thành HTML để gửi qua mail.HTMLBody.

    LÝ DO: bảng tổng kết các đợt thu tiền được canh cột bằng khoảng trắng
    (xem _build_rounds_table()) chỉ thẳng hàng nếu người NHẬN đọc bằng font
    monospace. Outlook mặc định dùng font tỉ lệ (Calibri...) nên bảng lệch
    tan nát. Giải pháp: giữ nguyên khung soạn thảo dạng text cho dễ sửa,
    nhưng ngay trước khi gửi thì DÒ ra các dòng bảng và dựng lại thành 1
    <table> HTML thật — trình bày chuẩn ở mọi font, mọi client.

    Cách dò: mọi dòng liên tiếp có chứa ký tự "│" gộp thành 1 khối bảng
    (dòng kẻ ngang ─┼─ bị bỏ qua); dòng ĐẦU của khối là tiêu đề. Nhờ vậy
    chế độ song ngữ có 2 bảng vẫn xử lý đúng, và nếu user xoá bảng đi thì
    phần còn lại vẫn gửi bình thường như text thường."""
    def esc(s):
        """Escape HTML RỒI tự biến các URL trần thành thẻ <a> thật.

        BUG ĐÃ SỬA: link đặt hàng món quà (Tab 6) hiện ra trong Outlook chỉ
        là CHỮ THƯỜNG, bấm vào không đi đâu cả. Nguyên nhân: Outlook chỉ tự
        nhận diện URL khi body là TEXT THUẦN; khi đã gán mail.HTMLBody thì
        nó hiển thị đúng y HTML mình đưa vào — không có <a href> thì không
        có link. Vậy nên phải tự dựng thẻ <a> ở đây.

        Thứ tự BẮT BUỘC là escape TRƯỚC rồi mới dò URL: nếu làm ngược lại,
        thẻ <a> vừa tạo sẽ bị chính hàm escape biến thành &lt;a...&gt; và
        hiện ra dưới dạng chữ. Việc escape trước cũng khiến dấu "&" trong
        URL thành "&amp;" — đó CHÍNH LÀ dạng đúng của một href trong HTML
        (trình đọc mail tự giải mã lại thành "&"), nên link nhiều tham số
        như của Rakuten (…&scid=…&_mpt=…) vẫn mở đúng."""
        s = (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
              .replace('"', "&quot;"))
        return _URL_IN_TEXT_RE.sub(_linkify_url_match, s)

    def render_table(rows):
        """rows: list các list ô đã tách. Dòng đầu = tiêu đề. Cột đầu canh
        trái (tên đợt), các cột còn lại canh phải (số liệu)."""
        if not rows:
            return ""
        head, body_rows = rows[0], rows[1:]
        th = "".join(
            f'<th style="border:1px solid #999;padding:4px 10px;background:#f2f2f2;'
            f'text-align:{"left" if i == 0 else "right"};">{esc(c)}</th>'
            for i, c in enumerate(head))
        trs = []
        for row in body_rows:
            tds = "".join(
                f'<td style="border:1px solid #999;padding:4px 10px;'
                f'text-align:{"left" if i == 0 else "right"};">{esc(c)}</td>'
                for i, c in enumerate(row))
            trs.append(f"<tr>{tds}</tr>")
        return (f'<table style="border-collapse:collapse;margin:10px 0;">'
                f"<tr>{th}</tr>{''.join(trs)}</table>")

    out = []
    table_rows = []

    def flush_table():
        if table_rows:
            out.append(render_table(table_rows))
            table_rows.clear()

    for line in body.split("\n"):
        stripped = line.strip()
        # Dòng kẻ ngang phân cách tiêu đề của BẢNG — bỏ đi vì HTML đã có
        # viền riêng. Bắt buộc phải chứa "┼" (dấu giao cột, chỉ xuất hiện ở
        # dòng kẻ của bảng) để không lỡ tay xoá nhầm đường phân cách song
        # ngữ hay bất kỳ dòng gạch ngang nào user tự gõ.
        if "┼" in stripped and set(stripped) <= set("─┼ "):
            continue
        if "│" in line:
            table_rows.append([c.strip() for c in line.split("│")])
            continue
        flush_table()
        out.append(esc(line) + "<br>")
    flush_table()

    return ('<div style="font-family:Segoe UI, Meiryo, sans-serif;font-size:11pt;">'
            + "\n".join(out) + "</div>")

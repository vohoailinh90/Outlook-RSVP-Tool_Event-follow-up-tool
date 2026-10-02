"""
outlook_com.py — Các hàm điều khiển Outlook desktop qua COM (pywin32)
=======================================================================
Chỉ chạy được trên WINDOWS, với Outlook desktop đã cài & đăng nhập.
Được rsvp_app.py (giao diện chính) gọi tới — tách riêng file này để
dễ đọc / dễ sửa mà không đụng vào code giao diện.

⚠️ QUAN TRỌNG: rsvp_app.py gọi các hàm ở đây từ trong background thread
(threading.Thread) để không làm đứng giao diện khi chờ Outlook phản hồi.
pywin32 COM YÊU CẦU mỗi thread dùng COM phải tự gọi pythoncom.CoInitialize()
trước — nếu không sẽ báo lỗi COM khó hiểu (hoặc bị nuốt mất traceback thật).
Vì vậy MỌI hàm public trong file này đều tự CoInitialize()/CoUninitialize()
ở đầu/cuối hàm — không cần rsvp_app.py phải lo việc này.

Cơ chế Voting Buttons:
    - Khi soạn mail, gán mail.VotingOptions = "Yes;No;Maybe"
      → Outlook tự thêm 3 nút bấm ở đầu email cho người nhận.
    - Khi người nhận bấm 1 nút, Outlook TỰ ĐỘNG gửi lại 1 email phản hồi
      rất ngắn (không cần họ gõ chữ gì). Email phản hồi này có thuộc tính
      COM `.VotingResponse` chứa đúng lựa chọn họ đã bấm (vd: "Yes").
    - Vì vậy collect_responses không cần parse text nữa — chỉ cần đọc
      thuộc tính VotingResponse của các email trả lời trong Inbox.
"""
import pythoncom
import time
import os


def _outlook_app():
    import win32com.client
    return win32com.client.Dispatch("Outlook.Application")


def _asset_path(filename):
    """Đường dẫn tuyệt đối tới file asset (vd ảnh minh hoạ) nằm CÙNG THƯ MỤC
    với outlook_com.py — để hoạt động đúng dù app được chạy từ thư mục nào
    (Windows Task Scheduler, shortcut, v.v.)."""
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)


def _attach_vote_illustration(mail):
    """Đính kèm ảnh hướng dẫn cách bấm nút Vote (how_to_vote.png, minh hoạ 3
    bước: mở email → bấm nút Vote → chọn Yes/No/Maybe, có chú thích 3 thứ
    tiếng Anh/Nhật/Việt) vào email, nếu file này tồn tại cùng thư mục với
    outlook_com.py. Dùng cho cả email mời gốc lẫn email nhắc nhở.
    Best-effort — lỗi gì cũng bỏ qua (khg đính kèm được thì thôi), không làm
    hỏng việc gửi email chính."""
    path = _asset_path("how_to_vote.png")
    if os.path.exists(path):
        try:
            mail.Attachments.Add(path)
            return True
        except Exception:
            return False
    return False


def send_voting_invite(recipients, subject, body, voting_options="Yes;No;Maybe", auto_send=False,
                        send_to_override=None, use_voting_buttons=True):
    """
    Tạo email và mở lên để review (mặc định) hoặc gửi luôn.
    recipients: list[(name, email)]
    send_to_override: nếu có (vd: 1 group email của cả phòng ban), dùng chuỗi này
        làm To thay vì nối các email cá nhân trong `recipients`. `recipients` vẫn
        được dùng làm ROSTER để đối chiếu phản hồi sau này (Tab 4), bất kể gửi tới
        group hay từng cá nhân.
    use_voting_buttons: MỚI — mặc định True (email mời RSVP có Voting
        Buttons + ảnh hướng dẫn vote như trước giờ). Đặt False cho email
        THÔNG BÁO THUẦN TUÝ không cần vote (vd "Send Gift Contribution
        Notice" ở Tab 3 — chỉ là thông báo kêu gọi đóng góp, không phải
        1 cuộc bỏ phiếu Yes/No/Maybe) — khi đó KHÔNG gán mail.VotingOptions
        và KHÔNG đính kèm ảnh hướng dẫn vote (vì không có nút gì để hướng
        dẫn bấm cả).
    Trả về: đối tượng MailItem (hoặc None nếu lỗi)
    """
    pythoncom.CoInitialize()
    try:
        outlook = _outlook_app()
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = send_to_override.strip() if send_to_override else ";".join(e for _, e in recipients if e)
        mail.Subject = subject
        mail.Body = body
        if use_voting_buttons:
            mail.VotingOptions = voting_options  # <-- đây chính là "Use Voting Buttons"
            _attach_vote_illustration(mail)  # ảnh minh hoạ cách bấm Vote (best-effort)
        if auto_send:
            mail.Send()
        else:
            mail.Display()
        return mail
    finally:
        pythoncom.CoUninitialize()


def _smtp_address(item):
    """Lấy địa chỉ SMTP thật, kể cả khi Exchange trả về dạng /O=EXCHANGELABS/..."""
    try:
        addr = item.SenderEmailAddress
        if addr and "@" in addr:
            return addr.lower()
        sender = item.Sender
        if sender is not None:
            exch = sender.GetExchangeUser()
            if exch is not None:
                return exch.PrimarySmtpAddress.lower()
    except Exception:
        pass
    return (getattr(item, "SenderEmailAddress", None) or "unknown").lower()


def list_folder_paths(max_depth=3):
    """
    Trả về danh sách đường dẫn folder (dạng chuỗi, không phải COM object) có trong
    mailbox — dùng để hiển thị lên UI cho người dùng chọn thêm folder cần quét
    (ngoài Inbox mặc định), vd: nếu có rule tự động chuyển mail sang folder khác.
    Đường dẫn con dùng dấu \\ để phân cách, vd: "Linh", "Inbox\\SubFolder".
    """
    pythoncom.CoInitialize()
    try:
        outlook = _outlook_app()
        ns = outlook.GetNamespace("MAPI")
        paths = []

        def walk(folder, path, depth):
            paths.append(path)
            if depth >= max_depth:
                return
            try:
                for sub in folder.Folders:
                    walk(sub, f"{path}\\{sub.Name}", depth + 1)
            except Exception:
                pass

        for store_folder in ns.Folders:
            try:
                for sub in store_folder.Folders:
                    walk(sub, sub.Name, 1)
            except Exception:
                continue
        return sorted(set(paths))
    finally:
        pythoncom.CoUninitialize()


def _find_folder(ns, path):
    """Tìm folder theo đường dẫn dạng 'Linh' hoặc 'Inbox\\SubFolder' (dò qua mọi mailbox/store)."""
    parts = [p for p in path.split("\\") if p]
    for store_folder in ns.Folders:
        try:
            current = store_folder
            matched = True
            for part in parts:
                found = None
                for sub in current.Folders:
                    if sub.Name == part:
                        found = sub
                        break
                if found is None:
                    matched = False
                    break
                current = found
            if matched:
                return current
        except Exception:
            continue
    return None


def _walk_all_folders(ns, max_depth=4):
    """Trả về list các COM folder object (không phải string path) — dùng nội bộ khi
    scan_all=True, để không phải resolve lại path -> folder object 1 lần nữa."""
    folders = []

    def walk(folder, depth):
        folders.append(folder)
        if depth >= max_depth:
            return
        try:
            for sub in folder.Folders:
                walk(sub, depth + 1)
        except Exception:
            pass

    for store_folder in ns.Folders:
        try:
            for sub in store_folder.Folders:
                walk(sub, 1)
        except Exception:
            continue
    return folders


def _build_scope(ns, folder_paths, scan_all):
    """Trả về chuỗi Scope cho Outlook AdvancedSearch (đường dẫn folder đầy đủ, có dấu nháy đơn)."""
    if scan_all:
        # Scope = toàn bộ mailbox (folder gốc), SearchSubFolders=True sẽ quét hết mọi folder con.
        parts = []
        for store_folder in ns.Folders:
            try:
                parts.append(f"'{store_folder.FolderPath}'")
            except Exception:
                continue
        if parts:
            return ",".join(parts)

    if folder_paths:
        parts = []
        for path in folder_paths:
            try:
                if path.strip().lower() == "inbox":
                    fobj = ns.GetDefaultFolder(6)
                else:
                    fobj = _find_folder(ns, path)
                if fobj is not None:
                    parts.append(f"'{fobj.FolderPath}'")
            except Exception:
                continue
        if parts:
            return ",".join(parts)

    inbox = ns.GetDefaultFolder(6)
    return f"'{inbox.FolderPath}'"


def _advanced_search_scan(outlook, ns, event_id, folder_paths, scan_all, timeout=25):
    """
    Quét NHANH bằng Windows Search Index của Outlook (AdvancedSearch) thay vì tự lặp
    qua từng email trong từng folder — nhanh hơn NHIỀU lần khi mailbox có nhiều folder,
    vì Outlook chỉ tra chỉ mục có sẵn thay vì mở từng email.

    ⚠️ Lưu ý quan trọng: AdvancedSearch chạy BẤT ĐỒNG BỘ — lúc mới gọi, Results.Count
    thường = 0 trong vài giây đầu (đang tra chỉ mục), rồi mới tăng dần. Vì vậy KHÔNG được
    coi "count không đổi" là "đã xong" khi count vẫn đang = 0 — nếu không sẽ thoát vòng lặp
    quá sớm và bỏ sót kết quả (đây là lỗi đã gặp ở bản trước). Logic dưới đây:
      - Luôn chờ tối thiểu `min_wait` giây trước khi bắt đầu xét "ổn định".
      - Chỉ coi là "ổn định/xong" khi count > 0 VÀ không đổi qua nhiều lần kiểm tra liên tiếp.
      - Nếu count vẫn = 0, tiếp tục chờ tới khi hết `timeout`.
    """
    event_id_escaped = event_id.replace("'", "''")
    scope = _build_scope(ns, folder_paths, scan_all)
    dasl_filter = f"\"urn:schemas:httpmail:subject\" LIKE '%{event_id_escaped}%'"

    search = outlook.AdvancedSearch(scope, dasl_filter, True, "RSVP_VoteSearch")

    min_wait = 3.0
    start = time.time()
    last_count = -1
    stable_checks = 0
    while time.time() - start < timeout:
        time.sleep(0.5)
        try:
            count = search.Results.Count
        except Exception:
            count = 0
        elapsed = time.time() - start
        if count > 0 and count == last_count and elapsed >= min_wait:
            stable_checks += 1
            if stable_checks >= 4:  # ~2 giây không đổi (và có ít nhất 1 kết quả) → coi là xong
                break
        else:
            stable_checks = 0
        last_count = count

    candidates = []
    try:
        for item in search.Results:
            try:
                if item.Class != 43:  # 43 = olMail
                    continue
                candidates.append(item)
            except Exception:
                continue
    except Exception:
        pass

    return _build_responses_from_candidates(candidates)


def _build_responses_from_candidates(candidates):
    candidates.sort(key=lambda it: it.ReceivedTime, reverse=True)
    responses = {}
    skipped_non_vote = 0
    for item in candidates:
        try:
            vote = getattr(item, "VotingResponse", "") or ""
            if not vote:
                skipped_non_vote += 1
                continue
            email = _smtp_address(item)
            if email in responses:
                continue
            responses[email] = {
                "name": item.SenderName or "",
                "vote": vote.strip(),
                "received": item.ReceivedTime,
            }
        except Exception:
            continue
    return responses, skipped_non_vote


def _manual_folder_scan(ns, event_id, folder_paths, scan_all):
    """Quét thủ công (lặp qua từng email trong từng folder) — chậm hơn AdvancedSearch
    nhưng không phụ thuộc vào chỉ mục Windows Search, nên luôn đáng tin cậy."""
    if scan_all:
        folders = _walk_all_folders(ns)
    elif not folder_paths:
        folders = [ns.GetDefaultFolder(6)]  # Inbox mặc định
    else:
        folders = []
        for path in folder_paths:
            if path.strip().lower() == "inbox":
                folders.append(ns.GetDefaultFolder(6))
                continue
            f = _find_folder(ns, path)
            if f is not None:
                folders.append(f)

    # Gom item từ TẤT CẢ folder trước, rồi sort chung theo thời gian mới nhất,
    # để dedup đúng ngay cả khi 1 người có mail rải rác ở nhiều folder khác nhau.
    candidates = []
    for folder in folders:
        try:
            for item in folder.Items:
                try:
                    if item.Class != 43:  # 43 = olMail
                        continue
                    subject = item.Subject or ""
                    if event_id not in subject:
                        continue
                    candidates.append(item)
                except Exception:
                    continue
        except Exception:
            continue

    return _build_responses_from_candidates(candidates)


def scan_voting_responses(event_id, folder_paths=None, scan_all=False):
    """
    Quét tìm các email PHẢN HỒI VOTE (có thuộc tính VotingResponse) có Subject chứa event_id.

    Cách quét CHÍNH: dùng Outlook AdvancedSearch (tra chỉ mục Windows Search có sẵn) —
    NHANH hơn nhiều so với việc tự lặp qua từng email trong từng folder, đặc biệt khi
    mailbox có nhiều folder.

    AN TOÀN 2 LỚP: nếu AdvancedSearch lỗi HOẶC trả về 0 kết quả (có thể do chỉ mục
    Windows Search chưa kịp cập nhật — hay gặp ngay sau khi vừa nhận được reply mới),
    tool sẽ TỰ ĐỘNG chạy thêm 1 lượt quét thủ công (chậm hơn nhưng không phụ thuộc chỉ
    mục) và gộp kết quả lại, để không bỏ sót phản hồi nào dù chậm hơn đôi chút.

    scan_all=True  → quét TOÀN BỘ folder trong mailbox — mặc định nên dùng cách này.
    folder_paths   → chỉ dùng khi scan_all=False: list[str] đường dẫn folder cụ thể
                     cần quét (vd: ["Inbox", "Linh"]). Nếu để None và scan_all=False
                     → chỉ quét Inbox mặc định (hành vi cũ, để tương thích ngược).

    Trả về dict: { email: {"name":.., "vote":"Yes"/"No"/"Maybe", "received": datetime} }
    Chỉ giữ lại phản hồi MỚI NHẤT của mỗi người, tính GỘP trên tất cả các folder đã quét.
    """
    import win32com.client  # noqa: F401  (chỉ để báo lỗi rõ ràng nếu thiếu pywin32)

    pythoncom.CoInitialize()
    try:
        outlook = _outlook_app()
        ns = outlook.GetNamespace("MAPI")

        responses, skipped = {}, 0
        try:
            responses, skipped = _advanced_search_scan(outlook, ns, event_id, folder_paths, scan_all)
        except Exception:
            pass  # AdvancedSearch lỗi → responses vẫn rỗng, sẽ chạy quét thủ công bên dưới

        if not responses:
            # AdvancedSearch lỗi hoặc không tìm thấy gì — tự động quét bổ sung bằng
            # cách thủ công để chắc chắn không bỏ sót, rồi gộp kết quả lại.
            manual_responses, manual_skipped = _manual_folder_scan(ns, event_id, folder_paths, scan_all)
            for email, data in manual_responses.items():
                if email not in responses or data["received"] > responses[email]["received"]:
                    responses[email] = data
            skipped = max(skipped, manual_skipped)

        return responses, skipped
    finally:
        pythoncom.CoUninitialize()


_MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _outlook_date_str(dt):
    """Format a datetime as an UNAMBIGUOUS 'DD Mon YYYY HH:MM' string (e.g.
    '10 Aug 2026 18:00') for assigning to appt.Start/appt.End.

    Why a STRING instead of a datetime/pywintypes.Time object: assigning any
    Python date OBJECT (raw datetime.datetime, or pywintypes.Time()) still
    goes through pywin32's automatic COM-DATE marshalling layer, which has
    been observed to still shift the time by the local UTC offset (e.g.
    18:00 configured → 9:00 shown in Outlook, exactly -9h = JST) even after
    switching to pywintypes.Time(timetuple()) — i.e. the marshalling itself
    is the unreliable part, not just which Python type triggers it.

    Assigning a STRING sidesteps that layer entirely: Outlook's COM property
    setter parses the string itself (same code path Outlook uses to parse
    what a person types into the Start/End box), so whatever it produces is
    guaranteed to match what typing the same text manually would produce —
    no separate datetime→COM-DATE conversion involved.

    Using an explicit 3-letter month name (not digits) avoids a SECOND,
    unrelated ambiguity: numeric "MM/DD" vs "DD/MM" order depends on the
    OS/Outlook locale (Japan generally expects YYYY/MM/DD), so a numeric
    string could be misread as the wrong day/month. A month name has no
    such ambiguity in any locale. The name is hardcoded here (not produced
    via strftime's locale-dependent "%b") so it's always the English
    abbreviation regardless of the machine's locale/language settings.
    """
    return f"{dt.day:02d} {_MONTH_ABBR[dt.month - 1]} {dt.year} {dt.hour:02d}:{dt.minute:02d}"


def send_calendar_invite(attendees, subject, location, start_dt, end_dt, body="",
                          attach_event_id=None, attach_hint_datetime=None, include_update=True):
    """
    Tạo Meeting Request (lời mời lịch Outlook) gửi tới danh sách attendees.
    attendees: list[email] — thường là những người đã vote "Yes"
    start_dt / end_dt: đối tượng datetime.datetime
    attach_event_id: (tuỳ chọn) EventID — nếu có, hàm sẽ tự tìm email Invite
        GỐC ('[Confirm-...]'/'【出欠確認-...】'/'[XacNhan-...]') và, nếu
        include_update=True, email UPDATE INVITE ('[UPDATED-...]'/
        '【変更のお知らせ-...】'/'[CapNhat-...]', nếu đã từng gửi qua chế độ
        'Send update invite' ở Tab 3) đã gửi trước đó trong Sent Items —
        CHỈ 2 loại này (không lấy email nhắc nhở, không lấy email reply/vote
        dù Subject của chúng có chứa event_id — xem docstring
        _find_all_sent_invite_mails()) — và đính kèm dạng file .msg vào
        Appointment, để người nhận có đủ lịch sử các lần mời/cập nhật
        thông tin sự kiện.
    include_update: đặt False nếu RSVP_History.xlsx cho biết sự kiện này
        chưa từng gửi update invite (cột UpdateInviteDate trống) — bỏ qua
        tìm loại này luôn, không chỉ lọc kết quả.
    attach_hint_datetime: không còn dùng để LỌC (trước đây dùng để chọn 1
        email khớp nhất khi có nhiều email trùng event_id) — giờ hàm đính
        kèm TẤT CẢ nên tham số này chỉ giữ lại cho tương thích ngược, không
        ảnh hưởng kết quả.
    Mở cửa sổ để review trước khi Send (an toàn hơn gửi thẳng).

    Trả về: (appt, attached_count) — `attached_count` là SỐ email đã đính
        kèm THÀNH CÔNG (0 nếu không tìm thấy hoặc lỗi) — xem chú thích BUG
        ĐÍNH KÈM bên dưới.

    ⚠️ BUG ĐÃ GẶP (giờ họp bị lệch, vd: cấu hình 18:00-21:00 nhưng Outlook
    hiển thị 9:00-12:00 — lệch đúng 9 tiếng = offset giờ Nhật JST UTC+9):
    lần đầu nghi do gán trực tiếp Python `datetime.datetime` khiến pywin32
    hiểu nhầm UTC → đã thử sửa bằng `pywintypes.Time(tuple)`, NHƯNG vẫn bị
    lệch y hệt — nghĩa là bản thân tầng marshalling datetime→COM DATE của
    pywin32 không đáng tin cậy cho trường hợp này, bất kể dùng kiểu Python
    nào. Cách sửa CHẮC CHẮN hơn: gán CHUỖI TEXT (không phải object ngày giờ)
    cho `.Start`/`.End` — để Outlook tự parse bằng đúng engine nó dùng khi
    người dùng gõ tay vào ô Start/End, bỏ qua hoàn toàn tầng pywin32 hay bị
    lỗi. Dùng tên tháng viết tắt (vd "10 Aug 2026 18:00") thay vì số để
    tránh thêm 1 lỗi khác: nhầm lẫn thứ tự ngày/tháng dạng số tuỳ theo locale
    máy (Nhật thường dùng yyyy/mm/dd, có thể đọc nhầm 10/08 thành tháng 10).

    ⚠️ BUG ĐÃ GẶP (báo "đã tìm thấy và đính kèm" nhưng thực tế Appointment
    KHÔNG có file đính kèm nào): bản trước tìm email bằng 1 hàm riêng
    (find_sent_mail_by_subject) — hàm đó tự mở/đóng phiên COM RIÊNG của nó
    (CoInitialize...CoUninitialize) TRƯỚC KHI trả MailItem về, rồi hàm NÀY
    lại mở 1 phiên COM KHÁC để tạo Appointment và gọi Attachments.Add() với
    MailItem đó. Việc dùng 1 COM object đã lấy từ phiên COM ĐÃ ĐÓNG (CoUninit
    xong) trong 1 phiên COM MỚI khác khiến reference bị hỏng/không hợp lệ —
    Attachments.Add() âm thầm ném lỗi (bị nuốt bởi try/except), nhưng code cũ
    vẫn báo "thành công" vì chỉ kiểm tra biến Python có None hay không, không
    kiểm tra việc Add() có thực sự chạy được hay không. Cách sửa: gộp việc
    TÌM + ĐÍNH KÈM vào CÙNG 1 phiên COM duy nhất (bên trong hàm này), và trả
    về SỐ email thực sự đính kèm thành công (không phải chỉ cờ True/False).
    """
    pythoncom.CoInitialize()
    try:
        outlook = _outlook_app()
        appt = outlook.CreateItem(1)  # 1 = olAppointmentItem
        appt.MeetingStatus = 1        # 1 = olMeeting
        appt.Subject = subject
        appt.Location = location
        appt.AllDayEvent = False
        # Assign as an unambiguous STRING — see _outlook_date_str() docstring
        # above for why this is more reliable than any datetime/pywintypes
        # object here.
        appt.Start = _outlook_date_str(start_dt)
        appt.End = _outlook_date_str(end_dt)
        appt.Body = body
        appt.ReminderSet = True
        appt.ReminderMinutesBeforeStart = 30
        for email in attendees:
            if not email:
                continue
            r = appt.Recipients.Add(email)
            r.Type = 1  # 1 = olRequired
        try:
            appt.Recipients.ResolveAll()
        except Exception:
            pass

        attached_count = 0
        if attach_event_id:
            try:
                ns = outlook.GetNamespace("MAPI")
                found_mails = _find_all_sent_invite_mails(ns, attach_event_id, include_update=include_update)
                for m in found_mails:
                    try:
                        # Adding an existing Outlook item directly (not a file
                        # path) embeds it as a .msg attachment on the
                        # Appointment. These MailItems were obtained from the
                        # SAME `outlook` Application object, within this SAME
                        # CoInitialize session — unlike the old bug, so the
                        # reference stays valid here.
                        appt.Attachments.Add(m)
                        attached_count += 1
                    except Exception:
                        continue  # 1 email lỗi không nên chặn đính kèm các email còn lại
            except Exception:
                attached_count = 0  # best-effort — don't fail the whole send over this

        appt.Display()  # review trước khi bấm Send trong Outlook
        return appt, attached_count
    finally:
        pythoncom.CoUninitialize()


def _find_all_sent_invite_mails(ns, event_id, include_update=True):
    """Tìm email đã gửi (Sent Items) liên quan tới event_id — nhưng CHỈ 2
    loại: email mời GỐC và email UPDATE INVITE. KHÔNG lấy email nhắc nhở,
    và KHÔNG lấy các email reply/vote (vd 'Yes: [Confirm-...]',
    'No: [...]', 'RE: [...]') dù Subject của chúng cũng chứa event_id.

    ⚠️ BUG ĐÃ SỬA: bản trước lọc bằng "event_id CÓ TRONG Subject" (contains),
    nên vô tình khớp luôn CẢ email reply tự động khi ai đó (kể cả chính
    người tổ chức, nếu email của họ cũng nằm trong danh sách mời) bấm Vote
    — Outlook tự tạo 1 email reply dạng "Yes: [Confirm-...]"/"No: [...]"/
    "RE: [...]" và lưu vào Sent Items, Subject vẫn giữ nguyên phần gốc nên
    vẫn chứa event_id — LẪN email nhắc nhở "[Reminder-{event_id}]"/
    "【リマインド-{event_id}】". Kết quả: Calendar Invite bị đính kèm dư thừa
    nhiều email không liên quan (xem ảnh chụp người dùng gửi). Giờ đổi
    sang so khớp "Subject BẮT ĐẦU BẰNG đúng 1 trong các tiền tố Invite/
    Update-invite THẬT" — đây CHÍNH XÁC là các Subject mà build_subject()
    (rsvp_app.py) tạo ra cho email Invite gốc/Update invite (ở CẢ 4 lựa
    chọn ngôn ngữ EN/JA/VI/Bilingual — bilingual luôn bắt đầu bằng tiền tố
    EN nên vẫn khớp) — loại hẳn được email reply (có thêm "Yes: "/"No: "/
    "RE: " phía trước, không khớp tiền tố) và email nhắc nhở (tiền tố khác
    hẳn "[Reminder-"/"【リマインド-"), tức là CHỈ đính kèm đúng 2 loại email
    mà cột SentDate/UpdateInviteDate trong RSVP_History.xlsx đang theo dõi.

    include_update: nếu False, không tìm email UPDATE INVITE (dùng khi
        RSVP_History.xlsx cho biết sự kiện này CHƯA từng gửi update invite
        — cột UpdateInviteDate trống — để không mất công quét thêm).

    Trả về list[MailItem], sắp xếp theo thời gian gửi (CŨ NHẤT trước)."""
    invite_prefixes = (
        f"[Confirm-{event_id}]",
        f"【出欠確認-{event_id}】",
        f"[XacNhan-{event_id}]",
    )
    update_prefixes = (
        f"[UPDATED-{event_id}]",
        f"【変更のお知らせ-{event_id}】",
        f"[CapNhat-{event_id}]",
    )
    sent_folder = ns.GetDefaultFolder(5)  # 5 = olFolderSentMail
    candidates = []
    for item in sent_folder.Items:
        try:
            if item.Class != 43:  # 43 = olMail
                continue
            subject = item.Subject or ""
            if subject.startswith(invite_prefixes):
                candidates.append(item)
            elif include_update and subject.startswith(update_prefixes):
                candidates.append(item)
        except Exception:
            continue
    candidates.sort(key=lambda it: it.SentOn)
    return candidates


def _find_sent_invite_mail(ns, event_id, hint_datetime=None):
    """Tìm email GỐC đã gửi trước đó (trong Sent Items, Subject chứa
    event_id) — dùng chung cho cả send_calendar_invite() (đính kèm vào
    Appointment) và send_reminder_email() (đính kèm vào email nhắc nhở).

    ⚠️ PHẢI được gọi trong CÙNG 1 phiên COM (pythoncom.CoInitialize) với nơi
    sẽ dùng MailItem trả về (vd Attachments.Add) — nếu gọi ở 1 phiên COM
    riêng rồi trả object qua phiên khác, reference sẽ bị hỏng (xem chú thích
    BUG ĐÍNH KÈM ở send_calendar_invite để biết chi tiết lỗi đã gặp).

    hint_datetime: (tuỳ chọn) datetime lấy từ cột SentDate trong
        RSVP_History.xlsx — dùng để chọn đúng email khi có NHIỀU email cùng
        chứa event_id trong Subject (vd: gửi lại/gửi nhắc nhiều lần). Không có
        hint → lấy email MỚI NHẤT khớp Subject (thường là lần gửi gần nhất).
    """
    sent_folder = ns.GetDefaultFolder(5)  # 5 = olFolderSentMail
    candidates = []
    for item in sent_folder.Items:
        try:
            if item.Class != 43:  # 43 = olMail
                continue
            if event_id not in (item.Subject or ""):
                continue
            candidates.append(item)
        except Exception:
            continue

    if not candidates:
        return None
    if hint_datetime is None or len(candidates) == 1:
        candidates.sort(key=lambda it: it.SentOn, reverse=True)
        return candidates[0]

    def _time_diff(it):
        try:
            naive_sent_on = it.SentOn.replace(tzinfo=None)
            return abs((naive_sent_on - hint_datetime).total_seconds())
        except Exception:
            return float("inf")

    candidates.sort(key=_time_diff)
    return candidates[0]


def send_reminder_email(pending, subject, body, voting_options="Yes;No;Maybe",
                         auto_send=False, attach_event_id=None, attach_hint_datetime=None):
    """
    Gửi email NHẮC NHỞ tới những người CHƯA trả lời (pending), vẫn giữ
    Voting Buttons (Yes/No/Maybe) như mail gốc để họ bấm vote ngay trên email
    nhắc — Subject truyền vào PHẢI vẫn chứa đúng event_id như mail gốc, để
    lần "Scan Inbox for Vote results" sau đó vẫn quét/khớp được các phiếu
    vote mới trả lời trên email nhắc này (cơ chế quét dựa vào event_id nằm
    trong Subject, không quan tâm mail nào — mail mời gốc hay mail nhắc).

    pending: list[(name, email)] — danh sách người CHƯA trả lời (lấy từ Tab 4).
    attach_event_id: nếu có, tự tìm + đính kèm email MỜI GỐC đã gửi trước đó
        (dùng chung logic với send_calendar_invite, xem _find_sent_invite_mail).
    attach_hint_datetime: (tuỳ chọn) SentDate của email mời gốc, lấy từ
        RSVP_History.xlsx — giúp chọn đúng email khi Sent Items có nhiều email
        cùng event_id (vd gửi nhắc nhiều lần, mỗi lần đều chứa event_id).

    Trả về: (mail, attached) — attached=True chỉ khi Attachments.Add() thực
        sự chạy không lỗi (không phải chỉ "có tìm thấy candidate").
    """
    pythoncom.CoInitialize()
    try:
        outlook = _outlook_app()
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = ";".join(e for _, e in pending if e)
        mail.Subject = subject
        mail.Body = body
        mail.VotingOptions = voting_options
        _attach_vote_illustration(mail)  # ảnh minh hoạ cách bấm Vote (best-effort)

        attached = False
        if attach_event_id:
            try:
                ns = outlook.GetNamespace("MAPI")
                found_mail = _find_sent_invite_mail(ns, attach_event_id, attach_hint_datetime)
                if found_mail is not None:
                    mail.Attachments.Add(found_mail)
                    attached = True
            except Exception:
                attached = False  # best-effort — không làm hỏng cả việc gửi vì lỗi đính kèm

        if auto_send:
            mail.Send()
        else:
            mail.Display()
        return mail, attached
    finally:
        pythoncom.CoUninitialize()


def _find_sent_gift_mail(ns, event_id):
    """Tìm email THÔNG BÁO QUYÊN GÓP QUÀ TẶNG gốc đã gửi trước đó (Sent
    Items) — dùng riêng cho Gift Contribution reminder, KHÔNG dùng chung với
    _find_sent_invite_mail()/_find_all_sent_invite_mails() (những hàm đó chỉ
    khớp tiền tố Invite/Update-invite '[Confirm-'/'[UPDATED-'/... — email
    Gift dùng tiền tố RIÊNG '[Gift-'/'【寄付のお願い-'/'[QuyenGop-', xem
    build_gift_subject() trong rsvp_app.py).

    Subject BẮT ĐẦU BẰNG đúng 1 trong các tiền tố Gift THẬT (cả 3 ngôn ngữ —
    bilingual luôn bắt đầu bằng tiền tố "en" nên vẫn khớp) — tránh khớp nhầm
    email reply hoặc chính email reminder Gift trước đó (nếu đã từng gửi).

    Trả về MailItem MỚI NHẤT khớp Subject, hoặc None nếu không tìm thấy.
    """
    gift_prefixes = (
        f"[Gift-{event_id}]",
        f"【寄付のお願い-{event_id}】",
        f"[QuyenGop-{event_id}]",
    )
    sent_folder = ns.GetDefaultFolder(5)  # 5 = olFolderSentMail
    candidates = []
    for item in sent_folder.Items:
        try:
            if item.Class != 43:  # 43 = olMail
                continue
            subject = item.Subject or ""
            if subject.startswith(gift_prefixes):
                candidates.append(item)
        except Exception:
            continue
    if not candidates:
        return None
    candidates.sort(key=lambda it: it.SentOn, reverse=True)
    return candidates[0]


def send_gift_reminder_email(pending, subject, body, auto_send=False, attach_event_id=None):
    """Gửi email NHẮC NHỞ QUYÊN GÓP QUÀ TẶNG tới những người CHƯA đóng góp
    (pending) — KHÔNG có Voting Buttons và KHÔNG đính kèm ảnh hướng dẫn vote
    (giống send_voting_invite(..., use_voting_buttons=False) dùng cho Gift
    mode nói chung — đây thuần là 1 email thông báo/nhắc nhở).

    pending: list[(name, email)] — danh sách người CHƯA tick "Đã góp" (lấy
        từ file Gift_Contribution_List_{EventID}.xlsx, cột "Contributed" != "Yes").
    attach_event_id: nếu có, tự tìm + đính kèm email THÔNG BÁO QUYÊN GÓP GỐC
        đã gửi trước đó (xem _find_sent_gift_mail ở trên).

    Trả về: (mail, attached) — attached=True chỉ khi Attachments.Add() thực
        sự chạy không lỗi (không phải chỉ "có tìm thấy candidate").
    """
    pythoncom.CoInitialize()
    try:
        outlook = _outlook_app()
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = ";".join(e for _, e in pending if e)
        mail.Subject = subject
        mail.Body = body

        attached = False
        if attach_event_id:
            try:
                ns = outlook.GetNamespace("MAPI")
                found_mail = _find_sent_gift_mail(ns, attach_event_id)
                if found_mail is not None:
                    mail.Attachments.Add(found_mail)
                    attached = True
            except Exception:
                attached = False  # best-effort — không làm hỏng cả việc gửi vì lỗi đính kèm

        if auto_send:
            mail.Send()
        else:
            mail.Display()
        return mail, attached
    finally:
        pythoncom.CoUninitialize()


def send_gift_report_email(recipients, subject, body, excel_path=None, auto_send=False):
    """MỚI: gửi email BÁO CÁO số tiền đã quyên góp được (Tab 6 'Gift
    Contribution') tới những người được TICK CHỌN ở cột "Send email" —
    KHÔNG có Voting Buttons và KHÔNG đính ảnh hướng dẫn vote (thuần là 1
    thông báo, giống send_gift_reminder_email/send_voting_invite(...,
    use_voting_buttons=False)) — nhưng KHÔNG dùng chung với
    send_voting_invite() vì hàm đó không hỗ trợ đính kèm file.

    excel_path: đường dẫn file Excel báo cáo (danh sách người ĐÃ đóng góp,
        đánh số lại, không gồm 2 cột checkbox) do rsvp_app.py tự lưu ra 1
        file tạm TRƯỚC khi gọi hàm này (xem
        RSVPApp._build_gift_report_workbook()) — LUÔN được đính kèm nếu
        file tồn tại; đây là ĐƯỜNG DẪN FILE (không phải Outlook item, khác
        với cách đính kèm ở send_calendar_invite()/send_reminder_email()).

    Trả về: (mail, attached) — attached=True chỉ khi Attachments.Add() cho
        file Excel thực sự chạy không lỗi."""
    pythoncom.CoInitialize()
    try:
        outlook = _outlook_app()
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = ";".join(e for _, e in recipients if e)
        mail.Subject = subject
        mail.Body = body

        attached = False
        if excel_path and os.path.exists(excel_path):
            try:
                mail.Attachments.Add(excel_path)
                attached = True
            except Exception:
                attached = False  # best-effort — không đính kèm được thì thôi, không chặn việc gửi

        if auto_send:
            mail.Send()
        else:
            mail.Display()
        return mail, attached
    finally:
        pythoncom.CoUninitialize()


def _find_calendar_invite_appointment(ns, event_name):
    """MỚI (Tab 5 'Attendance & Payment' — email cảm ơn sau sự kiện): tìm
    Calendar Invite (Appointment) của sự kiện, để đính lại vào email cảm ơn
    gửi sau khi sự kiện kết thúc.

    Một Meeting Request tạo qua send_calendar_invite() luôn để lại 1 bản
    sao trong CHÍNH folder Calendar của người tổ chức — kể cả khi mới chỉ
    Display() để review, chưa bấm Send — với Subject = đúng `subject`
    truyền vào send_calendar_invite() (chính là Event Name, xem
    _send_calendar() trong rsvp_app.py). Vì vậy quét folder Calendar
    (KHÔNG PHẢI Sent Items — Meeting Request không phải MailItem) tìm đúng
    Subject khớp là cách xác định lại nó, không cần lưu ID gì thêm.

    So khớp Subject SAU KHI .strip().lower() (bỏ khoảng trắng thừa/không
    phân biệt hoa thường) — Event Name gõ tay có thể lệch 1 khoảng trắng
    so với lúc gửi Calendar Invite ban đầu.

    Trả về AppointmentItem MỚI NHẤT (theo giờ bắt đầu) khớp Subject, hoặc
    None nếu không tìm thấy — CHỈ dùng NGAY TRONG CÙNG 1 phiên COM với nơi
    gọi (xem chú thích BUG ĐÍNH KÈM ở send_calendar_invite() để biết lý do
    — reference sẽ hỏng nếu tìm ở 1 phiên COM khác rồi đem dùng ở đây)."""
    target = (event_name or "").strip().lower()
    if not target:
        return None
    cal_folder = ns.GetDefaultFolder(9)  # 9 = olFolderCalendar
    candidates = []
    for item in cal_folder.Items:
        try:
            if item.Class != 26:  # 26 = olAppointment
                continue
            if (item.Subject or "").strip().lower() == target:
                candidates.append(item)
        except Exception:
            continue
    if not candidates:
        return None
    try:
        candidates.sort(key=lambda it: it.Start, reverse=True)
    except Exception:
        pass  # sort lỗi (vd .Start không đọc được) -> vẫn trả về candidate đầu tiên tìm thấy
    return candidates[0]


def send_thankyou_email(recipients, subject, body, excel_path=None, event_name=None, auto_send=False):
    """MỚI: gửi email CẢM ƠN sau sự kiện (Tab 5 'Attendance & Payment')
    tới những người "Actual Attend" = Yes. KHÔNG có Voting Buttons và
    KHÔNG đính ảnh hướng dẫn vote (giống send_gift_reminder_email — thuần
    là email thông báo, không phải 1 cuộc bỏ phiếu).

    Tự đính kèm 2 thứ, trong CÙNG 1 phiên COM (đúng nguyên tắc đã rút ra từ
    bug đính kèm ở send_calendar_invite — xem docstring hàm đó):
      1. File Excel báo cáo Attendance & Payment — `excel_path` là ĐƯỜNG
         DẪN FILE (không phải Outlook item), do rsvp_app.py tự lưu ra file
         tạm trước khi gọi hàm này (Attachments.Add() nhận cả 2 kiểu: 1
         đường dẫn file HOẶC 1 Outlook item).
      2. Calendar Invite của sự kiện — tìm theo `event_name` qua
         _find_calendar_invite_appointment() ở trên.
    Cả 2 đều best-effort: thiếu 1 cái không chặn việc tạo/gửi email.

    recipients: list[(name, email)] — những người "Actual Attend" = Yes.
    Trả về: (mail, calendar_attached) — calendar_attached=True CHỈ khi
        Attachments.Add() cho Calendar Invite thực sự chạy không lỗi (cùng
        quy ước với `attached` ở send_reminder_email/send_gift_reminder_email).
    """
    pythoncom.CoInitialize()
    try:
        outlook = _outlook_app()
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = ";".join(e for _, e in recipients if e)
        mail.Subject = subject
        mail.Body = body

        if excel_path and os.path.exists(excel_path):
            try:
                mail.Attachments.Add(excel_path)
            except Exception:
                pass  # best-effort — báo cáo Excel không đính kèm được thì thôi, không chặn việc gửi

        calendar_attached = False
        try:
            ns = outlook.GetNamespace("MAPI")
            found_appt = _find_calendar_invite_appointment(ns, event_name)
            if found_appt is not None:
                mail.Attachments.Add(found_appt)
                calendar_attached = True
        except Exception:
            calendar_attached = False  # best-effort — không tìm/đính được Calendar Invite thì thôi

        if auto_send:
            mail.Send()
        else:
            mail.Display()
        return mail, calendar_attached
    finally:
        pythoncom.CoUninitialize()


def _expand_dl_addr_entry(addr_entry, seen_keys, depth):
    """addr_entry: 1 AddressEntry COM object ĐÃ BIẾT CHẮC là 1 Exchange
    Distribution List (group email). Trả về list[(name, smtp_email)] của
    TẤT CẢ thành viên THẬT (không phải group) — đệ quy fan-out qua mọi
    sub-group lồng bên trong, loại trùng theo email (seen_keys dùng chung
    xuyên suốt đệ quy để 1 người nằm ở nhiều sub-group không bị lặp)."""
    if depth <= 0:
        return []
    try:
        dl = addr_entry.GetExchangeDistributionList()
        if dl is None:
            return []
        member_entries = dl.GetExchangeDistributionListMembers()
        count = member_entries.Count
    except Exception:
        return []

    result = []
    for i in range(1, count + 1):
        try:
            m = member_entries.Item(i)
        except Exception:
            continue
        sub_dl = None
        try:
            sub_dl = m.GetExchangeDistributionList()
        except Exception:
            pass
        if sub_dl is not None:
            # Thành viên này bản thân là 1 sub-group lồng bên trong -> đệ quy
            # fan-out tiếp, KHÔNG thêm chính cái sub-group vào kết quả.
            result.extend(_expand_dl_addr_entry(m, seen_keys, depth - 1))
            continue
        try:
            exch_user = m.GetExchangeUser()
            smtp = exch_user.PrimarySmtpAddress.lower() if exch_user else None
            name = (exch_user.Name if exch_user else None) or m.Name
        except Exception:
            smtp, name = None, getattr(m, "Name", None)
        if not smtp or "@" not in smtp or smtp in seen_keys:
            continue
        seen_keys.add(smtp)
        result.append((name or smtp, smtp))
    return result


def expand_group_members(email_or_name, max_depth=6):
    """
    Nếu `email_or_name` là 1 Exchange Distribution List (group email trong sổ
    địa chỉ công ty — GAL) — có thể chứa sub-group lồng bên trong — trả về
    list[(name, smtp_email)] của TẤT CẢ thành viên THẬT, đã fan-out đệ quy
    qua mọi sub-group, loại trùng theo email.

    Trả về None nếu `email_or_name` KHÔNG PHẢI group (là 1 người dùng bình
    thường, hoặc không resolve được qua Outlook) — caller nên giữ nguyên
    dòng gốc trong trường hợp này (xem _expand_group_recipients() ở
    rsvp_app.py: dòng nào expand ra None thì được giữ y nguyên trong danh
    sách, không bị xoá).

    ⚠️ YÊU CẦU: tài khoản Outlook đang dùng phải nằm trên Exchange/Microsoft
    365 và group đó có trong GAL của tổ chức. Với mailing-list/group ngoài
    Exchange (Google Group, group tự tạo trong Contacts cá nhân...), Outlook
    KHÔNG có cách nào liệt kê thành viên qua COM — hàm sẽ trả về None, và
    dòng group email đó vẫn được giữ nguyên như 1 người nhận bình thường
    (tool không biết đó là group nên không thể tách số người chưa trả lời
    chính xác cho trường hợp này).
    """
    pythoncom.CoInitialize()
    try:
        outlook = _outlook_app()
        ns = outlook.GetNamespace("MAPI")
        recip = ns.CreateRecipient(email_or_name)
        recip.Resolve()
        if not recip.Resolved:
            return None
        addr_entry = recip.AddressEntry
        if addr_entry is None:
            return None
        try:
            dl = addr_entry.GetExchangeDistributionList()
        except Exception:
            dl = None
        if dl is None:
            return None  # không phải group -> caller tự giữ nguyên dòng gốc

        seen = set()
        return _expand_dl_addr_entry(addr_entry, seen, max_depth)
    finally:
        pythoncom.CoUninitialize()

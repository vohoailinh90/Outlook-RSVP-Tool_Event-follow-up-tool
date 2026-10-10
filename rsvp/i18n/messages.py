"""Subject and body builders for every outgoing message.

Every language table here is read as TABLE.get(lang_code, TABLE["en"]), so a
missing language silently serves English rather than failing.
scripts/check_i18n_matrix.py exists to catch that.

Moved verbatim out of rsvp_app.py (lines 274-895) by the phase 1 extraction in
docs/agentic/ARCHITECTURE.md. The code is unchanged; only its location is.

Exception: build_thankyou_body and build_gift_report_body (and their label
tables) were later replaced by the versions in the separately developed copy
of this app that the user runs - a table of payment rounds, and the gift
item with Event + Gift totals. tests/test_i18n_parity.py lists them as the
only builders allowed to differ from the phase-1 golden file, and
tests/golden/i18n_merge_snapshot.json pins their new output.
"""

from rsvp.domain.money import parse_signed_amount

from .langs import BILINGUAL_SEPARATOR, GREETING, NOT_TRANSLATED_FLAG
from .tables import _build_aligned_table, _build_rounds_table

def build_greeting(lang_code):
    if lang_code == "bilingual":
        return GREETING["en"] + BILINGUAL_SEPARATOR + GREETING["ja"]
    return GREETING.get(lang_code, GREETING["en"])


def build_subject(lang_code, event_id, event_name, is_update=False):
    if is_update:
        subs = {
            "en": f"[UPDATED-{event_id}] {event_name} - Event Details Have Changed",
            "ja": f"【変更のお知らせ-{event_id}】{event_name} - 開催情報が変更されました",
            "vi": f"[CapNhat-{event_id}] {event_name} - Thông tin sự kiện đã thay đổi",
        }
    else:
        subs = {
            "en": f"[Confirm-{event_id}] {event_name} - Please Confirm Attendance",
            "ja": f"【出欠確認-{event_id}】{event_name} - ご出欠確認のお願い",
            "vi": f"[XacNhan-{event_id}] {event_name} - Vui lòng xác nhận tham gia",
        }
    if lang_code == "bilingual":
        return subs["en"] + " / " + subs["ja"]
    return subs.get(lang_code, subs["en"])


UPDATE_NOTICE = {
    "en": ("⚠️ IMPORTANT: The details of this event have been UPDATED since the original "
           "invite. Please review the new information below (date/location/deadline/budget "
           "may have changed) and re-confirm your attendance if anything changed for you.\n"),
    "ja": ("⚠️ 重要: このイベントの詳細が、当初のご案内から変更されました（日時・場所・回答期限・"
           "費用などが変わっている可能性があります）。以下の新しい情報をご確認のうえ、必要に応じて"
           "再度ご回答をお願いいたします。\n"),
    "vi": ("⚠️ QUAN TRỌNG: Thông tin của sự kiện này đã được CẬP NHẬT so với email mời ban đầu "
           "(ngày giờ/địa điểm/hạn phản hồi/chi phí có thể đã thay đổi). Vui lòng xem lại thông "
           "tin mới bên dưới và xác nhận lại nếu có gì thay đổi với bạn.\n"),
}


def build_update_notice(lang_code):
    """Banner cảnh báo chèn ĐẦU email khi gửi ở chế độ 'Send update invite' —
    báo cho người nhận biết thông tin sự kiện đã thay đổi so với lần mời gốc."""
    if lang_code == "bilingual":
        return build_update_notice("en") + BILINGUAL_SEPARATOR + build_update_notice("ja")
    return UPDATE_NOTICE.get(lang_code, UPDATE_NOTICE["en"])


def build_fixed_block(lang_code, event_name, event_date, location, deadline, budget):
    """FIXED part: event details (auto-filled from Tab 1) + voting instructions
    (NO greeting — the greeting is shown separately, before the EDITABLE note).
    This is the SYSTEM DEFAULT wording — the person can hand-edit it in Tab 3 and save
    their own version as the new default (see fixed_overrides in RSVPApp)."""
    if lang_code == "bilingual":
        return (build_fixed_block("en", event_name, event_date, location, deadline, budget)
                + BILINGUAL_SEPARATOR
                + build_fixed_block("ja", event_name, event_date, location, deadline, budget))

    if lang_code == "ja":
        loc = f"（{location}）" if location else ""
        bud = f"予定費用: {budget}\n" if budget else ""
        return (
            f'「{event_name}」を {event_date} に開催いたします{loc}。\n\n'
            f'ご返信期限: {deadline}\n'
            f'{bud}\n'
            '>>> このメール上部にある「Yes / No / Maybe」の投票ボタンのいずれかをクリックしてください <<<\n'
            '（投票ボタンは通常メール本文の上にあるインフォバー、または Outlook デスクトップのリボンの'
            '「返信」欄に表示されます。Outlook Web でも同様にご利用いただけます）\n'
            '📎 ボタンの場所が分からない場合は、添付の画像（クリック方法の説明）をご確認ください。\n\n'
            '  • Yes   = 参加を確認します\n'
            '  • No    = 参加できません\n'
            '  • Maybe = まだ未定です、後日改めてご連絡します\n\n'
            '文字を入力したり件名を変更したりする必要はありません。ボタンをクリックするだけで自動的に記録されます。'
        )

    if lang_code == "vi":
        loc = f", tại {location}" if location else ""
        bud = f"Chi phí dự kiến: {budget}\n" if budget else ""
        return (
            f'Chúng ta tổ chức "{event_name}" vào ngày {event_date}{loc}.\n\n'
            f'Hạn phản hồi: {deadline}\n'
            f'{bud}\n'
            '>>> VUI LÒNG BẤM VÀO 1 TRONG 3 NÚT "Yes / No / Maybe" Ở ĐẦU EMAIL NÀY <<<\n'
            '(Nút vote thường hiện ở thanh InfoBar phía trên nội dung email, hoặc trong Ribbon '
            '"Respond" nếu bạn dùng Outlook desktop — Outlook Web cũng hỗ trợ tương tự)\n'
            '📎 Nếu không tìm thấy nút bấm, xem hình ảnh đính kèm bên dưới để biết vị trí cụ thể.\n\n'
            '  • Yes   = Bạn XÁC NHẬN tham gia\n'
            '  • No    = Bạn KHÔNG tham gia được\n'
            '  • Maybe = Bạn CHƯA CHẮC, sẽ xác nhận sau\n\n'
            'Bạn KHÔNG cần gõ chữ hay đổi tiêu đề — chỉ cần bấm 1 nút, hệ thống sẽ tự ghi nhận.'
        )

    # default / "en"
    loc = f", at {location}" if location else ""
    bud = f"Estimated cost: {budget}\n" if budget else ""
    return (
        f'Our team is organizing "{event_name}" on {event_date}{loc}.\n\n'
        f'Response deadline: {deadline}\n'
        f'{bud}\n'
        '>>> PLEASE CLICK ONE OF THE 3 BUTTONS "Yes / No / Maybe" AT THE TOP OF THIS EMAIL <<<\n'
        '(The voting buttons usually appear in the InfoBar above the message, or under the '
        '"Respond" section of the Ribbon in Outlook desktop — Outlook Web supports this too)\n'
        "📎 Not sure where the button is? See the attached image below for a visual guide.\n\n"
        '  • Yes   = You CONFIRM attendance\n'
        '  • No    = You CANNOT attend\n'
        '  • Maybe = NOT SURE yet, will confirm later\n\n'
        'You do not need to type anything or change the subject line — just click a button '
        'and the system will record it automatically.'
    )


REMINDER_LABELS = {
    "en": {"event": "📋 Event", "date": "⏰ Date", "location": "📍 Location",
           "budget": "💰 Budget", "deadline": "⏰ Please respond by"},
    "ja": {"event": "📋 イベント", "date": "⏰ 日時", "location": "📍 場所",
           "budget": "💰 予算", "deadline": "⏰ 回答期限"},
    "vi": {"event": "📋 Sự kiện", "date": "⏰ Ngày", "location": "📍 Địa điểm",
           "budget": "💰 Ngân sách", "deadline": "⏰ Vui lòng phản hồi trước"},
}


def build_reminder_body(lang_code, event_name, event_date, location, deadline, budget):
    """Nội dung email NHẮC NHỞ (Tab 4 'Collect Responses') — tương tự
    build_fixed_block() nhưng dùng giọng văn ngắn gọn, thân thiện, phù hợp
    cho email nhắc lại (không phải mời lần đầu). Hỗ trợ 4 lựa chọn giống hệt
    Tab 3: English / Japanese / Vietnamese / Bilingual (Japanese + English)
    — bilingual ghép theo đúng convention đã dùng ở nơi khác trong app
    (banner "[English below]" + nội dung tiếng Nhật trước, tiếng Anh sau,
    ngăn cách bằng BILINGUAL_SEPARATOR — xem _refresh_compose_preview())."""
    if lang_code == "bilingual":
        ja_full = build_reminder_body("ja", event_name, event_date, location, deadline, budget)
        en_full = build_reminder_body("en", event_name, event_date, location, deadline, budget)
        return "[English below]\n\n" + ja_full + BILINGUAL_SEPARATOR + en_full

    L = REMINDER_LABELS.get(lang_code, REMINDER_LABELS["en"])
    budget_line = f"{L['budget']}: {budget}\n" if budget else ""

    if lang_code == "ja":
        return (
            "皆様\n\n"
            "こちらはリマインドメールです。まだご回答をいただいておりません。\n\n"
            f"{L['event']}: {event_name}\n"
            f"{L['date']}: {event_date}\n"
            f"{L['location']}: {location}\n"
            f"{budget_line}"
            f"{L['deadline']}: {deadline}\n\n"
            ">>> このメール上部にある「Yes / No / Maybe」の投票ボタンのいずれかをクリックして"
            "ください <<<\n"
            "（投票ボタンは通常メール本文の上にあるインフォバー、または Outlook のリボンの"
            "「応答」欄に表示されます。Outlook Web でも同様にご利用いただけます）\n\n"
            "✅ このリマインドメール自体に直接投票していただいて問題ございません。"
            "（下に添付している最初のご案内メールを開き直す必要はありません。どちらのメールで"
            "投票されても、同じように記録されます。）\n\n"
            "・Yes = 出席\n"
            "・No = 欠席\n"
            "・Maybe = 未定、後日改めてご連絡します\n\n"
            "文字を入力したり件名を変更したりする必要はありません。ボタンをクリックするだけで"
            "自動的に記録されます。\n\n"
            "📎 ボタンの場所が分からない場合は、添付の画像をご確認ください。\n\n"
            "参考までに、元のご案内メールを添付しております。\n\n"
            "よろしくお願いいたします。"
        )

    if lang_code == "vi":
        return (
            "Xin chào,\n\n"
            "Đây là email nhắc nhở thân thiện — chúng tôi vẫn chưa nhận được phản hồi của bạn "
            "cho:\n\n"
            f"{L['event']}: {event_name}\n"
            f"{L['date']}: {event_date}\n"
            f"{L['location']}: {location}\n"
            f"{budget_line}"
            f"{L['deadline']}: {deadline}\n\n"
            ">>> Vui lòng bấm vào một trong các nút Yes / No / Maybe ở phía trên email này <<<\n"
            "(các nút này thường xuất hiện ở thanh thông tin phía trên nội dung email, hoặc "
            "trong mục \"Respond\" trên thanh Ribbon của Outlook — Outlook Web cũng hoạt động "
            "tương tự)\n\n"
            "✅ Bạn có thể bấm vote NGAY TRÊN email nhắc nhở này, không cần mở lại email mời "
            "gốc đính kèm bên dưới — vote trên email nào cũng được ghi nhận như nhau.\n\n"
            "• Yes = xác nhận tham dự\n"
            "• No = không thể tham dự\n"
            "• Maybe = chưa chắc chắn, sẽ phản hồi lại sau\n\n"
            "Không cần gõ gì thêm hay đổi tiêu đề — chỉ cần bấm nút là phản hồi của bạn sẽ được "
            "ghi nhận tự động.\n\n"
            "📎 Chưa biết nút nằm ở đâu? Xem ảnh đính kèm để được hướng dẫn trực quan.\n\n"
            "Email mời gốc được đính kèm để bạn tham khảo.\n\n"
            "Cảm ơn bạn!"
        )

    # default / "en"
    return (
        "Hello,\n\n"
        "This is a friendly reminder — we haven't received your response yet for:\n\n"
        f"{L['event']}: {event_name}\n"
        f"{L['date']}: {event_date}\n"
        f"{L['location']}: {location}\n"
        f"{budget_line}"
        f"{L['deadline']}: {deadline}\n\n"
        ">>> Please click one of the Yes / No / Maybe VOTING BUTTONS at the top of this "
        "email <<<\n"
        "(the buttons usually appear in the info bar above the message body, or in the "
        "\"Respond\" section of the Outlook ribbon — Outlook Web works the same way)\n\n"
        "✅ You can vote directly on THIS reminder email — no need to reopen the original "
        "invite attached below. Voting on either email is recorded the same way.\n\n"
        "  • Yes   = confirming attendance\n"
        "  • No    = cannot attend\n"
        "  • Maybe = not sure yet, will follow up later\n\n"
        "No need to type anything or change the subject — clicking a button records your "
        "response automatically.\n\n"
        "📎 Not sure where the button is? See the attached image for a visual guide.\n\n"
        "The original invitation email is attached for reference.\n\n"
        "Thank you!"
    )


def build_reminder_subject(lang_code, event_id, event_name):
    """Subject của email nhắc nhở — dùng chung cấu trúc [Reminder-{EventID}]
    để cơ chế Scan Inbox (Tab 4) vẫn khớp được Event ID bất kể ngôn ngữ nào
    (Scan chỉ cần Event ID có mặt trong Subject, không quan tâm phần chữ
    còn lại — xem docstring _collect_responses/_lookup_sent_date_hint)."""
    subs = {
        "en": f"[Reminder-{event_id}] {event_name} - Please Confirm Attendance",
        "ja": f"【リマインド-{event_id}】{event_name} - ご回答のお願い",
        "vi": f"[NhacNho-{event_id}] {event_name} - Vui lòng phản hồi",
    }
    if lang_code == "bilingual":
        return subs["ja"] + " / " + subs["en"]
    return subs.get(lang_code, subs["en"])


CALENDAR_LABELS = {
    "en": {"event": "📋 Event", "when": "⏰ Date", "where": "📍 Location", "budget": "💰 Budget"},
    "ja": {"event": "📋 イベント", "when": "⏰ 日時", "where": "📍 場所", "budget": "💰 予算"},
    "vi": {"event": "📋 Sự kiện", "when": "⏰ Ngày", "where": "📍 Địa điểm", "budget": "💰 Ngân sách"},
}


def build_calendar_body(lang_code, event_name="", event_date="", location="", budget=""):
    """Nội dung mặc định của Appointment body (Tab 5 'Attendance & Payment') — lời
    cảm ơn đã vote + nhắc lại thông tin sự kiện (EventName/EventDate/
    Location/Expected event budget, lấy từ Tab 1), gửi kèm Calendar Invite
    chính thức cho những người đã trả lời Yes/Maybe. Hỗ trợ 4 lựa chọn
    giống Tab 3/4: English/Japanese/Vietnamese/Bilingual (Japanese + English,
    ghép theo đúng convention "[English below]" + JA trước + EN sau đã dùng
    ở build_reminder_body()/build_gift_fixed_block()).
    Các tham số event_name/event_date/location/budget để trống ("") vẫn hợp
    lệ (vd lúc mới mở app, Tab 1 chưa điền gì) — chỉ hiện dòng tương ứng
    rỗng, không lỗi."""
    if lang_code == "bilingual":
        ja_full = build_calendar_body("ja", event_name, event_date, location, budget)
        en_full = build_calendar_body("en", event_name, event_date, location, budget)
        return "[English below]\n\n" + ja_full + BILINGUAL_SEPARATOR + en_full

    L = CALENDAR_LABELS.get(lang_code, CALENDAR_LABELS["en"])

    if lang_code == "ja":
        return (
            f"{event_name}への投票、ありがとうございました！\n"
            "正式なカレンダー招待を送付いたしますので、ご確認をお願いいたします。\n\n"
            f"{L['event']}: {event_name}\n"
            f"{L['when']}: {event_date}\n"
            f"{L['where']}: {location}\n"
            f"{L['budget']}: {budget}\n\n"
            "皆様にお会いできるのを楽しみにしております。"
        )

    if lang_code == "vi":
        return (
            f"Cảm ơn bạn đã vote cho sự kiện {event_name}!\n"
            "Tôi xin phép gửi lời mời lịch (Calendar Invite) chính thức như sau đây.\n\n"
            f"{L['event']}: {event_name}\n"
            f"{L['when']}: {event_date}\n"
            f"{L['where']}: {location}\n"
            f"{L['budget']}: {budget}\n\n"
            "Rất mong được gặp mọi người ở buổi tiệc/event!"
        )

    # default / "en"
    return (
        f"Thank you for voting for {event_name}!\n"
        "Please find the official calendar invitation below.\n\n"
        f"{L['event']}: {event_name}\n"
        f"{L['when']}: {event_date}\n"
        f"{L['where']}: {location}\n"
        f"{L['budget']}: {budget}\n\n"
        "Looking forward to seeing everyone there!"
    )


def build_thankyou_subject(lang_code, event_id, event_name):
    """Subject của email cảm ơn sau sự kiện (Tab 5 'Attendance & Payment',
    tính năng mới) — dùng tiền tố RIÊNG '[ThankYou-...]'/'【御礼-...】'/
    '[CamOn-...]' để không lẫn với Invite ('[Confirm-...]'), Reminder
    ('[Reminder-...]') hay Gift Notice ('[Gift-...]') — vẫn giữ event_id
    trong Subject theo đúng convention chung của app (dù email này không
    cần Scan Inbox lại, giữ để tra cứu/lọc mail nhất quán với các loại
    khác)."""
    subs = {
        "en": f"[ThankYou-{event_id}] Thank You for Attending {event_name}",
        "ja": f"【御礼-{event_id}】{event_name}へのご参加ありがとうございました",
        "vi": f"[CamOn-{event_id}] Cảm ơn bạn đã tham gia {event_name}",
    }
    if lang_code == "bilingual":
        return subs["ja"] + " / " + subs["en"]
    return subs.get(lang_code, subs["en"])


THANKYOU_LABELS = {
    "en": {"attend": "👥 Total attendees",
           "col_round": "", "col_attend": "👥 Attendees", "col_collected": "💰 Collected",
           "col_paid": "💸 Paid", "col_remaining": "📊 Remaining",
           "grand_collected": "💰 Total collected (all rounds)",
           "grand_paid": "💸 Total paid (all rounds)",
           "grand_remaining": "📊 Total remaining (all rounds)"},
    "ja": {"attend": "👥 参加人数",
           "col_round": "", "col_attend": "👥 参加人数", "col_collected": "💰 集金額",
           "col_paid": "💸 支払済み", "col_remaining": "📊 残金",
           "grand_collected": "💰 総集金額（全回合計）",
           "grand_paid": "💸 総支払額（全回合計）",
           "grand_remaining": "📊 総残金（全回合計）"},
    "vi": {"attend": "👥 Tổng số người tham gia",
           "col_round": "", "col_attend": "👥 Số người dự", "col_collected": "💰 Đã thu",
           "col_paid": "💸 Đã trả", "col_remaining": "📊 Còn lại",
           "grand_collected": "💰 Tổng tiền thu (tất cả các lần)",
           "grand_paid": "💸 Tổng tiền đã trả (tất cả các lần)",
           "grand_remaining": "📊 Tổng tiền còn lại (tất cả các lần)"},
}


def build_thankyou_body(lang_code, event_name="", event_date="", location="",
                         total_attend="0", rounds_info=None,
                         grand_collected="0", grand_paid="0", grand_remaining="0"):
    """Nội dung mặc định của email cảm ơn sau sự kiện (Tab 5 'Attendance &
    Payment') — lời cảm ơn mọi người đã tham gia (nội dung sự kiện lấy từ
    Tab 1, giống build_calendar_body()), kèm BẢNG tổng kết số liệu
    Attendance & Payment: mỗi đợt thu tiền 1 dòng (số người dự / đã thu /
    đã trả / còn lại — rounds_info là list các tuple (label, attendees,
    collected, paid, remaining), đợt 1 "Amount" luôn là phần tử đầu, xem
    _thankyou_body_args()), rồi 3 dòng tổng của TOÀN BỘ các đợt cộng lại.
    Mọi số liệu LUÔN được tính lại từ dữ liệu MỚI NHẤT ngay trước khi gửi
    (xem _send_thank_you_email()), và có thể cập nhật lại bất cứ lúc nào
    bằng nút "🔄 Update from table" (xem _apply_thankyou_body_lang()) —
    không phải giá trị đông cứng lúc soạn mail. Hỗ trợ 4 lựa chọn ngôn ngữ
    giống các tab khác."""
    if rounds_info is None:
        rounds_info = []
    if lang_code == "bilingual":
        ja_full = build_thankyou_body(
            "ja", event_name, event_date, location,
            total_attend, rounds_info, grand_collected, grand_paid, grand_remaining)
        en_full = build_thankyou_body(
            "en", event_name, event_date, location,
            total_attend, rounds_info, grand_collected, grand_paid, grand_remaining)
        return "[English below]\n\n" + ja_full + BILINGUAL_SEPARATOR + en_full

    L = THANKYOU_LABELS.get(lang_code, THANKYOU_LABELS["en"])
    table = _build_rounds_table(L, rounds_info)
    totals = (
        f"{L['grand_collected']}: {grand_collected}\n"
        f"{L['grand_paid']}: {grand_paid}\n"
        f"{L['grand_remaining']}: {grand_remaining}"
    )

    if lang_code == "ja":
        return (
            f"{event_name}にご参加いただき、誠にありがとうございました！\n"
            "おかげさまで無事に終えることができました。\n\n"
            f"{table}\n\n"
            f"{totals}\n\n"
            "参加費用の詳細は添付の集金リストをご確認ください。\n"
            "改めて、ご参加ありがとうございました。"
        )

    if lang_code == "vi":
        return (
            f"Cảm ơn bạn đã tham gia sự kiện {event_name}!\n"
            "Sự kiện đã diễn ra thành công tốt đẹp.\n\n"
            f"{table}\n\n"
            f"{totals}\n\n"
            "Vui lòng xem file Excel đính kèm để biết chi tiết về người tham gia và chi phí.\n"
            "Một lần nữa xin cảm ơn mọi người đã tham gia!"
        )

    # default / "en"
    return (
        f"Thank you for attending {event_name}!\n"
        "The event went smoothly thanks to everyone's participation.\n\n"
        f"{table}\n\n"
        f"{totals}\n\n"
        "Please see the attached spreadsheet for full attendee and cost details.\n"
        "Thanks again for being there!"
    )


def build_gift_subject(lang_code, event_id, guest_of_honor):
    """Subject email THÔNG BÁO QUYÊN GÓP QUÀ TẶNG (Tab 3 'Send Gift
    Contribution Notice') — dùng tiền tố RIÊNG '[Gift-...]'/'【寄付のお願い-...】'
    /'[QuyenGop-...]' để không lẫn với Invite RSVP thường ('[Confirm-...]')
    hay email nhắc nhở ('[Reminder-...]') — người nhận phân biệt ngay đây
    KHÔNG phải email cần bấm Vote, mà là thông báo kêu gọi đóng góp."""
    subs = {
        "en": f"[Gift-{event_id}] Farewell Gift Contribution for {guest_of_honor}",
        "ja": f"【寄付のお願い-{event_id}】{guest_of_honor}さんへの記念品ご協力のお願い",
        "vi": f"[QuyenGop-{event_id}] Kêu gọi đóng góp quà tặng {guest_of_honor}",
    }
    if lang_code == "bilingual":
        return subs["en"] + " / " + subs["ja"]
    return subs.get(lang_code, subs["en"])


GIFT_LABELS = {
    "en": {"gift_for": "🎁 Gift for", "when": "⏰ When", "where": "📍 Location",
           "contact": "📮 Contact to contribute", "deadline": "⏰ Please contribute by",
           "budget": "💰 Expected gift budget"},
    "ja": {"gift_for": "🎁 贈呈対象", "when": "⏰ 日時", "where": "📍 場所",
           "contact": "📮 ご協力の連絡先", "deadline": "⏰ お申し出期限",
           "budget": "💰 想定予算"},
    "vi": {"gift_for": "🎁 Quà tặng cho", "when": "⏰ Thời gian", "where": "📍 Địa điểm",
           "contact": "📮 Liên hệ đóng góp", "deadline": "⏰ Vui lòng đóng góp trước",
           "budget": "💰 Ngân sách dự kiến"},
}


def build_gift_fixed_block(lang_code, guest_of_honor, start_time, event_date, location,
                            organizer, deadline, gift_budget):
    """FIXED part của email THÔNG BÁO QUYÊN GÓP QUÀ TẶNG — tương tự
    build_fixed_block()/build_reminder_body() nhưng nội dung là kêu gọi
    đóng góp mua quà tặng (KHÔNG phải RSVP, không có Voting Buttons). Nội
    dung theo đúng mẫu người dùng yêu cầu:
    'Chúng tôi sẽ tặng quà cho {GuestOfHonor} vào lúc {StartTime}, ngày
    {EventDate}, ở {Location}. Liên hệ {Organizer} trước {Deadline} nếu
    muốn đóng góp. Expected budget: {GiftBudget}.'
    Hỗ trợ 4 lựa chọn giống các nơi khác: English/Japanese/Vietnamese/
    Bilingual (Japanese + English, ghép theo đúng convention "[English
    below]" + JA trước + EN sau)."""
    if lang_code == "bilingual":
        ja_full = build_gift_fixed_block("ja", guest_of_honor, start_time, event_date, location,
                                          organizer, deadline, gift_budget)
        en_full = build_gift_fixed_block("en", guest_of_honor, start_time, event_date, location,
                                          organizer, deadline, gift_budget)
        return "[English below]\n\n" + ja_full + BILINGUAL_SEPARATOR + en_full

    L = GIFT_LABELS.get(lang_code, GIFT_LABELS["en"])
    when_str = f"{start_time}, {event_date}" if start_time else event_date

    if lang_code == "ja":
        return (
            f"{L['gift_for']}: {guest_of_honor}さん\n"
            f"{L['when']}: {when_str}\n"
            f"{L['where']}: {location}\n\n"
            f"当日、{guest_of_honor}さんへ記念品を贈呈する予定です。\n"
            f"ご協力いただける方は、{L['deadline']}（{deadline}）までに{L['contact']}（{organizer}）までご連絡ください。\n\n"
            f"{L['budget']}: {gift_budget}\n\n"
            "何卒よろしくお願いいたします。"
        )

    if lang_code == "vi":
        return (
            f"{L['gift_for']}: {guest_of_honor}\n"
            f"{L['when']}: {when_str}\n"
            f"{L['where']}: {location}\n\n"
            f"Chúng tôi sẽ tiến hành tặng món quà kỷ niệm cho {guest_of_honor} vào lúc {when_str}, "
            f"tại {location}.\n"
            f"Nếu bạn có ý muốn đóng góp, xin hãy liên hệ {organizer} trước ngày {deadline}.\n\n"
            f"{L['budget']}: {gift_budget}\n\n"
            "Xin chân thành cảm ơn!"
        )

    # default / "en"
    return (
        f"{L['gift_for']}: {guest_of_honor}\n"
        f"{L['when']}: {when_str}\n"
        f"{L['where']}: {location}\n\n"
        f"We will be presenting a farewell gift to {guest_of_honor} at {when_str}, at {location}.\n"
        f"If you would like to contribute, please contact {organizer} before {deadline}.\n\n"
        f"{L['budget']}: {gift_budget}\n\n"
        "Thank you very much!"
    )


def build_gift_reminder_subject(lang_code, event_id, guest_of_honor):
    """Subject email NHẮC NHỞ QUYÊN GÓP QUÀ TẶNG (Tab 6 'Gửi email nhắc nhở
    tới người CHƯA đóng góp') — tiền tố RIÊNG '[Reminder-Gift-...]'/
    '【寄付リマインド-...】'/'[NhacQuyenGop-...]', khác hẳn tiền tố email Gift
    gốc ('[Gift-...]', build_gift_subject()) lẫn tiền tố Reminder RSVP
    thường ('[Reminder-...]', build_reminder_subject()) — để người nhận
    phân biệt ngay đây là NHẮC LẠI lời kêu gọi đóng góp trước đó, không phải
    thông báo mới hay email cần bấm Vote."""
    subs = {
        "en": f"[Reminder-Gift-{event_id}] Friendly Reminder — Gift Contribution for {guest_of_honor}",
        "ja": f"【寄付リマインド-{event_id}】{guest_of_honor}さんへの記念品ご協力のお願い（リマインド）",
        "vi": f"[NhacQuyenGop-{event_id}] Nhắc nhở đóng góp quà tặng {guest_of_honor}",
    }
    if lang_code == "bilingual":
        return subs["en"] + " / " + subs["ja"]
    return subs.get(lang_code, subs["en"])


def _gift_reminder_single_lang_body(lang_code, guest_of_honor, start_time, event_date, location,
                                     organizer, deadline, gift_budget):
    L = GIFT_LABELS.get(lang_code, GIFT_LABELS["en"])
    when_str = f"{start_time}, {event_date}" if start_time else event_date

    if lang_code == "ja":
        return (
            "（このメールはリマインドです。まだご協力の連絡をいただいていない方へお送りしています。）\n\n"
            f"{L['gift_for']}: {guest_of_honor}さん\n"
            f"{L['when']}: {when_str}\n"
            f"{L['where']}: {location}\n\n"
            f"以前ご案内した通り、{guest_of_honor}さんへ記念品を贈呈する予定です。\n"
            f"まだご連絡いただいていない場合は、{L['deadline']}（{deadline}）までに"
            f"{L['contact']}（{organizer}）までご連絡いただけますと幸いです。\n\n"
            f"{L['budget']}: {gift_budget}\n\n"
            "（参考として、以前送付したご案内メールを添付しております。）\n\n"
            "お忙しいところ恐れ入りますが、何卒よろしくお願いいたします。"
        )

    if lang_code == "vi":
        return (
            "(Đây là email nhắc nhở nhẹ — gửi tới những bạn chưa phản hồi đóng góp.)\n\n"
            f"{L['gift_for']}: {guest_of_honor}\n"
            f"{L['when']}: {when_str}\n"
            f"{L['where']}: {location}\n\n"
            f"Như đã thông báo trước đó, chúng tôi sẽ tặng món quà kỷ niệm cho {guest_of_honor} "
            f"vào lúc {when_str}, tại {location}.\n"
            f"Nếu bạn vẫn muốn đóng góp nhưng chưa kịp liên hệ, xin vui lòng liên hệ {organizer} "
            f"trước ngày {deadline} giúp mình nhé.\n\n"
            f"{L['budget']}: {gift_budget}\n\n"
            "(Đính kèm email thông báo gốc để bạn tiện xem lại chi tiết.)\n\n"
            "Không có gì gấp, chỉ là nhắc nhẹ thôi — xin cảm ơn bạn rất nhiều!"
        )

    # default / "en"
    return (
        "(This is a gentle reminder sent only to those who haven't responded yet.)\n\n"
        f"{L['gift_for']}: {guest_of_honor}\n"
        f"{L['when']}: {when_str}\n"
        f"{L['where']}: {location}\n\n"
        f"As announced earlier, we'll be presenting a farewell gift to {guest_of_honor} "
        f"at {when_str}, at {location}.\n"
        f"If you'd still like to contribute but haven't had a chance to reach out yet, "
        f"please contact {organizer} before {deadline}.\n\n"
        f"{L['budget']}: {gift_budget}\n\n"
        "(The original notice email is attached for your reference.)\n\n"
        "No pressure at all — just a friendly nudge. Thank you so much!"
    )


def build_gift_reminder_body(lang_code, guest_of_honor, start_time, event_date, location,
                              organizer, deadline, gift_budget):
    """BODY email nhắc nhở quyên góp — cùng cấu trúc song ngữ với
    build_gift_fixed_block() ('[English below]' + JA trước + EN sau, ghép
    bằng BILINGUAL_SEPARATOR), chỉ khác nội dung là bản NHẮC LẠI thay vì
    thông báo lần đầu."""
    if lang_code == "bilingual":
        ja_full = _gift_reminder_single_lang_body("ja", guest_of_honor, start_time, event_date,
                                                    location, organizer, deadline, gift_budget)
        en_full = _gift_reminder_single_lang_body("en", guest_of_honor, start_time, event_date,
                                                    location, organizer, deadline, gift_budget)
        return "[English below]\n\n" + ja_full + BILINGUAL_SEPARATOR + en_full
    return _gift_reminder_single_lang_body(lang_code, guest_of_honor, start_time, event_date,
                                            location, organizer, deadline, gift_budget)


def build_gift_report_subject(lang_code, event_id, guest_of_honor):
    """Subject email BÁO CÁO số tiền quyên góp đã thu được (Tab 6, tính
    năng mới) — tiền tố RIÊNG '[GiftReport-...]'/'【集金報告-...】'/
    '[BaoCaoQuyenGop-...]' để không lẫn với thông báo kêu gọi đóng góp ban
    đầu ('[Gift-...]', xem build_gift_subject()) hay email nhắc nhở."""
    subs = {
        "en": f"[GiftReport-{event_id}] Contribution Report for {guest_of_honor}",
        "ja": f"【集金報告-{event_id}】{guest_of_honor}さんへの寄付金報告",
        "vi": f"[BaoCaoQuyenGop-{event_id}] Báo cáo quyên góp cho {guest_of_honor}",
    }
    if lang_code == "bilingual":
        return subs["en"] + " / " + subs["ja"]
    return subs.get(lang_code, subs["en"])


GIFT_REPORT_LABELS = {
    "en": {"count": "👥 Total contributors", "total": "💰 Total collected",
           # MỚI — thông tin món quà (Tab 6)
           "gift_item": "🎁 Gift", "gift_link": "🔗 Order link", "gift_price": "💸 Gift cost",
           # MỚI — bảng tổng kết Event + Gift (cột đầu để TRỐNG, giống bảng
           # trong email cảm ơn: tên dòng đã tự nói lên nó là gì)
           "col_kind": "", "col_collected": "💰 Collected", "col_paid": "💸 Paid",
           "col_remaining": "📊 Remaining",
           "row_event": "Party", "row_gift": "Gift",
           "grand_collected": "💰 Total collected (Event + Gift)",
           "grand_paid": "💸 Total paid (Event + Gift)",
           "grand_remaining": "📊 Total remaining (Event + Gift)",
           "gift_only_remaining": "📊 Gift remaining",
           # MỚI — 3 câu văn của phần "gộp với sự kiện (Party)". Đặt ở đây
           # thay vì viết thẳng trong build_gift_report_body() để cả 3 ngôn
           # ngữ nằm cạnh nhau, sửa 1 chỗ là xong. {amount} được thay bằng
           # số dư tổng đã format (vd "9,930").
           "intro_combined": "Together with the party itself, here is the overall summary of "
                              "money collected, money spent, and what is left over.",
           "closing_keep": "The remaining {amount} will be kept for now and put towards our "
                            "next get-together.",
           "closing_bye": "Looking forward to seeing everyone again next time!"},
    "ja": {"count": "👥 寄付者数", "total": "💰 集金総額",
           "gift_item": "🎁 記念品", "gift_link": "🔗 購入リンク", "gift_price": "💸 記念品代",
           "col_kind": "", "col_collected": "💰 集金額", "col_paid": "💸 支払済み",
           "col_remaining": "📊 残金",
           "row_event": "Party", "row_gift": "Gift",
           "grand_collected": "💰 総集金額（イベント＋記念品）",
           "grand_paid": "💸 総支払額（イベント＋記念品）",
           "grand_remaining": "📊 総残金（イベント＋記念品）",
           "gift_only_remaining": "📊 残金",
           "intro_combined": "イベント（Party）と合わせて、収支および残金を下記の通りご報告いたします。",
           "closing_keep": "今回の残金 {amount} 円は一旦保管し、次回の飲み会で活用させていただく予定です。",
           "closing_bye": "また次回もぜひよろしくお願いいたします。"},
    "vi": {"count": "👥 Tổng số người đóng góp", "total": "💰 Tổng tiền thu được",
           "gift_item": "🎁 Món quà", "gift_link": "🔗 Link đặt hàng", "gift_price": "💸 Giá món quà",
           "col_kind": "", "col_collected": "💰 Đã thu", "col_paid": "💸 Đã chi",
           "col_remaining": "📊 Còn lại",
           "row_event": "Sự kiện", "row_gift": "Quà tặng",
           "grand_collected": "💰 Tổng tiền thu (Sự kiện + Quà)",
           "grand_paid": "💸 Tổng tiền chi (Sự kiện + Quà)",
           "grand_remaining": "📊 Tổng tiền còn lại (Sự kiện + Quà)",
           "gift_only_remaining": "📊 Tiền quà còn lại",
           "intro_combined": "Cộng chung với phần sự kiện (Party), dưới đây là tổng kết thu — chi — "
                              "còn lại của cả hai phần.",
           "closing_keep": "Số tiền còn lại {amount} sẽ được giữ lại và dùng cho buổi họp mặt lần sau.",
           "closing_bye": "Rất mong được gặp lại mọi người!"},
}


def build_gift_report_body(lang_code, guest_of_honor="", event_name="", contributor_count="0",
                            total_amount="0", gift_name="", gift_link="", gift_price="",
                            summary_rows=None, grand_collected="0", grand_paid="0",
                            grand_remaining="0", link_event=False, gift_remaining="0"):
    """Nội dung mặc định của email báo cáo quyên góp (Tab 6).

    CẤU TRÚC (đã dựng lại theo đúng bố cục bạn yêu cầu):

      1. Lời cảm ơn + câu dẫn.
      2. KHỐI THÔNG TIN QUÀ — gộp CHUNG mọi con số của phần quà lại 1 chỗ,
         theo đúng thứ tự: tên quà → link đặt hàng → số người đóng góp →
         tổng thu → giá quà → còn lại. (Trước đây số người/tổng thu nằm
         TÁCH khỏi khối quà, phải đọc nhảy cóc 2 nơi mới ghép được bức
         tranh của phần quà.)
      3. CHỈ KHI tick "🔗 Link with Tab 5" (link_event=True):
         câu dẫn "cộng chung với phần sự kiện (Party)" → BẢNG so sánh
         Party/Gift → 3 dòng tổng → câu kết về việc GIỮ LẠI số dư cho lần
         sau.
         Không tick thì bỏ hẳn 3 phần này — bảng chỉ có mỗi dòng Gift là
         thừa, vì mọi số của phần quà đã nằm đủ ở khối (2) rồi.
      4. Nhắc xem file Excel đính kèm.

    gift_remaining / grand_remaining: chuỗi ĐÃ format sẵn (vd "1,500"),
        do _gift_report_body_args() dựng từ số đang hiển thị trên UI.

    Câu kết "giữ lại số dư cho lần sau" chỉ xuất hiện khi số dư THỰC SỰ
    DƯƠNG — nếu đang âm (đã chi nhiều hơn thu) mà vẫn viết "sẽ giữ lại để
    dùng lần sau" thì vô nghĩa và dễ gây hiểu nhầm là đang còn quỹ."""
    if summary_rows is None:
        summary_rows = []
    if lang_code == "bilingual":
        ja_full = build_gift_report_body(
            "ja", guest_of_honor, event_name, contributor_count, total_amount,
            gift_name, gift_link, gift_price, summary_rows,
            grand_collected, grand_paid, grand_remaining, link_event, gift_remaining)
        en_full = build_gift_report_body(
            "en", guest_of_honor, event_name, contributor_count, total_amount,
            gift_name, gift_link, gift_price, summary_rows,
            grand_collected, grand_paid, grand_remaining, link_event, gift_remaining)
        return "[English below]\n\n" + ja_full + BILINGUAL_SEPARATOR + en_full

    L = GIFT_REPORT_LABELS.get(lang_code, GIFT_REPORT_LABELS["en"])

    # ── (2) Khối thông tin quà — bỏ qua dòng nào còn trống ──
    gift_lines = []
    if (gift_name or "").strip():
        gift_lines.append(f"{L['gift_item']}: {gift_name.strip()}")
    if (gift_link or "").strip():
        gift_lines.append(f"{L['gift_link']}: {gift_link.strip()}")
    gift_lines.append(f"{L['count']}: {contributor_count}")
    gift_lines.append(f"{L['total']}: {total_amount}")
    if (gift_price or "").strip():
        gift_lines.append(f"{L['gift_price']}: {gift_price.strip()}")
    gift_lines.append(f"{L['gift_only_remaining']}: {gift_remaining}")
    gift_block = "\n".join(gift_lines) + "\n\n"

    # ── (3) Phần gộp với sự kiện — chỉ khi đang liên kết Tab 5 ──
    combined_block = ""
    closing_block = ""
    if link_event and summary_rows:
        headers = [L["col_kind"], L["col_collected"], L["col_paid"], L["col_remaining"]]
        table = _build_aligned_table(headers, summary_rows)
        combined_block = (
            L["intro_combined"] + "\n\n"
            + table + "\n\n"
            + f"{L['grand_collected']}: {grand_collected}\n"
            + f"{L['grand_paid']}: {grand_paid}\n"
            + f"{L['grand_remaining']}: {grand_remaining}\n\n"
        )
        # Số dư dương -> câu "giữ lại dùng cho lần sau". Âm hoặc bằng 0 ->
        # câu kết trung tính, không hứa hẹn khoản quỹ không tồn tại.
        if (parse_signed_amount(grand_remaining) or 0.0) > 0:
            closing_block = L["closing_keep"].format(amount=grand_remaining) + "\n" + L["closing_bye"]
        else:
            closing_block = L["closing_bye"]

    if lang_code == "ja":
        head = (
            f"{guest_of_honor}さんへの記念品にご協力いただき、誠にありがとうございました！\n"
            "現在までの集金状況を下記の通りご報告いたします。\n\n"
        )
        tail = "詳細（お一人おひとりの内訳）は、添付の集金リストをご確認ください。\n\n"
        if not closing_block:
            closing_block = "改めまして、ご協力いただき誠にありがとうございました。"
    elif lang_code == "vi":
        head = (
            f"Cảm ơn mọi người đã đóng góp quà tặng cho {guest_of_honor}!\n"
            "Đây là báo cáo tình hình quyên góp tính đến thời điểm hiện tại.\n\n"
        )
        tail = "Vui lòng xem file đính kèm để biết chi tiết đóng góp của từng người.\n\n"
        if not closing_block:
            closing_block = "Xin chân thành cảm ơn sự đóng góp của mọi người!"
    else:  # default / "en"
        head = (
            f"Thank you for contributing to the farewell gift for {guest_of_honor}!\n"
            "Here is the contribution report so far.\n\n"
        )
        tail = "Please see the attached spreadsheet for the per-person breakdown.\n\n"
        if not closing_block:
            closing_block = "Thank you again for your generosity!"

    return head + gift_block + combined_block + tail + closing_block


def build_editable_block(lang_code, note, note_is_translated):
    """EDITABLE part: ONLY the organizer's free-text background note for this event."""
    if lang_code == "bilingual":
        return note  # caller combines per-language notes separately for bilingual
    flag = "" if note_is_translated else NOT_TRANSLATED_FLAG.get(lang_code, "")
    return f"{note}{flag}"

"""Bilingual email on Compose & send: Japanese first, English second.

The note box holds only the organizer's note, in any language, and the fixed
box the event details and voting instructions, in Japanese and English either
side of a divider line. One Copilot prompt carries both - the note and the
fixed part as the boxes hold them - and the answer comes back as a Japanese
and an English version of each, under [JA NOTE], [JA DETAILS], [EN NOTE] and
[EN DETAILS]; the notes go back into the note box and the details into the
fixed box. Before this, the note box was filled with the whole bilingual
draft - greeting, event details and voting instructions included - and copying
it for Copilot was refused until a single source language was picked.

The first half needs nothing but rsvp.i18n. The second drives the real Tk app
against the fake Outlook, like tests/test_app_event_state.py, and skips where
that file skips.
"""
from __future__ import annotations

import pytest

from rsvp.i18n import (
    DEFAULT_PROMPT_BILINGUAL,
    LANG_LABELS,
    UPDATE_NOTICE,
    build_bilingual_body,
    build_bilingual_prompt,
    build_gift_fixed_block,
    join_bilingual,
    parse_bilingual_reply,
    split_bilingual,
    translation_extras,
    translation_gaps,
)
from tests.test_app_event_state import app  # noqa: F401

VI_NOTE = "Mọi người nhớ mang theo thẻ nhân viên nhé."
JA_NOTE = "社員証をお持ちください。"
EN_NOTE = "Please bring your staff badge."
JA_DETAILS = "⏰ 「Party」の詳細です。\n\n  • Yes   = 参加します\n  • No    = 参加できません"
EN_DETAILS = '⏰ Details of "Party".\n\n  • Yes   = You attend\n  • No    = You cannot attend'
PARTS = (JA_NOTE, JA_DETAILS, EN_NOTE, EN_DETAILS)


def _reply(ja_note=JA_NOTE, ja_details=JA_DETAILS, en_note=EN_NOTE, en_details=EN_DETAILS):
    return (f"[JA NOTE]\n{ja_note}\n[JA DETAILS]\n{ja_details}\n"
            f"[EN NOTE]\n{en_note}\n[EN DETAILS]\n{en_details}")


# ── rsvp.i18n ────────────────────────────────────────────────────────────

def test_join_and_split_are_inverse():
    assert split_bilingual(join_bilingual(JA_NOTE, EN_NOTE)) == (JA_NOTE, EN_NOTE)


def test_text_without_a_divider_line_is_one_part():
    assert split_bilingual(VI_NOTE) == (VI_NOTE, None)
    # A Japanese dash pair inside a sentence is not the divider line.
    assert split_bilingual("会場――ホールA") == ("会場――ホールA", None)


def test_only_the_first_divider_line_splits():
    text = join_bilingual("ja", join_bilingual("en", "more"))
    assert split_bilingual(text) == ("ja", join_bilingual("en", "more"))


def test_each_half_is_greeting_note_fixed_japanese_first():
    body = build_bilingual_body(JA_NOTE, EN_NOTE, "JA FIXED", "EN FIXED")
    assert body == ("[English below]\n\n皆様\n\n" + JA_NOTE + "\n\nJA FIXED"
                    + "\n\n" + "―" * 28 + "\n\n"
                    + "Hello everyone,\n\n" + EN_NOTE + "\n\nEN FIXED")


def test_an_update_puts_each_notice_before_the_note():
    body = build_bilingual_body(JA_NOTE, EN_NOTE, "JA FIXED", "EN FIXED", is_update=True)
    ja_half, en_half = split_bilingual(body)
    assert ja_half == f"[English below]\n\n皆様\n\n{UPDATE_NOTICE['ja'].strip()}\n\n{JA_NOTE}\n\nJA FIXED"
    assert en_half == f"Hello everyone,\n\n{UPDATE_NOTICE['en'].strip()}\n\n{EN_NOTE}\n\nEN FIXED"


def test_an_empty_note_leaves_no_gap():
    body = build_bilingual_body("", "  ", "JA FIXED", "EN FIXED")
    assert "皆様\n\nJA FIXED" in body and "Hello everyone,\n\nEN FIXED" in body


def test_the_prompts_own_reply_format_is_what_the_parser_reads():
    """The prompt and the parser are two halves of one contract: the example
    reply the prompt shows Copilot must parse."""
    start = DEFAULT_PROMPT_BILINGUAL.index("[JA NOTE]\n")
    last = "<the event details and voting instructions in English>"
    example = DEFAULT_PROMPT_BILINGUAL[start:DEFAULT_PROMPT_BILINGUAL.index(last) + len(last)]
    answered = (example.replace("<the note in Japanese>", JA_NOTE)
                .replace("<the event details and voting instructions in Japanese>", JA_DETAILS)
                .replace("<the note in English>", EN_NOTE)
                .replace("<the event details and voting instructions in English>", EN_DETAILS))
    assert answered != example
    assert parse_bilingual_reply(answered) == PARTS
    # Copilot echoing the template back, unfilled, is not an answer.
    assert parse_bilingual_reply(example) is None


def test_the_note_and_the_fixed_part_sit_in_marked_blocks_before_the_reply_format():
    """Reported from real use: with the note appended after the whole prompt,
    the format example read as an empty slot the note was never put into; and
    the fixed part has to go to Copilot too, for the email to be complete."""
    details = join_bilingual("JA FIXED", "EN FIXED\nP.S.")
    prompt = build_bilingual_prompt(f"  {VI_NOTE}\n", details)
    block = f"===== NOTE =====\n{VI_NOTE}\n===== EVENT DETAILS =====\n{details}\n===== END ====="
    assert block in prompt
    assert prompt.index(block) < prompt.index("[JA NOTE]\n<")
    assert "[NOTE]" not in prompt and "[DETAILS]" not in prompt
    assert prompt.rstrip().endswith("on its own line as in the source.")


def test_an_empty_note_reads_none():
    assert "===== NOTE =====\n(none)\n===== EVENT DETAILS =====" in build_bilingual_prompt(" ", "x")


def test_content_that_says_a_placeholder_is_kept():
    prompt = build_bilingual_prompt("See [DETAILS]", "Bring [NOTE] 2")
    assert "===== NOTE =====\nSee [DETAILS]\n===== EVENT DETAILS =====\nBring [NOTE] 2\n" in prompt


@pytest.mark.parametrize("reply", [
    _reply(),
    "Here is the translation:\n\n" + _reply().replace("[JA NOTE]", "### **[JA NOTE]**")
    .replace("[EN DETAILS]", "---\n\n### **[EN DETAILS]**:"),
    _reply().replace("\n[", "["),                                  # line breaks lost in the copy
    _reply().replace("[JA NOTE]", "【日本語 NOTE】").replace("[EN DETAILS]", "[English-Details]"),
    _reply().lower().replace(JA_NOTE.lower(), JA_NOTE).replace(EN_NOTE.lower(), EN_NOTE)
    .replace(EN_DETAILS.lower(), EN_DETAILS).replace(JA_DETAILS.lower(), JA_DETAILS),
    f"[EN NOTE]\n{EN_NOTE}\n[EN DETAILS]\n{EN_DETAILS}\n"          # English first
    f"[JA NOTE]\n{JA_NOTE}\n[JA DETAILS]\n{JA_DETAILS}",
    _reply("draft", "draft", "draft", "draft") + "\n" + _reply(),   # corrected copy wins
    f"```\n{_reply()}\n```",                                       # inside a code block
    _reply(f"<{JA_NOTE}>", f"<{JA_DETAILS}>", f"<{EN_NOTE}>", f"<{EN_DETAILS}>"),  # brackets kept
])
def test_a_copilot_answer_gives_both_parts_in_both_languages(reply):
    assert parse_bilingual_reply(reply) == PARTS


@pytest.mark.parametrize("ja_note, en_note", [
    ("(none)", "(none)"), ("なし", "None"), ("", ""), ("（なし）", "None."), ("なし。", "N/A"),
    ("特になし", "-"),
])
def test_a_note_written_as_none_is_empty(ja_note, en_note):
    assert parse_bilingual_reply(_reply(ja_note, JA_DETAILS, en_note, EN_DETAILS)) == (
        "", JA_DETAILS, "", EN_DETAILS)


def test_a_note_that_only_starts_like_none_is_kept():
    assert parse_bilingual_reply(_reply("なしでも大丈夫です。", JA_DETAILS, "None of us can be late.",
                                        EN_DETAILS))[::2] == ("なしでも大丈夫です。", "None of us can be late.")


def test_angle_brackets_that_belong_to_the_text_are_kept():
    assert parse_bilingual_reply(_reply(ja_note="<b>注意</b>"))[0] == "<b>注意</b>"
    # Found by Codex review: a part written as <TBD> lost its brackets. Only
    # brackets every part kept are the template's.
    assert parse_bilingual_reply(_reply(ja_note="<未定>", en_note="<TBD>"))[::2] == ("<未定>", "<TBD>")


@pytest.mark.parametrize("source, translated, gaps", [
    ("31/07/2026 18:00, 3,000 JPY", "2026年7月31日 18:00、3000円", []),
    ("Deadline 20/10/2026", "期限 2026年10月", ["20"]),
    ("• Yes • No • Maybe", "• Yes • No", ["Maybe"]),
    # Found by Codex review: the explanation bullets could go while the
    # instruction kept each name once.
    ("Click Yes / No / Maybe.\n• Yes = attend\n• No = cannot\n• Maybe = later",
     "Yes / No / Maybe を押してください。", ["Yes", "No", "Maybe"]),
    ("「Yes」を押す", "Yesを押してください", []),
    ("Ｙｅｓ 予算３，０００", "Yes budget 3000", []),        # full-width source
    ("", "anything 5", []),
    ("18:00 start", "18時開始", []),                         # times compared as times
    ("10:30 start", "10時半開始", []),
    ("10:30 start", "10時30分開始", []),
    ("10:30 start", "10時開始", ["10:30"]),
    # Found by Codex review: a changed on-the-hour time went unnoticed.
    ("18:00 start", "18:30 start", ["18:00"]),
    # Found by Codex review: AM and PM were left out of the comparison.
    ("Start 6:00 PM", "Start 6:00 AM", ["18:00"]),
    ("Start 6:00 PM", "開始 18:00", []),
    ("午後6時開始", "Starts at 6 p.m.", []),
    ("10:30 a.m.", "午前10時半", []),
    ("12:00 PM lunch", "正午 12:00", []),
    # Found by Codex review: zero was left out everywhere, so a lost budget
    # of 0 went unnoticed.
    ("Budget: 0 JPY", "予算: なし", ["0"]),
    # Found by Codex review: a lost minus sign went unnoticed.
    ("Balance: -500", "残高: 500", ["-500"]),
    ("残高 ¥-3,570", "Balance ¥-3570", []),
    ("Remaining \u22121,200", "残り －1200", []),        # typographic and full-width minus
    # Found by Codex review: a dropped plus sign went unnoticed.
    ("Phone: +81-3-1234-5678", "電話: 81-3-1234-5678", ["+81"]),
    ("Phone: +81-3-1234-5678", "電話：＋81-3-1234-5678", []),
    ("2026-12-20, floors 10-12", "2026年12月20日、10〜12階", []),   # hyphens, not signs
    # Found by Codex review: with a set, a dropped deadline equal to the event
    # date went unnoticed.
    ("Event 10/10/2026, reply by 10/10/2026", "イベント 2026年10月10日", ["10 (×2)", "2026"]),
])
def test_translation_gaps_reports_lost_numbers_and_buttons(source, translated, gaps):
    assert translation_gaps(source, translated) == gaps


def test_numbers_added_by_hand_to_one_half_must_reach_both_translations():
    """Found by Codex review: a deadline typed into the Japanese half only
    could be left out of the English translation unnoticed. The template's
    own differences between the halves are not counted."""
    ja_template, en_template = "31/10/2026 開催", "Held 31/10/2026. Click one of the 3 buttons."
    ja_edited = ja_template + "\n締切 25/12/2026"
    assert translation_gaps(en_template, "Held 31/10/2026. Click one of the 3 buttons.",
                            template=en_template, other=(ja_edited, ja_template)) == ["12", "25", "2026"]
    assert translation_gaps(ja_edited, "2026年10月31日開催、締切2026年12月25日",
                            template=ja_template, other=(en_template, en_template)) == []


def test_button_names_added_by_hand_to_one_half_must_reach_both_translations():
    """Found by Codex review: only numbers crossed between the halves."""
    ja_template, en_template = "Yes / No / Maybe を押す", "Click Yes / No / Maybe."
    ja_edited = ja_template + "\n迷ったら Maybe を押してください。"
    assert translation_gaps(en_template, "Click Yes / No / Maybe.", template=en_template,
                            other=(ja_edited, ja_template)) == ["Maybe"]


def test_a_number_the_translation_made_up_is_reported():
    """Found by Codex review: only lost numbers were looked for."""
    assert translation_extras(["Event 10/10/2026"], "Event 10/10/2026, budget 500") == ["500"]
    assert translation_extras(["18:00 start"], "18時開始、19:30終了") == ["19:30"]
    # A number one half has and the other lacks is not made up.
    halves = ["「Yes / No / Maybe」を押す", "Click one of the 3 BUTTONS"]
    assert translation_extras([halves], "3つのボタンのいずれかを押す") == []
    # Found by Codex review: a second copy of a known value is made up too.
    assert translation_extras(["Event 10/10/2026"], "Event 10/10/2026, deadline 10/10/2026") == [
        "10 (×2)", "2026"]
    # A value both in the note and in the details may be in both translations.
    assert translation_extras(["Bring it on 10/10", ["Held 10/10", "10/10 開催"]],
                              "Bring it on 10/10. Held 10/10") == []


def test_a_note_in_two_languages_is_two_versions_of_one_note():
    """Found by Codex review: a note copied after a saved translation holds
    both halves, and its numbers were expected twice in each translation."""
    versions = ["締切 10/10/2026", "Deadline 10/10/2026"]
    assert translation_gaps(versions, "締切 2026年10月10日") == []
    # A number written into one version only must still reach the translation.
    assert translation_gaps(["締切 10/10/2026 18:00", "Deadline 10/10/2026"],
                            "Deadline 10/10/2026") == ["18:00"]


def test_translating_a_saved_translation_again_does_not_expect_a_carried_value_twice():
    """Found by Codex review: once the Japanese translation carried the
    English template's "3" across, translating again expected two."""
    ja_template, en_template = "ボタンを押す 10/10", "Click one of the 3 BUTTONS by 10/10"
    ja_saved, en_saved = "3つのボタンのいずれかを 10/10 までに押す", "Click one of the 3 buttons by 10/10"
    assert translation_gaps(en_saved, "Click one of the 3 buttons by 10/10", template=en_template,
                            other=(ja_saved, ja_template)) == []
    assert translation_gaps(ja_saved, "3つのボタンを 10/10 までに", template=ja_template,
                            other=(en_saved, en_template)) == []


def test_a_hand_added_date_equal_to_the_event_date_is_counted_again():
    """Found by Codex review: taking the larger count let a deadline equal to
    the event date, typed into one half, be dropped from the other."""
    ja_template, en_template = "10/10/2026 開催", "Held 10/10/2026."
    ja_edited = ja_template + "\n締切 10/10/2026"
    assert translation_gaps(en_template, "Held 10/10/2026.", template=en_template,
                            other=(ja_edited, ja_template)) == ["10 (×2)", "2026"]
    # The same deadline typed into both halves counts once.
    en_edited = en_template + " Reply by 10/10/2026."
    assert translation_gaps(en_edited, "Held 10/10/2026. Reply by 10/10/2026.",
                            template=en_template, other=(ja_edited, ja_template)) == []


def test_an_answer_without_note_markers_has_empty_notes():
    reply = f"[JA DETAILS]\n{JA_DETAILS}\n[EN DETAILS]\n{EN_DETAILS}"
    assert parse_bilingual_reply(reply) == ("", JA_DETAILS, "", EN_DETAILS)


@pytest.mark.parametrize("glued", ["[JA DETAILS]", "[EN NOTE]", "[EN DETAILS]"])
def test_one_marker_glued_to_the_line_before_it_still_splits_there(glued):
    """Found by Codex review: with only some line breaks lost, the line-start
    markers alone made a complete answer, and the glued marker's text was
    left inside the part before it."""
    reply = _reply().replace("\n" + glued, glued)
    assert reply.count("\n" + glued) == 0
    assert parse_bilingual_reply(reply) == PARTS


def test_a_corrected_copy_without_notes_replaces_the_drafts_notes():
    """Found by Codex review: a corrected copy of the details only kept the
    draft's notes beside them."""
    reply = (_reply("draft note", "draft", "draft note", "draft")
             + f"\n[JA DETAILS]\n{JA_DETAILS}\n[EN DETAILS]\n{EN_DETAILS}")
    assert parse_bilingual_reply(reply) == ("", JA_DETAILS, "", EN_DETAILS)


def test_template_brackets_go_also_beside_a_none_note():
    """Found by Codex review: "(none)" has no brackets, so the details kept
    the template's."""
    reply = _reply("(none)", f"<{JA_DETAILS}>", "(none)", f"<{EN_DETAILS}>")
    assert parse_bilingual_reply(reply) == ("", JA_DETAILS, "", EN_DETAILS)


def test_a_corrected_copy_with_glued_markers_wins_over_the_draft():
    """Found by Codex review: the draft's line-start markers named every
    part, so the corrected copy's glued markers were thrown away."""
    reply = _reply("draft", "draft", "draft", "draft") + "\n" + _reply().replace("\n[", "[")
    assert parse_bilingual_reply(reply) == PARTS


def test_a_part_that_mentions_its_own_marker_is_not_cut_there():
    """Found by Codex review: after the real [EN NOTE], the same marker
    inside the English note counted as a glued one and cut the note short."""
    en = "Please keep the [EN NOTE] tag here."
    assert parse_bilingual_reply(_reply(en_note=en)) == (JA_NOTE, JA_DETAILS, en, EN_DETAILS)


@pytest.mark.parametrize("en_note", [
    "Please keep [JA DETAILS] exactly.",     # found by Codex review: after the real one
    "Please keep [JA NOTE] and [EN DETAILS] as written.",
])
def test_a_part_that_mentions_another_parts_marker_is_not_cut_there(en_note):
    assert parse_bilingual_reply(_reply(en_note=en_note)) == (
        JA_NOTE, JA_DETAILS, en_note, EN_DETAILS)


def test_glued_markers_of_an_answer_without_notes_still_split():
    assert parse_bilingual_reply(f"[JA DETAILS]\n{JA_DETAILS}[EN DETAILS]\n{EN_DETAILS}") == (
        "", JA_DETAILS, "", EN_DETAILS)


def test_text_that_mentions_a_marker_mid_line_is_not_cut_there():
    ja = "英語版は [EN NOTE] をご覧ください。"
    assert parse_bilingual_reply(_reply(ja_note=ja)) == (ja, JA_DETAILS, EN_NOTE, EN_DETAILS)


@pytest.mark.parametrize("reply", [
    "",
    f"{JA_NOTE}\n\n{EN_NOTE}",
    f"[JA NOTE]\n{JA_NOTE}\n[JA DETAILS]\n{JA_DETAILS}",
    _reply(en_details="  "),
    f"[JA]\n{JA_NOTE}\n[EN]\n{EN_NOTE}",       # the earlier two-marker format
])
def test_an_answer_without_both_details_parts_is_not_read(reply):
    assert parse_bilingual_reply(reply) is None


# ── the app ──────────────────────────────────────────────────────────────

def _box(text_widget):
    return text_widget.get("1.0", "end").strip()


def _set_box(text_widget, text):
    text_widget.delete("1.0", "end")
    text_widget.insert("1.0", text)


def _compose(app, note=VI_NOTE, mode="Send first Invite"):
    """EV1 with one recipient, Tab 1's note set, Email language Bilingual."""
    app.var_event_id.set("EV1")
    app.var_event_name.set("Party")
    app.var_location.set("Hall A")
    app.recipients = [("Alice Example", "alice@example.com")]
    app._refresh_recipient_tree()
    _set_box(app.entry_note, note)
    app.combo_send_mode.set(mode)
    app.combo_email_lang.set(LANG_LABELS["bilingual"])
    app._refresh_compose_preview()


def _answer(app, ja_note=JA_NOTE, en_note=EN_NOTE):
    """A Copilot answer to what the boxes hold now: the details 'translated'
    with every number and button name kept. Returns (reply, ja, en)."""
    ja_fixed, en_fixed = split_bilingual(_box(app.txt_fixed_preview))
    ja, en = f"⏰ {ja_fixed}", f"⏰ {en_fixed}"
    return _reply(ja_note, ja, en_note, en), ja, en


def _paste_and_save(app, reply):
    """Paste and Save, copying first unless the test already did: a
    bilingual Save needs the Copy it answers."""
    if app._current_lang_code() == "bilingual" and app._bilingual_copied is None:
        app._copy_email_for_translation()
    _set_box(app.txt_translation_paste, reply)
    app._save_translated_email()


def _sent_bodies(app):
    return [kw["body"] for name, kw in app.fake.calls if name == "send_voting_invite"]


def _send(app):
    app._send_invite()
    app.update()
    return _sent_bodies(app)


def test_the_note_box_holds_only_the_note(app):
    _compose(app)
    assert _box(app.txt_editable_preview) == VI_NOTE
    ja_fixed, en_fixed = split_bilingual(_box(app.txt_fixed_preview))
    assert ja_fixed.startswith("「Party」を") and "Yes / No / Maybe" in ja_fixed
    assert en_fixed.startswith('Our team is organizing "Party"')
    assert str(app.txt_fixed_preview.cget("state")) == "normal"


def test_bilingual_has_one_translation_target(app, monolith):
    _compose(app)
    assert app.combo_translate_target.get() == monolith.BILINGUAL_TARGET
    assert app.combo_translate_target.instate(["disabled"])
    assert app.var_copy_translation_label.get().startswith("📋 Copy note + fixed part + prompt")


def test_copy_sends_the_note_and_the_fixed_part_in_one_prompt(app):
    _compose(app)
    fixed = _box(app.txt_fixed_preview)
    app._copy_email_for_translation()
    assert app.clipboard_get() == build_bilingual_prompt(VI_NOTE, fixed)
    assert [title for _kind, title, _msg in app.dialogs] == ["Copied"]
    assert VI_NOTE in app.dialogs[-1][2]      # the dialog shows what was copied


def test_copy_uses_what_the_boxes_hold_edits_included(app):
    _compose(app)
    _set_box(app.txt_editable_preview, "Ghi chú mới")
    ja_fixed, en_fixed = split_bilingual(_box(app.txt_fixed_preview))
    edited = join_bilingual(ja_fixed, en_fixed + "\nP.S. Dress code: casual")
    _set_box(app.txt_fixed_preview, edited)
    app._copy_email_for_translation()
    assert (f"===== NOTE =====\nGhi chú mới\n===== EVENT DETAILS =====\n{edited}\n"
            in app.clipboard_get())


def test_a_saved_answer_fills_both_boxes_and_the_sent_email(app):
    _compose(app)
    app._copy_email_for_translation()
    reply, ja, en = _answer(app)
    _paste_and_save(app, "Here it is:\n" + reply)
    assert app._current_lang_code() == "bilingual"
    assert split_bilingual(_box(app.txt_editable_preview)) == (JA_NOTE, EN_NOTE)
    assert split_bilingual(_box(app.txt_fixed_preview)) == (ja, en)

    bodies = _send(app)
    assert len(bodies) == 1
    ja_half, en_half = split_bilingual(bodies[0])
    assert ja_half == f"[English below]\n\n皆様\n\n{JA_NOTE}\n\n{ja}"
    assert en_half == f"Hello everyone,\n\n{EN_NOTE}\n\n{en}"
    assert not [d for d in app.dialogs if d[0] == "askyesno"]


def test_a_saved_answer_with_no_note_leaves_the_note_box_empty(app):
    _compose(app, note="")
    reply, ja, en = _answer(app, "(none)", "（なし）")
    _paste_and_save(app, reply)
    assert _box(app.txt_editable_preview) == ""
    (body,) = _send(app)
    assert f"皆様\n\n{ja}" in body and f"Hello everyone,\n\n{en}" in body
    assert not [d for d in app.dialogs if d[0] == "askyesno"]


def test_saving_without_a_copy_is_refused(app):
    _compose(app)
    _set_box(app.txt_translation_paste, _answer(app)[0])
    app._save_translated_email()
    assert app.dialogs[-1][:2] == ("showwarning", "Copy the email first")
    assert app.full_translations["bilingual"] == ""


def test_an_answer_without_both_parts_saves_nothing(app):
    _compose(app)
    before = (_box(app.txt_editable_preview), _box(app.txt_fixed_preview))
    _paste_and_save(app, f"[JA]\n{JA_NOTE}\n[EN]\n{EN_NOTE}")      # the earlier format
    assert app.full_translations["bilingual"] == ""
    assert (_box(app.txt_editable_preview), _box(app.txt_fixed_preview)) == before
    assert app.dialogs[-1][:2] == ("showwarning", "Japanese and English parts not found")


def test_the_translation_stays_through_a_refresh_and_clearing_rebuilds_from_tab1(app):
    _compose(app)
    untranslated = (_box(app.txt_editable_preview), _box(app.txt_fixed_preview))
    reply, ja, en = _answer(app)
    _paste_and_save(app, reply)
    app._refresh_compose_preview()
    app._refresh_compose_preview_unless_edited()
    assert split_bilingual(_box(app.txt_fixed_preview)) == (ja, en)
    app._clear_translated_email()
    assert (_box(app.txt_editable_preview), _box(app.txt_fixed_preview)) == untranslated
    assert not [d for d in app.dialogs if d[0] in ("askyesno", "showwarning")]


def test_an_edit_to_the_saved_translation_is_what_is_sent(app):
    _compose(app)
    reply, ja, en = _answer(app)
    _paste_and_save(app, reply)
    _set_box(app.txt_fixed_preview, join_bilingual(ja + "\n追記", en + "\nP.S."))
    (body,) = _send(app)
    ja_half, en_half = split_bilingual(body)
    assert ja_half.endswith("追記") and en_half.endswith("P.S.")


def test_a_translation_made_before_event_setup_changed_is_set_aside(app):
    """Found in review: the fixed box came from the saved translation even
    after Tab 1 changed, so an update invite went out with the old venue."""
    _compose(app, mode="Send update invite")
    _paste_and_save(app, _answer(app)[0])
    app.var_location.set("Hall B")
    app._refresh_compose_preview_unless_edited()          # opening the tab again
    assert app.dialogs[-1][:2] == ("showwarning", "Translation set aside")
    assert app.full_translations["bilingual"] == ""
    assert _box(app.txt_editable_preview) == VI_NOTE
    assert "Hall B" in _box(app.txt_fixed_preview)
    (body,) = _send(app)                                  # untranslated note: the fixture says Yes
    assert "Hall B" in body and "Hall A" not in body


def test_the_old_answer_cannot_be_saved_again_after_it_was_set_aside(app):
    """Found by Codex review: Save again with the old answer still in the
    paste box took the new Event setup as its basis."""
    _compose(app)
    _paste_and_save(app, _answer(app)[0])
    app.var_location.set("Hall B")
    app._refresh_compose_preview_unless_edited()          # set aside
    app._save_translated_email()                          # the old answer, no new Copy
    assert app.dialogs[-1][:2] == ("showwarning", "Event setup changed since the Copy")
    assert app.full_translations["bilingual"] == ""
    assert "Hall B" in _box(app.txt_fixed_preview)


def test_an_edited_translation_after_event_setup_changed_is_asked_about(app, monolith, monkeypatch):
    _compose(app)
    reply, ja, en = _answer(app)
    _paste_and_save(app, reply)
    _set_box(app.txt_fixed_preview, join_bilingual(ja + "\n追記", en))   # hand edit: not rebuilt
    app.var_location.set("Hall B")
    app._refresh_compose_preview_unless_edited()
    asked = []
    monkeypatch.setattr(monolith.messagebox, "askyesno",
                        lambda title, message=None, **_: asked.append(title) or False)
    assert _send(app) == []
    assert asked == ["Boxes may be out of date"]


def test_a_copy_of_boxes_left_behind_by_event_setup_is_asked_about(app, monolith, monkeypatch):
    """Found by Codex review: copying hand-edited boxes that still held the
    old details stamped them with the new Event setup, so Save and Send
    took them as current."""
    _compose(app)
    reply, ja, en = _answer(app)
    _paste_and_save(app, reply)
    _set_box(app.txt_fixed_preview, join_bilingual(ja + "\n追記", en))   # hand edit: kept
    app.var_location.set("Hall B")
    app._refresh_compose_preview_unless_edited()
    asked = []
    monkeypatch.setattr(monolith.messagebox, "askyesno",
                        lambda title, message=None, **_: asked.append(title) or False)
    copies_before = app._bilingual_copied
    app._copy_email_for_translation()
    assert asked == ["Boxes may be out of date"]
    assert app._bilingual_copied is copies_before          # nothing copied


def test_edited_untranslated_boxes_left_behind_by_event_setup_are_asked_about(app, monolith,
                                                                             monkeypatch):
    _compose(app, note="")
    _set_box(app.txt_editable_preview, "Ghi chú mới")      # hand edit: kept
    app.var_location.set("Hall B")
    app._refresh_compose_preview_unless_edited()
    assert "Hall B" not in _box(app.txt_fixed_preview)
    asked = []
    monkeypatch.setattr(monolith.messagebox, "askyesno",
                        lambda title, message=None, **_: asked.append(title) or False)
    assert _send(app) == []
    assert asked == ["Boxes may be out of date"]


def test_an_answer_to_a_copy_made_before_event_setup_changed_is_refused(app):
    _compose(app)
    app._copy_email_for_translation()
    reply = _answer(app)[0]
    app.var_location.set("Hall B")
    _paste_and_save(app, reply)
    assert app.dialogs[-1][:2] == ("showwarning", "Event setup changed since the Copy")
    assert app.full_translations["bilingual"] == ""


def test_saving_asks_before_replacing_edits_made_after_the_copy(app, monolith, monkeypatch):
    """Found in review: an edit typed after Copy was silently replaced."""
    _compose(app)
    app._copy_email_for_translation()
    reply = _answer(app)[0]
    after = _box(app.txt_fixed_preview) + "\nP.S. AFTER COPY"
    _set_box(app.txt_fixed_preview, after)
    asked = []
    monkeypatch.setattr(monolith.messagebox, "askyesno",
                        lambda title, message=None, **_: asked.append(title) or False)
    _paste_and_save(app, reply)
    assert asked == ["Replace your edits?"]
    assert _box(app.txt_fixed_preview) == after and app.full_translations["bilingual"] == ""


def test_clearing_leaves_the_boxes_alone_when_nothing_was_saved(app):
    _compose(app)
    _set_box(app.txt_editable_preview, "Ghi chú mới")
    app._clear_translated_email()
    assert _box(app.txt_editable_preview) == "Ghi chú mới"


def test_clearing_asks_before_replacing_edits(app, monolith, monkeypatch):
    _compose(app)
    reply, ja, en = _answer(app)
    _paste_and_save(app, reply)
    _set_box(app.txt_fixed_preview, join_bilingual(ja + "\n追記", en))
    monkeypatch.setattr(monolith.messagebox, "askyesno", lambda *a, **k: False)
    app._clear_translated_email()
    assert _box(app.txt_fixed_preview).startswith(ja + "\n追記")
    assert app.full_translations["bilingual"]
    # Found by Codex review: answering No still emptied the paste box.
    assert _box(app.txt_translation_paste) == reply


def test_an_answer_to_a_copy_from_the_other_send_mode_is_refused(app):
    """Found by Codex review: copied as an invite, saved after switching to
    Gift mode, the invite's details went into the gift notice."""
    app.var_guest_of_honor.set("Guest Example")
    _compose(app)
    app._copy_email_for_translation()
    reply = _answer(app)[0]
    app.combo_send_mode.set("Send Gift Contribution Notice")
    app._refresh_compose_preview()
    _paste_and_save(app, reply)
    assert app.dialogs[-1][:2] == ("showwarning", "Copied in another Send mode")
    assert not app.gift_full_translations["bilingual"] and not app.full_translations["bilingual"]


def test_translating_a_saved_translation_again_raises_no_false_alarm(app):
    _compose(app, note="Hạn chót 10/10/2026")
    reply, ja, en = _answer(app, "締切 2026年10月10日", "Deadline 10/10/2026")
    _paste_and_save(app, reply)
    app._copy_email_for_translation()                  # the note is now JA + EN
    _paste_and_save(app, _answer(app, "締切 2026年10月10日", "Deadline 10/10/2026")[0])
    assert not [d for d in app.dialogs if d[0] == "askyesno"]


def test_a_second_save_after_a_saved_copy_does_not_ask_about_edits(app):
    _compose(app)
    app._copy_email_for_translation()
    reply = _answer(app)[0]
    _paste_and_save(app, reply)
    _paste_and_save(app, reply)          # e.g. after 'Clean up spacing'
    assert not [d for d in app.dialogs if d[0] == "askyesno"]


def test_an_answer_that_drops_the_note_is_refused(app):
    """Found in review: with no NOTE parts, the note vanished from the email."""
    _compose(app)
    ja_fixed, en_fixed = split_bilingual(_box(app.txt_fixed_preview))
    _paste_and_save(app, f"[JA DETAILS]\n{ja_fixed}\n[EN DETAILS]\n{en_fixed}")
    assert app.dialogs[-1][:2] == ("showwarning", "Translated note missing")
    assert app.full_translations["bilingual"] == ""
    assert _box(app.txt_editable_preview) == VI_NOTE


def test_a_date_added_to_one_half_and_lost_from_the_other_translation_is_asked_about(
        app, monolith, monkeypatch):
    _compose(app)
    ja_fixed, en_fixed = split_bilingual(_box(app.txt_fixed_preview))
    _set_box(app.txt_fixed_preview, join_bilingual(ja_fixed + "\n締切 25/12/2026", en_fixed))
    app._copy_email_for_translation()
    reply = _answer(app)[0]              # the English details lack the added date
    asked = []
    monkeypatch.setattr(monolith.messagebox, "askyesno",
                        lambda title, message=None, **_: asked.append((title, message)) or False)
    _paste_and_save(app, reply)
    assert [title for title, _ in asked] == ["Check the translation"]
    assert "English details: 12, 25, 2026" in asked[0][1] and "Japanese details" not in asked[0][1]


def test_a_note_copilot_added_to_an_empty_one_is_asked_about(app, monolith, monkeypatch):
    """Found by Codex review: a note invented for an empty one was saved and
    sent without a word."""
    _compose(app, note="")
    asked = []
    monkeypatch.setattr(monolith.messagebox, "askyesno",
                        lambda title, message=None, **_: asked.append(title) or False)
    _paste_and_save(app, _answer(app, "皆様のご参加をお待ちしております。", "We look forward to seeing you.")[0])
    assert asked == ["Note added by Copilot"]
    assert app.full_translations["bilingual"] == ""
    assert _box(app.txt_editable_preview) == ""


def test_an_answer_with_a_made_up_amount_is_asked_about(app, monolith, monkeypatch):
    _compose(app)
    ja_fixed, en_fixed = split_bilingual(_box(app.txt_fixed_preview))
    asked = []
    monkeypatch.setattr(monolith.messagebox, "askyesno",
                        lambda title, message=None, **_: asked.append((title, message)) or False)
    _paste_and_save(app, _reply(JA_NOTE, ja_fixed + "\n予算 98,765円", EN_NOTE, en_fixed))
    assert [title for title, _ in asked] == ["Check the translation"]
    assert "nowhere in the text you copied" in asked[0][1] and "Japanese details: 98765" in asked[0][1]
    assert app.full_translations["bilingual"] == ""


def test_an_answer_that_loses_a_date_is_asked_about(app, monolith, monkeypatch):
    _compose(app)
    ja_fixed, en_fixed = split_bilingual(_box(app.txt_fixed_preview))
    deadline = monolith.get_date_str(app.date_deadline)
    asked = []
    monkeypatch.setattr(monolith.messagebox, "askyesno",
                        lambda title, message=None, **_: asked.append((title, message)) or False)
    _paste_and_save(app, _reply(JA_NOTE, ja_fixed, EN_NOTE, en_fixed.replace(deadline, "soon")))
    assert [title for title, _ in asked] == ["Check the translation"]
    assert "English details" in asked[0][1] and "Japanese details" not in asked[0][1]
    assert app.full_translations["bilingual"] == ""


def test_an_untranslated_note_is_asked_about_and_no_keeps_it_unsent(app, monolith, monkeypatch):
    _compose(app)
    asked = []
    monkeypatch.setattr(monolith.messagebox, "askyesno",
                        lambda title, message=None, **_: asked.append(title) or False)
    assert _send(app) == []
    assert asked == ["Note not translated"]


def test_an_untranslated_note_sent_anyway_is_in_both_halves(app):
    _compose(app)
    (body,) = _send(app)       # the fixture answers Yes
    ja_half, en_half = split_bilingual(body)
    assert VI_NOTE in ja_half and VI_NOTE in en_half


def test_an_empty_note_can_be_sent_untranslated_or_copied_as_none(app):
    _compose(app, note="")
    app._copy_email_for_translation()
    assert "===== NOTE =====\n(none)\n" in app.clipboard_get()
    (body,) = _send(app)
    assert "皆様\n\n「Party」を" in body and 'Hello everyone,\n\nOur team is organizing' in body
    assert not [d for d in app.dialogs if d[0] == "askyesno"]


def test_a_fixed_part_that_lost_its_divider_line_is_not_sent(app):
    _compose(app, note="")
    _set_box(app.txt_fixed_preview, _box(app.txt_fixed_preview).replace("―", ""))
    assert _send(app) == []
    assert app.dialogs[-1][:2] == ("showwarning", "Divider line missing")


@pytest.mark.parametrize("side", [0, 1])
def test_a_fixed_part_with_an_empty_half_is_not_sent(app, side):
    _compose(app, note="")
    halves = list(split_bilingual(_box(app.txt_fixed_preview)))
    halves[side] = ""
    _set_box(app.txt_fixed_preview, join_bilingual(*halves))
    assert _send(app) == []
    assert app.dialogs[-1][:2] == ("showwarning", "Half of the fixed part is empty")


def test_a_note_missing_from_one_half_is_asked_about(app, monolith, monkeypatch):
    _compose(app)
    _set_box(app.txt_editable_preview, join_bilingual(JA_NOTE, ""))
    asked = []
    monkeypatch.setattr(monolith.messagebox, "askyesno",
                        lambda title, message=None, **_: asked.append(title) or False)
    assert _send(app) == []
    assert asked == ["Note not translated"]


def test_a_hand_edited_fixed_part_is_what_is_sent(app):
    _compose(app, note="")
    ja_fixed, en_fixed = split_bilingual(_box(app.txt_fixed_preview))
    _set_box(app.txt_fixed_preview, join_bilingual(ja_fixed + "\n追記", en_fixed + "\nP.S."))
    (body,) = _send(app)
    ja_half, en_half = split_bilingual(body)
    assert ja_half.endswith("追記") and en_half.endswith("P.S.")


def test_update_mode_sends_each_notice_before_the_note_and_shows_it_in_no_box(app):
    """As in a single-language update: greeting, notice, note, details. The
    notice is fixed wording, so neither box holds it; the greeting line says
    it is added."""
    _compose(app, mode="Send update invite")
    assert "⚠️" not in _box(app.txt_editable_preview) + _box(app.txt_fixed_preview)
    reply, ja, en = _answer(app)
    _paste_and_save(app, reply)
    assert "⚠️ 重要" not in _box(app.txt_editable_preview) + _box(app.txt_fixed_preview)
    assert "change notice" in app.var_greeting_preview.get()
    (body,) = _send(app)
    ja_half, en_half = split_bilingual(body)
    assert ja_half == f"[English below]\n\n皆様\n\n{UPDATE_NOTICE['ja'].strip()}\n\n{JA_NOTE}\n\n{ja}"
    assert en_half == f"Hello everyone,\n\n{UPDATE_NOTICE['en'].strip()}\n\n{EN_NOTE}\n\n{en}"


def test_gift_mode_fixed_part_is_the_gift_notice_in_both_languages(app, monolith):
    app.var_guest_of_honor.set("Guest Example")
    _compose(app, mode="Send Gift Contribution Notice")
    args = ("Guest Example", app.var_start_time.get(), monolith.get_date_str(app.date_event),
            "Hall A", app.var_organizer.get(), monolith.get_date_str(app.date_gift_deadline),
            app.var_gift_budget.get())
    assert split_bilingual(_box(app.txt_fixed_preview)) == (
        build_gift_fixed_block("ja", *args), build_gift_fixed_block("en", *args))
    reply, ja, en = _answer(app)
    _paste_and_save(app, reply)
    assert app.gift_full_translations["bilingual"] and not app.full_translations["bilingual"]
    assert split_bilingual(_box(app.txt_fixed_preview)) == (ja, en)
    assert not [d for d in app.dialogs if d[0] == "askyesno"]


def test_a_single_language_offers_the_three_targets_again(app, monolith):
    _compose(app)
    app.combo_email_lang.set(LANG_LABELS["vi"])
    app._refresh_compose_preview()
    assert not app.combo_translate_target.instate(["disabled"])
    assert list(app.combo_translate_target.cget("values")) == monolith.SINGLE_TARGETS
    assert app.combo_translate_target.get() == "Japanese"
    assert app.var_copy_translation_label.get() == "📋 Copy full email + prompt"

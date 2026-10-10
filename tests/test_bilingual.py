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
    assert parse_bilingual_reply(example) == (
        "the note in Japanese", "the event details and voting instructions in Japanese",
        "the note in English", "the event details and voting instructions in English")


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


@pytest.mark.parametrize("ja_note, en_note", [("(none)", "(none)"), ("なし", "None"), ("", "")])
def test_a_note_written_as_none_is_empty(ja_note, en_note):
    assert parse_bilingual_reply(_reply(ja_note, JA_DETAILS, en_note, EN_DETAILS)) == (
        "", JA_DETAILS, "", EN_DETAILS)


def test_an_answer_without_note_markers_has_empty_notes():
    reply = f"[JA DETAILS]\n{JA_DETAILS}\n[EN DETAILS]\n{EN_DETAILS}"
    assert parse_bilingual_reply(reply) == ("", JA_DETAILS, "", EN_DETAILS)


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


def _paste_and_save(app, reply):
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
    _paste_and_save(app, "Here it is:\n" + _reply())
    assert app._current_lang_code() == "bilingual"
    assert split_bilingual(_box(app.txt_editable_preview)) == (JA_NOTE, EN_NOTE)
    assert split_bilingual(_box(app.txt_fixed_preview)) == (JA_DETAILS, EN_DETAILS)

    bodies = _send(app)
    assert len(bodies) == 1
    ja_half, en_half = split_bilingual(bodies[0])
    assert ja_half == f"[English below]\n\n皆様\n\n{JA_NOTE}\n\n{JA_DETAILS}"
    assert en_half == f"Hello everyone,\n\n{EN_NOTE}\n\n{EN_DETAILS}"
    assert not [d for d in app.dialogs if d[0] == "askyesno"]


def test_a_saved_answer_with_no_note_leaves_the_note_box_empty(app):
    _compose(app, note="")
    _paste_and_save(app, _reply("(none)", JA_DETAILS, "(none)", EN_DETAILS))
    assert _box(app.txt_editable_preview) == ""
    (body,) = _send(app)
    assert f"皆様\n\n{JA_DETAILS}" in body and f"Hello everyone,\n\n{EN_DETAILS}" in body
    assert not [d for d in app.dialogs if d[0] == "askyesno"]


def test_an_answer_without_both_parts_saves_nothing(app):
    _compose(app)
    before = (_box(app.txt_editable_preview), _box(app.txt_fixed_preview))
    _paste_and_save(app, f"[JA]\n{JA_NOTE}\n[EN]\n{EN_NOTE}")
    assert app.full_translations["bilingual"] == ""
    assert (_box(app.txt_editable_preview), _box(app.txt_fixed_preview)) == before
    assert app.dialogs[-1][:2] == ("showwarning", "Japanese and English parts not found")


def test_the_translation_stays_through_a_refresh_and_clearing_rebuilds_from_tab1(app):
    _compose(app)
    untranslated = (_box(app.txt_editable_preview), _box(app.txt_fixed_preview))
    _paste_and_save(app, _reply())
    app._refresh_compose_preview()
    assert split_bilingual(_box(app.txt_fixed_preview)) == (JA_DETAILS, EN_DETAILS)
    app._clear_translated_email()
    assert (_box(app.txt_editable_preview), _box(app.txt_fixed_preview)) == untranslated


def test_an_edit_to_the_saved_translation_is_what_is_sent(app):
    _compose(app)
    _paste_and_save(app, _reply())
    _set_box(app.txt_fixed_preview, join_bilingual(JA_DETAILS + "\n追記", EN_DETAILS + "\nP.S."))
    (body,) = _send(app)
    ja_half, en_half = split_bilingual(body)
    assert ja_half.endswith("追記") and en_half.endswith("P.S.")


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
    _paste_and_save(app, _reply())
    assert "⚠️ 重要" not in _box(app.txt_editable_preview) + _box(app.txt_fixed_preview)
    assert "change notice" in app.var_greeting_preview.get()
    (body,) = _send(app)
    ja_half, en_half = split_bilingual(body)
    assert ja_half == f"[English below]\n\n皆様\n\n{UPDATE_NOTICE['ja'].strip()}\n\n{JA_NOTE}\n\n{JA_DETAILS}"
    assert en_half == f"Hello everyone,\n\n{UPDATE_NOTICE['en'].strip()}\n\n{EN_NOTE}\n\n{EN_DETAILS}"


def test_gift_mode_fixed_part_is_the_gift_notice_in_both_languages(app, monolith):
    app.var_guest_of_honor.set("Guest Example")
    _compose(app, mode="Send Gift Contribution Notice")
    args = ("Guest Example", app.var_start_time.get(), monolith.get_date_str(app.date_event),
            "Hall A", app.var_organizer.get(), monolith.get_date_str(app.date_gift_deadline),
            app.var_gift_budget.get())
    assert split_bilingual(_box(app.txt_fixed_preview)) == (
        build_gift_fixed_block("ja", *args), build_gift_fixed_block("en", *args))
    _paste_and_save(app, _reply())
    assert app.gift_full_translations["bilingual"] and not app.full_translations["bilingual"]
    assert split_bilingual(_box(app.txt_fixed_preview)) == (JA_DETAILS, EN_DETAILS)


def test_a_single_language_offers_the_three_targets_again(app, monolith):
    _compose(app)
    app.combo_email_lang.set(LANG_LABELS["vi"])
    app._refresh_compose_preview()
    assert not app.combo_translate_target.instate(["disabled"])
    assert list(app.combo_translate_target.cget("values")) == monolith.SINGLE_TARGETS
    assert app.combo_translate_target.get() == "Japanese"
    assert app.var_copy_translation_label.get() == "📋 Copy full email + prompt"

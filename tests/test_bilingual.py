"""Bilingual email on Compose & send: Japanese first, English second.

The note box holds only the organizer's note, in any language; the greeting
and the fixed part already exist in both languages, so only the note goes to
Copilot, in one prompt, and comes back as a [JA] and an [EN] part. Before
this, the note box was filled with the whole bilingual draft - greeting,
event details and voting instructions included - and copying it for Copilot
was refused until a single source language was picked.

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
    build_gift_fixed_block,
    join_bilingual,
    parse_bilingual_reply,
    split_bilingual,
)
from tests.test_app_event_state import app  # noqa: F401

VI_NOTE = "Mọi người nhớ mang theo thẻ nhân viên nhé."
JA_NOTE = "社員証をお持ちください。"
EN_NOTE = "Please bring your staff badge."


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


def test_an_empty_note_leaves_no_gap():
    body = build_bilingual_body("", "  ", "JA FIXED", "EN FIXED")
    assert "皆様\n\nJA FIXED" in body and "Hello everyone,\n\nEN FIXED" in body


def test_the_prompts_own_reply_format_is_what_the_parser_reads():
    """The prompt and the parser are two halves of one contract: the example
    reply the prompt shows Copilot must parse."""
    start = DEFAULT_PROMPT_BILINGUAL.index("[JA]\n")
    end = DEFAULT_PROMPT_BILINGUAL.index("(the note in English)") + len("(the note in English)")
    example = DEFAULT_PROMPT_BILINGUAL[start:end]
    assert parse_bilingual_reply(example) == ("(the note in Japanese)", "(the note in English)")


@pytest.mark.parametrize("reply", [
    f"[JA]\n{JA_NOTE}\n[EN]\n{EN_NOTE}",
    f"Here is the translation:\n\n### **[JA]**\n{JA_NOTE}\n\n---\n\n### **[EN]**:\n{EN_NOTE}\n",
    f"[JA]{JA_NOTE}[EN]{EN_NOTE}",                       # line breaks lost in the copy
    f"【日本語】\n{JA_NOTE}\n【English】\n{EN_NOTE}",
    f"[ja] {JA_NOTE}\n[en] {EN_NOTE}",
    f"[EN]\n{EN_NOTE}\n[JA]\n{JA_NOTE}",                 # English first
    f"[JA]\ndraft\n[EN]\ndraft\n[JA]\n{JA_NOTE}\n[EN]\n{EN_NOTE}",  # corrected copy wins
])
def test_a_copilot_answer_gives_the_japanese_and_english_note(reply):
    assert parse_bilingual_reply(reply) == (JA_NOTE, EN_NOTE)


@pytest.mark.parametrize("reply", [
    "", f"{JA_NOTE}\n\n{EN_NOTE}", f"[JA]\n{JA_NOTE}", f"[JA]\n{JA_NOTE}\n[EN]\n  \n",
])
def test_an_answer_without_both_parts_is_not_read(reply):
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
    assert app.var_copy_translation_label.get().startswith("📋 Copy note + prompt")


def test_copy_sends_only_the_note_in_one_prompt(app):
    _compose(app)
    app._copy_email_for_translation()
    assert app.clipboard_get() == DEFAULT_PROMPT_BILINGUAL + "\n\n" + VI_NOTE
    assert [title for _kind, title, _msg in app.dialogs] == ["Copied"]


def test_copy_uses_a_note_typed_on_compose(app):
    _compose(app)
    _set_box(app.txt_editable_preview, "Ghi chú mới")
    app._copy_email_for_translation()
    assert app.clipboard_get().endswith("\n\nGhi chú mới")


def test_a_saved_answer_becomes_the_note_in_each_half_of_the_sent_email(app):
    _compose(app)
    _paste_and_save(app, f"Here it is:\n[JA]\n{JA_NOTE}\n[EN]\n{EN_NOTE}")
    assert app._current_lang_code() == "bilingual"
    assert split_bilingual(_box(app.txt_editable_preview)) == (JA_NOTE, EN_NOTE)

    bodies = _send(app)
    assert len(bodies) == 1
    ja_half, en_half = split_bilingual(bodies[0])
    assert ja_half.startswith(f"[English below]\n\n皆様\n\n{JA_NOTE}\n\n「Party」を")
    assert en_half.startswith(f'Hello everyone,\n\n{EN_NOTE}\n\nOur team is organizing "Party"')
    assert VI_NOTE not in bodies[0]
    assert "未翻訳" not in bodies[0] and "not yet translated" not in bodies[0]
    assert not [d for d in app.dialogs if d[0] == "askyesno"]


def test_an_answer_without_both_parts_saves_nothing(app):
    _compose(app)
    _paste_and_save(app, f"{JA_NOTE}\n\n{EN_NOTE}")
    assert app.full_translations["bilingual"] == ""
    assert _box(app.txt_editable_preview) == VI_NOTE
    assert app.dialogs[-1][:2] == ("showwarning", "Japanese and English parts not found")


def test_clearing_the_translation_brings_back_the_note(app):
    _compose(app)
    _paste_and_save(app, f"[JA]\n{JA_NOTE}\n[EN]\n{EN_NOTE}")
    app._clear_translated_email()
    assert _box(app.txt_editable_preview) == VI_NOTE


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


def test_an_empty_note_needs_no_translation(app):
    _compose(app, note="")
    app._copy_email_for_translation()
    assert app.dialogs[-1][:2] == ("showinfo", "Nothing to translate")
    (body,) = _send(app)
    assert "皆様\n\n「Party」を" in body and 'Hello everyone,\n\nOur team is organizing' in body
    assert not [d for d in app.dialogs if d[0] == "askyesno"]


def test_a_fixed_part_that_lost_its_divider_line_is_not_sent(app):
    _compose(app, note="")
    _set_box(app.txt_fixed_preview, _box(app.txt_fixed_preview).replace("―", ""))
    assert _send(app) == []
    assert app.dialogs[-1][:2] == ("showwarning", "Divider line missing")


def test_a_hand_edited_fixed_part_is_what_is_sent(app):
    _compose(app, note="")
    ja_fixed, en_fixed = split_bilingual(_box(app.txt_fixed_preview))
    _set_box(app.txt_fixed_preview, join_bilingual(ja_fixed + "\n追記", en_fixed + "\nP.S."))
    (body,) = _send(app)
    ja_half, en_half = split_bilingual(body)
    assert ja_half.endswith("追記") and en_half.endswith("P.S.")


def test_update_mode_opens_each_fixed_half_with_its_own_notice(app):
    _compose(app, mode="Send update invite")
    ja_fixed, en_fixed = split_bilingual(_box(app.txt_fixed_preview))
    assert ja_fixed.startswith(UPDATE_NOTICE["ja"].strip())
    assert en_fixed.startswith(UPDATE_NOTICE["en"].strip())
    assert "⚠️" not in _box(app.txt_editable_preview)


def test_gift_mode_fixed_part_is_the_gift_notice_in_both_languages(app, monolith):
    app.var_guest_of_honor.set("Guest Example")
    _compose(app, mode="Send Gift Contribution Notice")
    args = ("Guest Example", app.var_start_time.get(), monolith.get_date_str(app.date_event),
            "Hall A", app.var_organizer.get(), monolith.get_date_str(app.date_gift_deadline),
            app.var_gift_budget.get())
    assert split_bilingual(_box(app.txt_fixed_preview)) == (
        build_gift_fixed_block("ja", *args), build_gift_fixed_block("en", *args))
    _paste_and_save(app, f"[JA]\n{JA_NOTE}\n[EN]\n{EN_NOTE}")
    assert app.gift_full_translations["bilingual"] and not app.full_translations["bilingual"]


def test_a_single_language_offers_the_three_targets_again(app, monolith):
    _compose(app)
    app.combo_email_lang.set(LANG_LABELS["vi"])
    app._refresh_compose_preview()
    assert not app.combo_translate_target.instate(["disabled"])
    assert list(app.combo_translate_target.cget("values")) == monolith.SINGLE_TARGETS
    assert app.combo_translate_target.get() == "Japanese"
    assert app.var_copy_translation_label.get() == "📋 Copy full email + prompt"

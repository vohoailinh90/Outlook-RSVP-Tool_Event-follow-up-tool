"""Render every message builder, for comparison against tests/golden/.

Shared by the parity test and the regeneration script so the two can never
drift apart and silently compare different things.
"""
from __future__ import annotations

import rsvp.i18n as i18n

LANGS = ["en", "ja", "vi", "bilingual", "de", ""]

E = dict(event_name="Year End Party", event_date="2026-12-20 18:00",
         location="Hall A", deadline="2026-12-10", budget="3,000 JPY / person")
G = dict(guest_of_honor="Tanaka-san", start_time="18:00",
         event_date="2026-12-20", location="Hall A")

PASTE = [
    "", "Hello\nHello\nWorld", "  spaced  \n\n\n\nlines  ",
    "**bold** and *italic*\n- bullet\n# header",
    "Line one\nLine one\nLine two\nLine two\nLine three",
    "日本語のテキスト\n日本語のテキスト", "Tiếng Việt có dấu\n\nvà xuống dòng",
    "A" * 200 + "\n" + "A" * 200,
    # Emoji inputs exercise EMOJI_PATTERN, which the cleanup functions use as
    # an anchor when repairing lost line breaks. Without one of these the
    # pattern is exported but never executed, so a corrupted character range
    # would survive the move undetected. Found in review of phase 1.
    "⏰ 18:00📍 Hall A💰 3,000 JPY",
    "Hello everyone,⏰ Please reply by Friday📋 Bring your badge👥 12 people",
    "⚠️ Updated:⏰ new time📍 new room",
]

CONSTANTS = [
    "LANG_LABELS", "LANG_LABEL_TO_CODE", "TRANSLATE_TARGETS", "BILINGUAL_SEPARATOR",
    "NOT_TRANSLATED_FLAG", "GREETING", "UPDATE_NOTICE", "REMINDER_LABELS",
    "CALENDAR_LABELS", "THANKYOU_LABELS", "GIFT_LABELS", "GIFT_REPORT_LABELS",
    "DEFAULT_PROMPT_SINGLE", "DEFAULT_PROMPT_BILINGUAL", "EMOJI_PATTERN",
]


def _calls():
    for L in LANGS:
        yield "build_greeting", "build_greeting", (L,), {}
        yield "build_update_notice", "build_update_notice", (L,), {}
        yield "build_subject", "build_subject", (L, "E01", "Year End Party"), {}
        yield "build_subject_upd", "build_subject", (L, "E01", "Year End Party"), {"is_update": True}
        yield "build_reminder_subject", "build_reminder_subject", (L, "E01", "Year End Party"), {}
        yield "build_thankyou_subject", "build_thankyou_subject", (L, "E01", "Year End Party"), {}
        yield "build_gift_subject", "build_gift_subject", (L, "E01", "Tanaka-san"), {}
        yield "build_gift_reminder_subject", "build_gift_reminder_subject", (L, "E01", "Tanaka-san"), {}
        yield "build_gift_report_subject", "build_gift_report_subject", (L, "E01", "Tanaka-san"), {}
        yield ("build_fixed_block", "build_fixed_block",
               (L, E["event_name"], E["event_date"], E["location"], E["deadline"], E["budget"]), {})
        yield ("build_reminder_body", "build_reminder_body",
               (L, E["event_name"], E["event_date"], E["location"], E["deadline"], E["budget"]), {})
        yield ("build_calendar_body", "build_calendar_body", (L,),
               dict(event_name=E["event_name"], event_date=E["event_date"],
                    location=E["location"], budget=E["budget"]))
        yield ("build_thankyou_body", "build_thankyou_body", (L,),
               dict(event_name=E["event_name"], event_date=E["event_date"], location=E["location"],
                    total_attend="12", total_collected="36000", amount_paid="30000",
                    remaining_amount="6000"))
        yield ("build_gift_report_body_full", "build_gift_report_body", (L,),
               dict(guest_of_honor=G["guest_of_honor"], event_name=E["event_name"],
                    contributor_count="7", total_amount="35000"))
        yield ("build_gift_fixed_block", "build_gift_fixed_block",
               (L, G["guest_of_honor"], G["start_time"], G["event_date"], G["location"],
                "Sato-san", "2026-12-15", "5,000 JPY"), {})
        yield ("build_gift_reminder_body", "build_gift_reminder_body",
               (L, G["guest_of_honor"], G["start_time"], G["event_date"], G["location"],
                "Sato-san", "2026-12-15", "5,000 JPY"), {})
        yield ("build_gift_report_body", "build_gift_report_body", (L,),
               dict(guest_of_honor=G["guest_of_honor"], event_name=E["event_name"],
                    contributor_count="7"))
        yield "build_editable_block", "build_editable_block", (L, "See you there!", False), {}
        yield "build_editable_block_t", "build_editable_block", (L, "See you there!", True), {}

    for text in PASTE:
        for fn in ("cleanup_pasted_translation", "dedupe_pasted_translation",
                   "detect_possible_duplicate_paste"):
            yield fn, fn, (text,), {}


def build_snapshot(module=None) -> dict[str, str]:
    """Render everything. Keys match the golden file exactly.

    `module` defaults to rsvp.i18n. scripts/verify_golden_baseline.py passes
    the PRE-EXTRACTION rsvp_app.py instead, reconstructed from git history, so
    the golden file can be re-derived from the parent commit rather than taken
    on trust. Both must produce identical output; that is the whole claim.
    """
    module = module or i18n
    out: dict[str, str] = {}
    for label, fn_name, args, kw in _calls():
        key = f"{label}({args!r},{sorted(kw.items())!r})"
        fn = getattr(module, fn_name, None)
        if fn is None:
            out[key] = "!!MISSING"
            continue
        try:
            out[key] = repr(fn(*args, **kw))
        except Exception as exc:
            out[key] = f"!!RAISED {type(exc).__name__}: {exc}"
    for name in CONSTANTS:
        out[f"CONST:{name}"] = repr(getattr(module, name, "!!MISSING"))
    return out


# ── After phase 1 ────────────────────────────────────────────────────────
#
# Two builders were later replaced on purpose by the versions in the
# separately developed copy of this app that the user runs: the thank-you
# email gained a table of payment rounds, the gift report the gift item and
# Event + Gift totals. Their phase-1 calls above use the old signatures and
# no longer render. These sets are the ONLY entries of the phase-1 golden
# file allowed to differ; tests/test_i18n_parity.py also requires each of
# them to really differ, so the list cannot quietly hide an unchanged
# builder. Widening it is a reviewed code change, not a data edit.
RETIRED_BUILDERS = frozenset({"build_thankyou_body", "build_gift_report_body"})
RETIRED_CONSTANTS = frozenset({"THANKYOU_LABELS", "GIFT_REPORT_LABELS"})


def is_retired(key: str) -> bool:
    if key.startswith("CONST:"):
        return key[len("CONST:"):] in RETIRED_CONSTANTS
    return key.split("(")[0].removesuffix("_full") in RETIRED_BUILDERS


# Rounds as _thankyou_body_args() passes them: (label, attendees, collected,
# paid, remaining), already formatted.
ROUNDS = [
    [],
    [("Round 1", "14", "82,000", "70,000", "12,000")],
    [("Round 1", "14", "82,000", "70,000", "12,000"),
     ("Karaoke", "5", "10,000", "13,570", "-3,570")],
    [("Dinner (tiệc tối)", "3", "0", "0", "0")],   # renamed, wide characters
]
GIFT_ROWS = [("Party", "82,000", "70,000", "12,000"), ("Gift", "0", "3,570", "-3,570")]
HTML_BODIES = [
    "Hello\n\nNo table here.",
    "Totals:\n\n  │ 👥 Attendees │ 💰 Collected\n  ─┼─\n  Round 1 │ 14 │ 82,000\n\nThanks!",
    "Order at https://shop.example.com/item?a=1&b=2. Thanks",
    "A <b>tag</b> & an \"ampersand\"\n――――――――――――――――――――――――――――\nsecond half",
    "See https://example.com/a), (https://example.com/b). 詳細：https://example.com/c。",
    "  Round │ Link\n  ─┼─\n  A & B │ https://example.com/r?x=1&y=2",
]
MERGE_CONSTANTS = ["THANKYOU_LABELS", "GIFT_REPORT_LABELS"]


def _merge_calls():
    for L in LANGS:
        for i, rounds in enumerate(ROUNDS):
            yield (f"build_thankyou_body_r{i}", "build_thankyou_body", (L,),
                   dict(event_name=E["event_name"], event_date=E["event_date"],
                        location=E["location"], total_attend="14", rounds_info=rounds,
                        grand_collected="92,000", grand_paid="83,570",
                        grand_remaining="8,430"))
        for remaining in ("8,430", "0", "-1,200"):
            yield (f"build_gift_report_body_linked_{remaining}", "build_gift_report_body", (L,),
                   dict(guest_of_honor=G["guest_of_honor"], event_name=E["event_name"],
                        contributor_count="7", total_amount="0", gift_name="Speaker",
                        gift_link="https://shop.example.com/item", gift_price="3,570",
                        summary_rows=GIFT_ROWS, grand_collected="82,000",
                        grand_paid="73,570", grand_remaining=remaining, link_event=True,
                        gift_remaining="-3,570"))
        yield ("build_gift_report_body_unlinked", "build_gift_report_body", (L,),
               dict(guest_of_honor=G["guest_of_honor"], event_name=E["event_name"],
                    contributor_count="7", total_amount="35,000", gift_name="",
                    gift_link="", gift_price="", summary_rows=[GIFT_ROWS[1]],
                    link_event=False, gift_remaining="35,000"))
        # A gift item without a link, not added to the event's money.
        yield ("build_gift_report_body_item_no_link", "build_gift_report_body", (L,),
               dict(guest_of_honor=G["guest_of_honor"], event_name=E["event_name"],
                    contributor_count="7", total_amount="35,000", gift_name="Speaker",
                    gift_link="", gift_price="3,570", summary_rows=[GIFT_ROWS[1]],
                    link_event=False, gift_remaining="31,430"))
        # Added to the event's money with no gift item chosen yet.
        yield ("build_gift_report_body_linked_no_item", "build_gift_report_body", (L,),
               dict(guest_of_honor=G["guest_of_honor"], event_name=E["event_name"],
                    contributor_count="0", total_amount="0", gift_name="",
                    gift_link="", gift_price="", summary_rows=GIFT_ROWS,
                    grand_collected="82,000", grand_paid="70,000", grand_remaining="12,000",
                    link_event=True, gift_remaining="0"))
    for i, body in enumerate(HTML_BODIES):
        yield f"text_body_to_html_{i}", "text_body_to_html", (body,), {}


def build_merge_snapshot(module=None) -> dict[str, str]:
    """The new thank-you and gift-report output, pinned by
    tests/golden/i18n_merge_snapshot.json. Same key scheme as build_snapshot."""
    module = module or i18n
    out: dict[str, str] = {}
    for label, fn_name, args, kw in _merge_calls():
        key = f"{label}({args!r},{sorted(kw.items())!r})"
        fn = getattr(module, fn_name, None)
        try:
            out[key] = "!!MISSING" if fn is None else repr(fn(*args, **kw))
        except Exception as exc:
            out[key] = f"!!RAISED {type(exc).__name__}: {exc}"
    for name in MERGE_CONSTANTS:
        out[f"CONST:{name}"] = repr(getattr(module, name, "!!MISSING"))
    return out

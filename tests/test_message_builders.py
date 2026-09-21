"""Characterization tests: lock in what the message builders do TODAY.

These are not a specification. They record current behavior so that the staged
extraction in docs/agentic/ARCHITECTURE.md can move this code without changing
it. If one of these fails during a refactor, the refactor changed behavior.

They now import rsvp.i18n directly. Before phase 1 they had to import rsvp_app,
which pulls in tkinter and outlook_com, so they SKIPPED on any machine without
a display and a signed-in Outlook - a test that skips is a test that is not
protecting anything. This file is the concrete payoff of the extraction.
"""
import pytest

import rsvp.i18n as i18n

LANGS = ["en", "ja", "vi"]


@pytest.fixture(scope="module")
def app():
    """Kept so the test bodies below are unchanged from before the extraction.

    Same names, same assertions, now resolved from the extracted package
    instead of the monolith - which is exactly the claim phase 1 has to make.
    """
    return i18n


def test_every_language_produces_a_distinct_subject(app):
    subs = {lang: app.build_subject(lang, "E01", "Year End Party") for lang in LANGS}
    assert len(set(subs.values())) == len(LANGS), (
        f"two languages produced the same subject - a .get(lang, T['en']) "
        f"fallback is silently serving English: {subs}"
    )
    for lang, s in subs.items():
        assert "E01" in s and "Year End Party" in s, (lang, s)


def test_update_subject_differs_from_invite_subject(app):
    for lang in LANGS:
        invite = app.build_subject(lang, "E01", "Party", is_update=False)
        update = app.build_subject(lang, "E01", "Party", is_update=True)
        assert invite != update, f"{lang}: update invite is indistinguishable"


def test_bilingual_carries_both_languages(app):
    both = app.build_subject("bilingual", "E01", "Party")
    assert app.build_subject("en", "E01", "Party") in both
    assert app.build_subject("ja", "E01", "Party") in both


def test_unknown_language_falls_back_to_english(app):
    # Documenting the fallback, not endorsing it: this is exactly why
    # scripts/check_i18n_matrix.py exists.
    assert app.build_greeting("de") == app.build_greeting("en")


@pytest.mark.parametrize("lang", LANGS)
def test_fixed_block_mentions_every_event_field(app, lang):
    block = app.build_fixed_block(
        lang, "Year End Party", "2026-12-20 18:00", "Hall A", "2026-12-10", "3000 JPY")
    for field in ("Year End Party", "Hall A", "3000 JPY"):
        assert field in block, f"{lang}: {field!r} missing from the fixed block"


class TestParseAmountFromText:
    """parse_amount_from_text feeds the gift/attendance money totals.

    Still in rsvp_app.py: it is money/domain code, not i18n, so phase 1
    deliberately left it behind. Phase 2 extracts rsvp/domain/ and it moves
    there. Until then these reach it through the `monolith` fixture, which
    resolves on Windows CI and skips elsewhere - so the known defect below
    stays covered where the suite can actually reach it, rather than being
    silently dropped for the duration of the refactor.
    """

    @pytest.mark.parametrize("text,expected", [
        ("3,000 JPY / person", 3000.0),
        ("3000", 3000.0),
        ("", 0.0),
        (None, 0.0),
        ("no digits here", 0.0),
    ])
    def test_documented_cases(self, monolith, text, expected):
        assert monolith.parse_amount_from_text(text) == expected

    def test_grabs_the_first_number_not_the_amount(self, monolith):
        """KNOWN DEFECT, characterized so a refactor cannot hide it.

        The regex takes the first numeric run in the string, so a year or any
        other leading number wins over the real amount. Tracked as U3 in
        docs/agentic/ARCHITECTURE.md; fixing it is a behavior change that
        belongs to the domain-extraction phase, not to a scaffolding pass.
        """
        assert monolith.parse_amount_from_text("2026 year-end party, 3000 JPY") == 2026.0

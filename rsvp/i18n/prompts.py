"""Default Copilot translation prompts.

User overrides live in the JSON files owned by rsvp/storage/settings.py; these are the
fallbacks used until one is saved.

Moved verbatim out of rsvp_app.py (lines 85-154) by the phase 1 extraction in
docs/agentic/ARCHITECTURE.md. The code is unchanged; only its location is.

Exception: DEFAULT_PROMPT_BILINGUAL was later rewritten to translate only the
organizer's note, in any language, into Japanese and English, since the rest of
a bilingual email is already written in both. tests/i18n_snapshot.py lists it as
retired from the phase-1 golden file; tests/test_bilingual.py pins its reply format.
"""

# ══════════════════════════════════════════════════════════════════════════
# Copilot translation prompt — SYSTEM DEFAULTS (used unless user customizes
# and saves an override via Tab 3's "Customize Copilot prompt" box).
# Both defaults explicitly ask for emoji icons so translated emails are more
# visually scannable (⏰ time, 📍 location, 💰 cost, 📋 deadline/instructions,
# 👥 attendees). Use "[TARGET_LANGUAGE]" as a placeholder in the SINGLE
# template — it gets swapped for the real language name before copying.
# ══════════════════════════════════════════════════════════════════════════
DEFAULT_PROMPT_SINGLE = (
    "Translate the following internal event email into [TARGET_LANGUAGE]. "
    "Reply with ONLY the final, ready-to-send translated email as continuous text — "
    "do NOT add section labels, headers, quotation marks, or any explanation/commentary "
    "about your process. Keep the same paragraph order as the original. Do not merge, "
    "duplicate, summarize, or repeat any sentence — translate everything exactly once. "
    "Keep a polite, professional tone suitable for a workplace event email, and keep "
    "names, dates, and numbers exactly as written.\n\n"
    "Also make the email more visually scannable by adding ONE relevant emoji icon "
    "inline right before these types of information (do not add a legend/key — just "
    "place the icon naturally in the sentence):\n"
    "  ⏰ before any time / date / deadline\n"
    "  📍 before any location / venue\n"
    "  💰 before any cost / budget / price\n"
    "  📋 before instructions or action items (e.g. what to bring, what to do)\n"
    "  👥 before attendee / participant / headcount information\n"
    "Use each icon at most once per relevant sentence — do not overuse them, and do "
    "not add icons to the Yes/No/Maybe voting-button instructions.\n\n"
    "FORMATTING — plain text only, this matters: do NOT use Markdown syntax anywhere "
    "(no **bold**, no '-' or '*' bullet lists, no '#' headers, no code blocks/backticks). "
    "Write each paragraph as continuous flowing text — do NOT break a single sentence or "
    "short phrase onto its own separate line; keep each paragraph together and only start "
    "a new line at real paragraph boundaries, separated by ONE blank line. (Copying from "
    "some AI chat interfaces can silently drop the line breaks between many short lines, "
    "gluing words together with no space — writing in fewer, longer flowing paragraphs "
    "avoids that problem.) For the Yes/No/Maybe list, keep each bullet ('•' character) on "
    "its own line as in the source."
)

# The bilingual prompt sends the whole email content except the greeting:
# the organizer's note, in any language, and the fixed part (event details and
# voting instructions) as the fixed box holds it, in Japanese and English and
# possibly edited. build_bilingual_prompt() puts both in marked blocks near
# the top. The reply - a Japanese and an English version of each part, under
# four markers - is what parse_bilingual_reply() in cleanup.py reads.
NO_NOTE = "(none)"
DEFAULT_PROMPT_BILINGUAL = (
    "Translate the email content below into BOTH Japanese AND English. It has two parts: "
    "the organizer's NOTE, which may be written in ANY language, and the EVENT DETAILS with "
    "the voting instructions, given in Japanese and in English (if the two versions differ, "
    "keep everything from both). If a part is already in Japanese or English, keep its "
    "meaning and polish it in that language.\n\n"
    "===== NOTE =====\n"
    "[NOTE]\n"
    "===== EVENT DETAILS =====\n"
    "[DETAILS]\n"
    "===== END =====\n\n"
    "Do NOT add a greeting or a sign-off: the email adds them.\n\n"
    "Reply in EXACTLY this format and write nothing else (no title, no explanation). Keep "
    "the four marker lines exactly as they are, and replace each line in angle brackets with "
    "your translation (if the note is " + NO_NOTE + ", write " + NO_NOTE + " under both NOTE "
    "markers):\n"
    "[JA NOTE]\n"
    "<the note in Japanese>\n"
    "[JA DETAILS]\n"
    "<the event details and voting instructions in Japanese>\n"
    "[EN NOTE]\n"
    "<the note in English>\n"
    "[EN DETAILS]\n"
    "<the event details and voting instructions in English>\n\n"
    "Keep a polite, professional tone suitable for a workplace event email (Japanese: polite "
    "business style). Keep names, dates, numbers and amounts exactly as written. Do not "
    "summarize, merge or repeat sentences: translate everything exactly once.\n\n"
    "Make both versions easy to scan by adding ONE relevant emoji icon inline right before "
    "these types of information (no legend, just place the icon naturally in the sentence):\n"
    "  ⏰ before any time / date / deadline\n"
    "  📍 before any location / venue\n"
    "  💰 before any cost / budget / price\n"
    "  📋 before instructions or action items\n"
    "  👥 before attendee / participant / headcount information\n"
    "Use each icon at most once per sentence, and add none to the Yes/No/Maybe voting-button "
    "list.\n\n"
    "FORMATTING: plain text only. Do NOT use Markdown anywhere (no **bold**, no '-' or '*' "
    "bullet lists, no '#' headers, no code blocks). Write each paragraph as continuous text "
    "and separate paragraphs with ONE blank line. For the Yes/No/Maybe list, keep each bullet "
    "('•' character) on its own line as in the source."
)


def build_bilingual_prompt(note, details):
    """DEFAULT_PROMPT_BILINGUAL with the note and the fixed part in their
    marked blocks; an empty note reads NO_NOTE. The template's placeholders
    are filled before the content is, so content that says "[NOTE]" or
    "[DETAILS]" is kept as written."""
    head, rest = DEFAULT_PROMPT_BILINGUAL.split("[NOTE]", 1)
    middle, tail = rest.split("[DETAILS]", 1)
    return head + (note.strip() or NO_NOTE) + middle + details.strip() + tail

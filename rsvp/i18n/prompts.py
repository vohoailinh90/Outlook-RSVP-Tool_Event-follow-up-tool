"""Default Copilot translation prompts.

User overrides live in the JSON files owned by history.py; these are the
fallbacks used until one is saved.

Moved verbatim out of rsvp_app.py (lines 85-154) by the phase 1 extraction in
docs/agentic/ARCHITECTURE.md. The code is unchanged; only its location is.
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

DEFAULT_PROMPT_BILINGUAL = (
    "Translate the following internal event email into BOTH Japanese AND English, "
    "no matter what language the source text below is written in. "
    "Reply with ONE bilingual document structured like this: the complete JAPANESE "
    "version FIRST — starting with '[English below]' as its very first line — "
    "followed by a blank line and a divider line, then the complete ENGLISH version "
    "SECOND. Keep the two language versions COMPLETELY SEPARATE — do not interleave "
    "or mix sentences between languages; each version must be a full, independently "
    "readable translation covering everything in the source. Do not add any other "
    "section labels, headers, or commentary about your process. Do not merge, "
    "duplicate, summarize, or repeat any sentence within a language version. Keep a "
    "polite, professional tone suitable for a workplace event email, and keep names, "
    "dates, and numbers exactly as written.\n\n"
    "Also make BOTH versions more visually scannable by adding ONE relevant emoji icon "
    "inline right before these types of information (do not add a legend/key — just "
    "place the icon naturally in the sentence, in both language versions):\n"
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

"""Repairs for text pasted back from Copilot's web UI.

Pure string handling - no Tkinter, no Outlook, no I/O.

Moved verbatim out of rsvp_app.py (lines 156-271) by the phase 1 extraction in
docs/agentic/ARCHITECTURE.md. The code is unchanged; only its location is.
"""

import re

# ══════════════════════════════════════════════════════════════════════════
# Cleanup helper for a known Copilot-copy quirk: pasting a reply copied from
# Copilot's web UI sometimes silently drops line breaks between short lines
# (NOT replaced by a space — just gone), so adjacent lines' words run
# together with zero separator, e.g. "Hello everyone," + "Please respond..."
# becomes "Helloeveryone,Pleaserespond...". This happens in Copilot's own
# copy mechanism (outside this tool's control), so it can't be fixed at the
# source — this is a best-effort, SAFE repair applied after pasting: it only
# touches spots anchored to something recognizable (emoji icons, bullet
# markers, Japanese full stops) and does not attempt to guess-reconstruct
# spaces lost strictly inside a plain sentence with no such anchor nearby.
#
# ALSO observed: Copilot's reply can contain the WHOLE email duplicated
# WITHIN THE SAME RESPONSE — an initial malformed/glued draft immediately
# followed by a self-corrected, properly-spaced version of the identical
# content. This is NOT the user pasting twice; it's literally what Copilot
# returned in one go. The two copies differ in SPACING (one glued, one not),
# so a plain exact-text duplicate check misses it — dedupe below compares a
# WHITESPACE-STRIPPED version of the text so spacing differences don't hide
# the match, then removes everything before the LAST copy (self-correction
# is normally the cleaner one).
# ══════════════════════════════════════════════════════════════════════════
EMOJI_PATTERN = re.compile(
    "["
    "\u2300-\u23FF"          # misc technical (⏰ alarm clock, etc.)
    "\u2600-\u27BF"          # misc symbols & dingbats (✅❌☀️ etc.)
    "\U0001F300-\U0001F5FF"  # misc symbols & pictographs (📍💰📋📅 etc.)
    "\U0001F600-\U0001F64F"  # emoticons
    "\U0001F680-\U0001F6FF"  # transport & map symbols
    "\U0001F900-\U0001F9FF"  # supplemental symbols & pictographs (👥 etc.)
    "\U0001FA70-\U0001FAFF"  # symbols & pictographs extended-A
    "]"
)


def _normalized_to_original_index(text, normalized_index):
    """Map an index into the whitespace-stripped version of `text` back to
    the corresponding index in the ORIGINAL `text` (helper for dedupe below,
    since detection runs on a whitespace-stripped copy but we need to cut
    the ORIGINAL text at the right spot)."""
    count = 0
    for i, ch in enumerate(text):
        if not ch.isspace():
            if count == normalized_index:
                return i
            count += 1
    return len(text)


def dedupe_pasted_translation(text, min_repeat_len=60):
    """If `text` contains the WHOLE email duplicated within itself, keep only
    the LAST copy and discard everything before it. Detection is done on a
    whitespace-STRIPPED copy of the text (see module comment above for why:
    the two copies typically differ only in spacing), then the cut point is
    mapped back to the correct index in the original text.
    Returns `text` unchanged if no duplication is detected."""
    if not text:
        return text
    normalized = re.sub(r"\s+", "", text)
    if len(normalized) < min_repeat_len * 2:
        return text
    head = normalized[:min_repeat_len]
    second_pos = normalized.find(head, min_repeat_len)
    if second_pos == -1:
        return text
    cut_at = _normalized_to_original_index(text, second_pos)
    return text[cut_at:].strip()


def detect_possible_duplicate_paste(text, min_repeat_len=60):
    """True if dedupe_pasted_translation() would meaningfully shorten `text`
    — i.e. the email content appears to be duplicated within itself (either
    from Copilot's own reply containing a malformed-then-corrected repeat, or
    from pasting a new result without clearing the box first)."""
    if not text:
        return False
    deduped = dedupe_pasted_translation(text, min_repeat_len=min_repeat_len)
    return (len(text) - len(deduped)) > 20


def cleanup_pasted_translation(text):
    """Best-effort repair for lost line breaks after pasting a Copilot reply
    (see module-level comment above for why this happens). Fixes, in order:
      1. Strip stray literal '**' markdown bold markers that leaked through
         as plain text.
      2. Ensure a space exists immediately before AND after every emoji icon
         — the icon is usually exactly where a lost line break used to be,
         so this recovers the most visible/common cases.
      3. Ensure a newline appears before each '•' bullet character, so the
         Yes/No/Maybe list renders as a proper list again.
      4. Insert a paragraph break after each Japanese full stop '。' that is
         glued directly to the next character with no space/newline.
    Does NOT deduplicate (see dedupe_pasted_translation() for that — the
    caller decides whether to dedupe first, since it's a more impactful
    change that's worth confirming with the user).
    Does NOT attempt to reconstruct spaces lost strictly inside a plain
    sentence with no icon/bullet/punctuation anchor nearby — please review
    the result before saving/sending."""
    if not text:
        return text

    text = text.replace("**", "")
    text = text.replace("∗∗∗", "").replace("***", "")  # stray emphasis markers

    text = EMOJI_PATTERN.sub(lambda m: f" {m.group(0)} ", text)
    text = re.sub(r"[ \t]{2,}", " ", text)     # collapse doubled spaces just created
    text = re.sub(r"[ \t]+\n", "\n", text)     # trim trailing spaces before a newline
    text = re.sub(r"\n[ \t]+", "\n", text)     # trim leading spaces after a newline

    text = re.sub(r"(?<!\n)[ \t]*•", "\n•", text)

    text = re.sub(r"。(?=[^\s\n])", "。\n\n", text)

    text = re.sub(r"\n{3,}", "\n\n", text)     # collapse 3+ blank lines to 1

    return text.strip()


# Copilot's answer to DEFAULT_PROMPT_BILINGUAL: the note and the event
# details, each in Japanese and in English, under the markers [JA NOTE],
# [JA DETAILS], [EN NOTE] and [EN DETAILS]. The '#', bold, 【】 and colon
# Copilot sometimes puts around a marker belong to it. Markers are looked for
# at the start of a line first, so text that mentions "[EN NOTE]" mid-line is
# not cut there; only if that finds no complete answer are they looked for
# anywhere, since the copy quirk above can glue them to the words around them.
_MARKER = (r"(?:[#>]+[ \t]*)?(?:\*\*|__)?[\[【]\s*"
           r"(?:(?P<ja>JA|JP|JAPANESE|日本語)|(?P<en>EN|ENGLISH|英語))"
           r"[ \t_\-]*(?:(?P<note>NOTE)|(?P<details>DETAILS?))"
           r"\s*[\]】](?:\*\*|__)?[ \t]*[:：]?")
_LINE_START_MARKER = re.compile(r"^[ \t]*" + _MARKER, re.IGNORECASE | re.MULTILINE)
_ANYWHERE_MARKER = re.compile(_MARKER, re.IGNORECASE)
# Divider lines and code fences Copilot draws around a part.
_EDGE = r"(?:[―—–\-=_─━]{3,}|`{3}\w*)"
_EDGE_LINES = re.compile(r"\A(?:" + _EDGE + r"\s*)+|(?:\s*" + _EDGE + r")+\Z")
_NO_NOTE = re.compile(r"\(?\s*(?:none|なし)\s*\)?", re.IGNORECASE)


def _reply_parts(text, marker_pattern):
    markers = list(marker_pattern.finditer(text))
    found = {}
    for i, marker in enumerate(markers):
        end = markers[i + 1].start() if i + 1 < len(markers) else len(text)
        part = _EDGE_LINES.sub("", text[marker.end():end].strip()).strip()
        if part.startswith("<") and part.endswith(">"):
            part = part[1:-1].strip()   # the prompt's <...> placeholder brackets, kept
        is_note = bool(marker.group("note"))
        if is_note and _NO_NOTE.fullmatch(part):
            part = ""
        key = ("ja" if marker.group("ja") else "en", "note" if is_note else "details")
        if part or is_note:
            found[key] = part
    if not found.get(("ja", "details")) or not found.get(("en", "details")):
        return None
    return (found.get(("ja", "note"), ""), found[("ja", "details")],
            found.get(("en", "note"), ""), found[("en", "details")])


def parse_bilingual_reply(text):
    """(ja_note, ja_details, en_note, en_details) from Copilot's answer to
    DEFAULT_PROMPT_BILINGUAL, or None unless both a Japanese and an English
    details part with text in them are found. A missing note part, or one
    that says "(none)", is "". Anything before the first marker ("Here is the
    translation:") is dropped, and so are divider lines, code fences and the
    prompt's <...> brackets around a part. A marker found twice keeps its
    later part, as dedupe_pasted_translation() keeps the later copy. Text
    Copilot adds after the last part stays in it: the app shows the result
    for checking before anything is sent."""
    text = text or ""
    return _reply_parts(text, _LINE_START_MARKER) or _reply_parts(text, _ANYWHERE_MARKER)

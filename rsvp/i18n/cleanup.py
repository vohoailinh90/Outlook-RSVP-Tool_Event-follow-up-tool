"""Repairs for text pasted back from Copilot's web UI.

Pure string handling - no Tkinter, no Outlook, no I/O.

Moved verbatim out of rsvp_app.py (lines 156-271) by the phase 1 extraction in
docs/agentic/ARCHITECTURE.md. The code is unchanged; only its location is.
"""

import re
import unicodedata
from collections import Counter

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
# Copilot sometimes puts around a marker belong to it. Markers at the start of
# a line are taken first; one found mid-line counts only when no line-start
# marker of the same part comes after it, and the marker before it opened a
# different part. The copy quirk above can glue a marker - any one of them,
# also in a corrected copy that follows a draft - to the words before it, while
# text that merely mentions "[EN NOTE]" mid-line comes before the real marker
# on its own line, or sits inside the [EN NOTE] part itself, and is not cut.
_MARKER = (r"(?:[#>]+[ \t]*)?(?:\*\*|__)?[\[【]\s*"
           r"(?:(?P<ja>JA|JP|JAPANESE|日本語)|(?P<en>EN|ENGLISH|英語))"
           r"[ \t_\-]*(?:(?P<note>NOTE)|(?P<details>DETAILS?))"
           r"\s*[\]】](?:\*\*|__)?[ \t]*[:：]?")
_LINE_START_MARKER = re.compile(r"^[ \t]*" + _MARKER, re.IGNORECASE | re.MULTILINE)
_ANYWHERE_MARKER = re.compile(_MARKER, re.IGNORECASE)
# Divider lines and code fences Copilot draws around a part.
_EDGE = r"(?:[―—–\-=_─━]{3,}|`{3}\w*)"
_EDGE_LINES = re.compile(r"\A(?:" + _EDGE + r"\s*)+|(?:\s*" + _EDGE + r")+\Z")
# What Copilot writes for an empty note, as asked or in its own words; and the
# prompt's own placeholders, which a reply that only echoes the template holds.
_NO_NOTE = {"none", "なし", "無し", "特になし", "n/a", "na", "-", "—", "ー"}
_PLACEHOLDER = re.compile(
    r"the (?:note|event details and voting instructions) in (?:japanese|english)")


def _is_no_note(part):
    text = unicodedata.normalize("NFKC", part).strip().strip("()[]「」『』.。!！ ").lower()
    return text in _NO_NOTE


def _marker_key(marker):
    return ("ja" if marker.group("ja") else "en", "note" if marker.group("note") else "details")


def _reply_markers(text):
    at_line_start = list(_LINE_START_MARKER.finditer(text))
    last_at_line_start = {_marker_key(m): m.start() for m in at_line_start}
    line_start_ends = {m.end() for m in at_line_start}   # the same markers, found again
    glued = [m for m in _ANYWHERE_MARKER.finditer(text)
             if m.end() not in line_start_ends
             and m.start() > last_at_line_start.get(_marker_key(m), -1)]
    glued_ids = {id(m) for m in glued}
    markers = []
    for m in sorted(at_line_start + glued, key=lambda m: m.start()):
        if id(m) in glued_ids and markers and _marker_key(markers[-1]) == _marker_key(m):
            continue    # its own part's text, mentioning its marker
        markers.append(m)
    return markers


def parse_bilingual_reply(text):
    """(ja_note, ja_details, en_note, en_details) from Copilot's answer to
    DEFAULT_PROMPT_BILINGUAL, or None unless both a Japanese and an English
    details part with text in them are found. A missing note part, or one
    that says "(none)" in any of the ways Copilot writes it, is "". A part
    that is still the prompt's placeholder counts as missing. Anything before
    the first marker ("Here is the translation:") is dropped, and so are
    divider lines, code fences and the prompt's <...> brackets around a
    part (but not the brackets of a tag like <b>...</b> in the text). A marker found twice keeps its later part, as
    dedupe_pasted_translation() keeps the later copy. Text Copilot adds after
    the last part stays in it: the app shows the result for checking before
    anything is sent."""
    text = text or ""
    markers = _reply_markers(text)
    found = {}
    for i, marker in enumerate(markers):
        end = markers[i + 1].start() if i + 1 < len(markers) else len(text)
        part = _EDGE_LINES.sub("", text[marker.end():end].strip()).strip()
        if (part.startswith("<") and part.endswith(">")
                and not any(c in part[1:-1] for c in "<>")):
            part = part[1:-1].strip()   # the prompt's <...> placeholder brackets, kept
        if _PLACEHOLDER.fullmatch(part.lower()):
            continue
        key = _marker_key(marker)
        if key[1] == "note" and _is_no_note(part):
            part = ""
        if part or key[1] == "note":
            found[key] = part
    if not found.get(("ja", "details")) or not found.get(("en", "details")):
        return None
    return (found.get(("ja", "note"), ""), found[("ja", "details")],
            found.get(("en", "note"), ""), found[("en", "details")])


# Numbers and the voting buttons are what a translation must never change: a
# date, an amount or a deadline reaches colleagues as written, and the buttons
# Outlook shows are always Yes, No and Maybe.
_NUMBER = re.compile(r"\d+(?:,\d{3})*")
_ON_THE_HOUR = re.compile(r"(?<!\d)(\d{1,2}):00(?!\d)")
_BUTTONS = ("Yes", "No", "Maybe")


def _numbers(text):
    """How often each number occurs. The ':00' of a time on the hour is not
    counted - "18:00" becomes "18時" in a good Japanese translation - but
    every other zero is, a budget of 0 included."""
    text = _ON_THE_HOUR.sub(r"\1", unicodedata.normalize("NFKC", text))
    return Counter(int(n.replace(",", "")) for n in _NUMBER.findall(text))


def _button_names(text):
    return [b for b in _BUTTONS if re.search(rf"(?<![A-Za-z]){b}(?![A-Za-z])", text)]


def translation_gaps(source, translated, template=None, other=None):
    """What `translated` lacks that `source` has: numbers, compared by value
    and by how often they occur (so 07 matches 7, 3,000 matches 3000, a date
    may change its order, and of two identical dates neither may go), and the
    voting-button names Yes, No and Maybe. [] when nothing is missing.

    other=(other_source, other_template) adds the numbers typed into the
    other language's source beyond its built-in template - less those typed
    into this source beyond `template` - so a deadline added by hand to one
    half of the fixed part must reach both translations, also when it repeats
    the event date, and one added to both halves counts once. The halves' own
    differences in wording ("3 BUTTONS" in English only) do not count. A
    deterministic check, not a judgement: a correct translation that spells a
    number out ("three") is reported too, for a person to look at."""
    expected = _numbers(source)
    if other:
        other_source, other_template = other
        added_there = _numbers(other_source) - _numbers(other_template)
        added_here = expected - _numbers(template) if template is not None else Counter()
        expected = expected + (added_there - added_here)
    missing = expected - _numbers(translated)
    gaps = [str(n) if k == 1 else f"{n} (×{k})" for n, k in sorted(missing.items())]
    return gaps + [b for b in _button_names(source) if b not in _button_names(translated)]

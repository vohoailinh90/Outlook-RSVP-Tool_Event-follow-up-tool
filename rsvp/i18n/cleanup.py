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
# a line always count. One found mid-line - the copy quirk above can glue a
# marker to the words before it, also in a corrected copy that follows a
# draft - counts only when it opens the part that comes next in the reply's
# order (a note may be skipped, and after the last part a new copy may start),
# and no line-start marker of the same part follows it. Text that merely
# mentions a marker mid-line, inside any part, is not cut there.
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


# The parts that may follow each part, in the order the prompt asks for.
_NEXT_PARTS = {
    ("ja", "note"): {("ja", "details")},
    ("ja", "details"): {("en", "note"), ("en", "details")},
    ("en", "note"): {("en", "details")},
    ("en", "details"): {("ja", "note"), ("ja", "details")},   # a corrected copy
}


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
        if (id(m) in glued_ids and markers
                and _marker_key(m) not in _NEXT_PARTS[_marker_key(markers[-1])]):
            continue    # text of the open part that mentions a marker
        markers.append(m)
    return markers


def _unwrapped(part):
    """The part without one pair of <...> around it, or None if it has none
    (the brackets of a tag like <b>...</b> inside do not count)."""
    if part.startswith("<") and part.endswith(">") and not any(c in part[1:-1] for c in "<>"):
        return part[1:-1].strip()
    return None


def parse_bilingual_reply(text):
    """(ja_note, ja_details, en_note, en_details) from Copilot's answer to
    DEFAULT_PROMPT_BILINGUAL, or None unless both a Japanese and an English
    details part with text in them are found. A missing note part, or one
    that says "(none)" in any of the ways Copilot writes it, is "". A part
    that is still the prompt's placeholder counts as missing. Anything before
    the first marker ("Here is the translation:") is dropped, and so are
    divider lines and code fences around a part. The prompt's <...> brackets
    are dropped only when every part kept them - one part written as <TBD>
    is the text itself. A part named a second time starts a later copy, which
    replaces the whole earlier answer - parts it leaves out included - as
    dedupe_pasted_translation() keeps the later copy. Text Copilot adds after
    the last part stays in it: the app shows the result for checking before
    anything is sent."""
    text = text or ""
    markers = _reply_markers(text)
    raw = []
    for i, marker in enumerate(markers):
        end = markers[i + 1].start() if i + 1 < len(markers) else len(text)
        part = _EDGE_LINES.sub("", text[marker.end():end].strip()).strip()
        if _PLACEHOLDER.fullmatch((_unwrapped(part) or part).lower()):
            continue    # the template, echoed back unfilled
        raw.append((_marker_key(marker), part))
    # "(none)" is written without brackets, as the prompt asks: it does not
    # tell whether Copilot kept the template's brackets around its text.
    texts = [part for key, part in raw
             if part and not (key[1] == "note" and _is_no_note(part))]
    if texts and all(_unwrapped(part) is not None for part in texts):
        raw = [(key, _unwrapped(part) if _unwrapped(part) is not None else part)
               for key, part in raw]
    found = {}
    for key, part in raw:
        if key in found:
            found = {}  # a part named again starts a later copy, which replaces all of the earlier one
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
# A sign belongs to a number when it is attached to it and does not follow a
# letter, digit or another sign: "-500", "¥-500" and the "+81" of a phone
# number keep theirs, while the hyphens of "2026-12-20" or "10-12" are not
# signs.
_NUMBER = re.compile(r"((?<![\w+-])[+-])?(\d+(?:,\d{3})*)")
# Times are compared as times of day on a 24-hour clock, so "18:00" matches
# "18時", "6:00 PM", "6 PM" or "午後6時", and "10:30" matches "10時半" -
# while "18:00" turned into "18:30", or PM into AM, is caught.
_MERIDIEM = r"(?:\s*([AaPp])\.?\s*[Mm]\.?(?![A-Za-z]))"
_CLOCK_TIME = re.compile(r"(午前|午後)?\s*(?<!\d)(\d{1,2}):(\d{2})(?!\d)" + _MERIDIEM + "?")
_JA_TIME = re.compile(r"(午前|午後)?(?<!\d)(\d{1,2})時(?!間)(?:(\d{1,2})分|(半))?")   # 3時間 is a duration
_HOUR_MERIDIEM = re.compile(r"(?<![\d:])(\d{1,2})" + _MERIDIEM)
_BUTTONS = ("Yes", "No", "Maybe")


def _numbers(text):
    """How often each number occurs, with its sign, and each time of day as
    "H:MM". Every zero counts, a budget of 0 included."""
    text = unicodedata.normalize("NFKC", text).replace("\u2212", "-")   # − MINUS SIGN
    times = Counter()

    def take(hours, minutes, half_of_day=None):
        hours = int(hours)
        if half_of_day in ("p", "P", "午後"):
            hours = hours % 12 + 12
        elif half_of_day in ("a", "A", "午前"):
            hours = hours % 12
        times[f"{hours}:{int(minutes):02d}"] += 1
        return " "
    text = _CLOCK_TIME.sub(lambda m: take(m[2], m[3], m[4] or m[1]), text)
    text = _JA_TIME.sub(lambda m: take(m[2], 30 if m[4] else (m[3] or 0), m[1]), text)
    text = _HOUR_MERIDIEM.sub(lambda m: take(m[1], 0, m[2]), text)
    numbers = Counter(f"{sign}{int(digits.replace(',', ''))}"
                      for sign, digits in _NUMBER.findall(text))
    return numbers + times


def _button_names(text):
    """How often each voting-button name occurs: the built-in fixed part
    names each twice, in the instruction and in the bullet that explains it."""
    return Counter({b: len(re.findall(rf"(?<![A-Za-z]){b}(?![A-Za-z])", text)) for b in _BUTTONS})


def _number_order(n):
    return (0, int(n), n) if ":" not in n else (1, 0, n)


def _counted(text, count):
    """count(text), or for a list of versions of one text, each value as often
    as the version that has it most."""
    versions = text if isinstance(text, (list, tuple)) else [text]
    most = Counter()
    for version in versions:
        most |= count(version)
    return most


def _listed(counter, key=None):
    return [n if k == 1 else f"{n} (×{k})" for n, k in sorted(counter.items(), key=key)]


def translation_extras(texts, translated):
    """Numbers and times in `translated` beyond those in the copied `texts` -
    a fact Copilot made up ("budget 500", or a second date), which
    translation_gaps(), looking only for what went missing, cannot see. Each
    of `texts` is a string, or a list of versions of one text (the Japanese
    and English halves of the fixed part), whose numbers count as often as
    the version that has them most - so a number only one half has
    ("3 BUTTONS" in English only) is not an invention."""
    allowed = Counter()
    for text in texts:
        allowed += _counted(text, _numbers)
    extra = _numbers(translated) - allowed
    return _listed(extra, key=lambda item: _number_order(item[0]))


def translation_gaps(source, translated, template=None, other=None):
    """What `translated` lacks that `source` has: numbers, compared by value
    and by how often they occur (so 07 matches 7, 3,000 matches 3000, a date
    may change its order, and of two identical dates neither may go), and the
    voting-button names Yes, No and Maybe, counted the same way. [] when
    nothing is missing.

    `source` may also be a list of versions of the same text - a note already
    in Japanese and English: each value is then expected as often as the
    version that has it most.

    other=(other_source, other_template) is the other language's version of
    the same fixed part, with `template` and other_template the two built-in
    texts. Each value is then expected as often as the version that has it
    most, after taking from the other version what only its template has
    ("3 BUTTONS" in English only): so a deadline typed into one half must
    reach both translations, also when it repeats the event date; one typed
    into both counts once; and a value an earlier translation already carried
    across is not expected twice. A deterministic check, not a judgement: a
    correct translation that spells a number out ("three") is reported too,
    for a person to look at."""
    expected = _counted(source, _numbers)
    buttons = _counted(source, _button_names)
    if other:
        other_source, other_template = other
        for counted, count in ((expected, _numbers), (buttons, _button_names)):
            only_there = count(other_template) - (count(template) if template is not None else Counter())
            counted |= count(other_source) - only_there
    missing = expected - _numbers(translated)
    lost_buttons = buttons - _button_names(translated)
    return (_listed(missing, key=lambda item: _number_order(item[0]))
            + _listed(lost_buttons, key=lambda item: _BUTTONS.index(item[0])))

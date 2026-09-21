"""Message building, language tables and paste cleanup.

Pure: imports nothing that needs a display, Outlook or openpyxl, which is
what makes it testable anywhere. scripts/check_layering.py enforces that.
"""

from .langs import (  # noqa: F401
    LANG_LABELS,
    LANG_LABEL_TO_CODE,
    TRANSLATE_TARGETS,
    BILINGUAL_SEPARATOR,
    NOT_TRANSLATED_FLAG,
    GREETING,
)
from .prompts import (  # noqa: F401
    DEFAULT_PROMPT_SINGLE,
    DEFAULT_PROMPT_BILINGUAL,
)
from .cleanup import (  # noqa: F401
    EMOJI_PATTERN,
    dedupe_pasted_translation,
    detect_possible_duplicate_paste,
    cleanup_pasted_translation,
)
from .messages import (  # noqa: F401
    build_greeting,
    build_subject,
    UPDATE_NOTICE,
    build_update_notice,
    build_fixed_block,
    REMINDER_LABELS,
    build_reminder_body,
    build_reminder_subject,
    CALENDAR_LABELS,
    build_calendar_body,
    build_thankyou_subject,
    THANKYOU_LABELS,
    build_thankyou_body,
    build_gift_subject,
    GIFT_LABELS,
    build_gift_fixed_block,
    build_gift_reminder_subject,
    build_gift_reminder_body,
    build_gift_report_subject,
    GIFT_REPORT_LABELS,
    build_gift_report_body,
    build_editable_block,
)

__all__ = [
    "LANG_LABELS",
    "LANG_LABEL_TO_CODE",
    "TRANSLATE_TARGETS",
    "BILINGUAL_SEPARATOR",
    "NOT_TRANSLATED_FLAG",
    "GREETING",
    "DEFAULT_PROMPT_SINGLE",
    "DEFAULT_PROMPT_BILINGUAL",
    "EMOJI_PATTERN",
    "dedupe_pasted_translation",
    "detect_possible_duplicate_paste",
    "cleanup_pasted_translation",
    "build_greeting",
    "build_subject",
    "UPDATE_NOTICE",
    "build_update_notice",
    "build_fixed_block",
    "REMINDER_LABELS",
    "build_reminder_body",
    "build_reminder_subject",
    "CALENDAR_LABELS",
    "build_calendar_body",
    "build_thankyou_subject",
    "THANKYOU_LABELS",
    "build_thankyou_body",
    "build_gift_subject",
    "GIFT_LABELS",
    "build_gift_fixed_block",
    "build_gift_reminder_subject",
    "build_gift_reminder_body",
    "build_gift_report_subject",
    "GIFT_REPORT_LABELS",
    "build_gift_report_body",
    "build_editable_block",
]

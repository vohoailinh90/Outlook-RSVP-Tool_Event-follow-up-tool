"""Language codes, labels and the shared separators.

No imports: this is the bottom of the i18n layer and everything else in it
depends on this module, never the other way round.

Moved verbatim out of rsvp_app.py (lines 58-83) by the phase 1 extraction in
docs/agentic/ARCHITECTURE.md. The code is unchanged; only its location is.
"""

# ══════════════════════════════════════════════════════════════════════════
# Language config
# ══════════════════════════════════════════════════════════════════════════
LANG_LABELS = {
    "en": "English",
    "ja": "Japanese",
    "vi": "Vietnamese",
    "bilingual": "Bilingual (English + Japanese)",
}
LANG_LABEL_TO_CODE = {v: k for k, v in LANG_LABELS.items()}
TRANSLATE_TARGETS = ["English", "Japanese", "Vietnamese", "Bilingual (Japanese + English)"]

BILINGUAL_SEPARATOR = "\n\n――――――――――――――――――――――――――――\n\n"

NOT_TRANSLATED_FLAG = {
    "en": "  [not yet translated — using original text]",
    "ja": "  【未翻訳 — 原文のまま表示しています】",
    "vi": "  [chưa dịch — đang hiển thị bản gốc]",
}


GREETING = {
    "en": "Hello everyone,",
    "ja": "皆様",
    "vi": "Chào các bạn,",
}

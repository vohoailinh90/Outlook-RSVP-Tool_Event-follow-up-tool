"""
User wording settings, stored as JSON next to the app:

    fixed_wording_overrides.json  the FIXED part of the invite, per language,
                                  as edited and saved on Tab 3
    prompt_overrides.json         the Copilot prompt template saved on Tab 3

The Excel event history this module used to own is gone: events live in
rsvp_data.db (rsvp/storage/db.py), and an old RSVP_History.xlsx is imported
once by rsvp/export/legacy_excel.py.
"""
import json
import os

FIXED_OVERRIDES_FILE_DEFAULT = "fixed_wording_overrides.json"
PROMPT_OVERRIDES_FILE_DEFAULT = "prompt_overrides.json"


def load_fixed_overrides(path=FIXED_OVERRIDES_FILE_DEFAULT):
    """Đọc các bản 'câu văn mặc định' đã được người dùng tự chỉnh & lưu lại cho FIXED
    part (áp dụng cho mọi sự kiện sau này, cho tới khi người dùng Reset hoặc sửa lại)."""
    if not os.path.exists(path):
        return {"en": "", "ja": "", "vi": ""}
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return {"en": data.get("en", ""), "ja": data.get("ja", ""), "vi": data.get("vi", "")}
    except Exception:
        return {"en": "", "ja": "", "vi": ""}


def save_fixed_overrides(overrides: dict, path=FIXED_OVERRIDES_FILE_DEFAULT):
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(overrides, fh, ensure_ascii=False, indent=2)
    return path


def load_prompt_overrides(path=PROMPT_OVERRIDES_FILE_DEFAULT):
    """Đọc custom prompt template được người dùng tùy chỉnh.
    Trả về dict {"single": "...", "bilingual": "..."} hoặc rỗng nếu chưa lưu."""
    if not os.path.exists(path):
        return {"single": "", "bilingual": ""}
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return {
            "single": data.get("single", ""),
            "bilingual": data.get("bilingual", "")
        }
    except Exception:
        return {"single": "", "bilingual": ""}


def save_prompt_overrides(overrides: dict, path=PROMPT_OVERRIDES_FILE_DEFAULT):
    """Lưu custom prompt template được người dùng tùy chỉnh."""
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(overrides, fh, ensure_ascii=False, indent=2)
    return path

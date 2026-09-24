"""Excel reading and writing.

The only layer allowed to depend on openpyxl.
"""
from .legacy_excel import migrate_from_excel_if_needed  # noqa: F401

__all__ = ["migrate_from_excel_if_needed"]

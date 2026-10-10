"""The design tokens in rsvp/ui/tokens.py.

They are the Automation UI Kit's oklch values converted to sRGB. The neutral
ones must come out as Tailwind's published neutral palette - an independent
check that the conversion is right, not a restatement of it - and every tone
used for text must stay readable on its own background (WCAG AA, 4.5:1), the
rule the kit tuned its status colours for.
"""
from __future__ import annotations

import pytest

from rsvp.ui import tokens


@pytest.mark.parametrize("name, tailwind", [
    ("foreground", "#0a0a0a"),        # neutral-950
    ("primary", "#171717"),           # neutral-900
    ("muted_foreground", "#737373"),  # neutral-500
    ("border", "#e5e5e5"),            # neutral-200
    ("muted", "#f5f5f5"),             # neutral-100
    ("sidebar", "#fafafa"),           # neutral-50
    ("background", "#ffffff"),
])
def test_neutrals_match_tailwind(name, tailwind):
    assert tokens.COLORS[name] == tailwind


def _luminance(color):
    def channel(c):
        return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (channel(c) for c in tokens.from_hex(color))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _contrast(a, b):
    hi, lo = sorted((_luminance(a), _luminance(b)), reverse=True)
    return (hi + 0.05) / (lo + 0.05)


@pytest.mark.parametrize("tone", sorted(tokens.TONES))
def test_tone_text_is_readable_on_its_background(tone):
    text, background, _stripe = tokens.TONES[tone]
    assert _contrast(text, background) >= 4.5


@pytest.mark.parametrize("name", ["foreground", "muted_foreground", "destructive", "success", "info"])
def test_text_tokens_are_readable_on_white(name):
    assert _contrast(tokens.COLORS[name], "#ffffff") >= 4.5


def test_warning_is_a_fill_not_a_text_colour():
    # The kit's reason for neutral warning text: this orange fails as text.
    assert _contrast(tokens.COLORS["warning"], "#ffffff") < 4.5
    assert _contrast(tokens.COLORS["warning"], "#ffffff") >= 3.0


def test_tint_blends_toward_white():
    assert tokens.tint("#000000", 0.5) == "#808080"
    assert tokens.tint("#123456", 1.0) == "#123456"
    assert tokens.tint("#123456", 0.0) == "#ffffff"

"""Design tokens of the Automation UI Kit, for a Tkinter app.

The kit (github.com/vohoailinh90/automation-ui-kit, src/index.css) defines its
light theme in oklch - shadcn/ui's neutral palette plus its own status
colours, each tuned for contrast there (see the comments on --success,
--warning and --destructive). Tk only understands sRGB, so the oklch values
are copied here verbatim and converted, rather than eyeballed into hex:
a token keeps meaning exactly what the kit means by it.

Only the light theme is carried over.
"""
from __future__ import annotations

import math

# name -> (L, C, h), copied from the kit's :root block.
OKLCH = {
    "background": (1.0, 0.0, 0.0),
    "foreground": (0.145, 0.0, 0.0),
    "card": (1.0, 0.0, 0.0),
    "primary": (0.205, 0.0, 0.0),
    "primary_foreground": (0.985, 0.0, 0.0),
    "secondary": (0.97, 0.0, 0.0),
    "muted": (0.97, 0.0, 0.0),
    "muted_foreground": (0.556, 0.0, 0.0),
    "accent": (0.97, 0.0, 0.0),
    "border": (0.922, 0.0, 0.0),
    "ring": (0.556, 0.0, 0.0),
    # Text colours (>= 4.5:1 on white) and fill colours (>= 3:1), kept apart
    # the way the kit keeps them apart.
    "destructive": (0.52, 0.22, 27.0),
    "success": (0.48, 0.11, 163.0),
    "success_fill": (0.6, 0.13, 163.0),
    "warning": (0.64, 0.17, 55.0),       # fill only: as text it fails contrast
    "info": (0.5, 0.17, 255.0),
    "sidebar": (0.985, 0.0, 0.0),
    "sidebar_accent": (0.97, 0.0, 0.0),
    "sidebar_border": (0.922, 0.0, 0.0),
}


def oklch_to_rgb(lightness, chroma, hue):
    """sRGB channels in 0..1 for an oklch colour, clipped to the gamut."""
    a = chroma * math.cos(math.radians(hue))
    b = chroma * math.sin(math.radians(hue))
    l_ = lightness + 0.3963377774 * a + 0.2158037573 * b
    m_ = lightness - 0.1055613458 * a - 0.0638541728 * b
    s_ = lightness - 0.0894841775 * a - 1.2914855480 * b
    l, m, s = l_ ** 3, m_ ** 3, s_ ** 3
    linear = (
        4.0767416621 * l - 3.3077115913 * m + 0.2309699292 * s,
        -1.2684380046 * l + 2.6097574011 * m - 0.3413193965 * s,
        -0.0041960863 * l - 0.7034186147 * m + 1.7076147010 * s,
    )

    def encode(x):
        x = min(max(x, 0.0), 1.0)
        return 12.92 * x if x <= 0.0031308 else 1.055 * x ** (1 / 2.4) - 0.055

    return tuple(encode(x) for x in linear)


def to_hex(rgb):
    return "#%02x%02x%02x" % tuple(round(c * 255) for c in rgb)


def from_hex(color):
    color = color.lstrip("#")
    return tuple(int(color[i:i + 2], 16) / 255 for i in (0, 2, 4))


def tint(color, alpha, over="#ffffff"):
    """`color` at `alpha` opacity over `over` - the kit's `bg-success/10`
    and friends, flattened, because Tk has no transparency."""
    top, bottom = from_hex(color), from_hex(over)
    return to_hex(tuple(alpha * t + (1 - alpha) * b for t, b in zip(top, bottom)))


COLORS = {name: to_hex(oklch_to_rgb(*value)) for name, value in OKLCH.items()}

# Derived the way the kit's components derive them.
COLORS.update({
    "primary_hover": tint(COLORS["primary"], 0.9),          # hover:bg-primary/90
    "destructive_hover": tint(COLORS["destructive"], 0.9),
    "selection": tint(COLORS["primary"], 0.08),             # a selected table row
    "success_soft": tint(COLORS["success"], 0.10),          # badge / banner backgrounds
    "warning_soft": tint(COLORS["warning"], 0.12),
    "info_soft": tint(COLORS["info"], 0.08),
    "destructive_soft": tint(COLORS["destructive"], 0.08),
    "success_wash": tint(COLORS["success"], 0.05),          # large text areas
    "warning_wash": tint(COLORS["warning"], 0.06),
})

# Tone -> (text, background, stripe) for banners and badges. Warning text is
# neutral on purpose: orange that reaches 4.5:1 on white turns brown, so the
# kit says "warning" with a coloured mark and neutral words.
TONES = {
    "neutral": (COLORS["foreground"], COLORS["muted"], COLORS["muted_foreground"]),
    "info": (COLORS["info"], COLORS["info_soft"], COLORS["info"]),
    "success": (COLORS["success"], COLORS["success_soft"], COLORS["success_fill"]),
    "warning": (COLORS["foreground"], COLORS["warning_soft"], COLORS["warning"]),
    "danger": (COLORS["destructive"], COLORS["destructive_soft"], COLORS["destructive"]),
}

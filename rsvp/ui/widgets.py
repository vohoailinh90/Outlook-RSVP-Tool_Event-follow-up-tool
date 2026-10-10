"""The kit's building blocks, rebuilt in Tkinter.

Card, StatusBadge-style Banner, KpiTile, Sidebar and PageHeader follow the
components of the same names in automation-ui-kit (src/components/ui/card.tsx,
src/components/dashboard/kpi-card.tsx, src/components/layout/*). Tk has no
rounded corners or shadows, so a card is a 1px hairline box - the rest
(spacing, type, colour rules) is the kit's.
"""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from .tokens import COLORS as C
from .tokens import TONES, tint


class WrapLabel(ttk.Label):
    """A label that re-wraps to the width its parent is given, so help text
    reflows when the window is resized instead of being cut off."""

    def __init__(self, master, pad=4, **kwargs):
        kwargs.setdefault("justify", "left")
        super().__init__(master, **kwargs)
        self._pad = pad
        master.bind("<Configure>", self._rewrap, add="+")

    def _rewrap(self, event):
        width = event.width - self._pad
        if width > 80:
            self.configure(wraplength=width)


def autohide(scrollbar):
    """A yscrollcommand/xscrollcommand that shows a grid-managed scrollbar only
    while there is something to scroll. For widgets that do not re-wrap when
    their width changes (Treeview, Listbox), so it cannot oscillate."""
    def set_(first, last):
        if float(first) <= 0.0 and float(last) >= 1.0:
            scrollbar.grid_remove()
        else:
            scrollbar.grid()
        scrollbar.set(first, last)
    return set_


def bordered(parent, color=None, **kwargs):
    """A frame drawn with a 1px border - the kit's `border` - and a white fill."""
    return tk.Frame(parent, bg=kwargs.pop("bg", C["card"]), bd=0, highlightthickness=1,
                    highlightbackground=color or C["border"], highlightcolor=color or C["border"],
                    **kwargs)


class Card(tk.Frame):
    """The kit's Card: header (title, description, actions on the right) and a
    body. Build content in `.body`; put header buttons in `.actions`."""

    def __init__(self, parent, title=None, description=None, padding=18):
        super().__init__(parent, bg=C["card"], bd=0, highlightthickness=1,
                         highlightbackground=C["border"], highlightcolor=C["border"])
        inner = ttk.Frame(self, style="Card.TFrame", padding=padding)
        inner.pack(fill="both", expand=True)
        self.actions = None
        self.description = None
        if title:
            header = ttk.Frame(inner, style="Card.TFrame")
            header.pack(fill="x", pady=(0, 14))
            self.actions = ttk.Frame(header, style="Card.TFrame")
            self.actions.pack(side="right", anchor="n", padx=(12, 0))
            heading = ttk.Frame(header, style="Card.TFrame")
            heading.pack(side="left", fill="x", expand=True)
            ttk.Label(heading, text=title, style="CardTitle.TLabel").pack(anchor="w")
            if description:
                self.description = WrapLabel(heading, text=description, style="Muted.TLabel")
                self.description.pack(anchor="w", fill="x", pady=(3, 0))
        self.body = ttk.Frame(inner, style="Card.TFrame")
        self.body.pack(fill="both", expand=True)


class Banner(ttk.Frame):
    """An alert line: tinted background, a coloured stripe, and text. Empty
    text hides it. Tones: neutral, info, success, warning, danger."""

    def __init__(self, parent, fonts, text="", tone="neutral"):
        super().__init__(parent)
        self._box = tk.Frame(self, bd=0, highlightthickness=1)
        self._stripe = tk.Frame(self._box, width=3)
        self._stripe.pack(side="left", fill="y")
        self._label = tk.Label(self._box, anchor="w", justify="left", font=fonts.body_medium,
                               padx=12, pady=9)
        self._label.pack(side="left", fill="both", expand=True)
        self._box.bind("<Configure>",
                       lambda e: self._label.configure(wraplength=max(e.width - 40, 120)))
        self.set(text, tone)

    def set(self, text, tone="neutral"):
        fg, bg, stripe = TONES[tone]
        self._box.configure(bg=bg, highlightbackground=tint(stripe, 0.35),
                            highlightcolor=tint(stripe, 0.35))
        self._stripe.configure(bg=stripe)
        self._label.configure(text=text, bg=bg, fg=fg)
        if text:
            if not self._box.winfo_manager():
                self._box.pack(fill="x")
        else:
            self._box.pack_forget()

    def cget_text(self):
        return self._label.cget("text")


class KpiTile(tk.Frame):
    """A KPI card: label, big number, optional coloured dot and footnote.
    `value` is a StringVar, so the number follows whatever the app sets - or
    a callable that builds an input in its place (a figure the user types)."""

    def __init__(self, parent, label, value, dot=None, value_style="KpiValue.TLabel", footnote=None):
        super().__init__(parent, bg=C["card"], bd=0, highlightthickness=1,
                         highlightbackground=C["border"], highlightcolor=C["border"])
        inner = ttk.Frame(self, style="Card.TFrame", padding=(16, 12))
        inner.pack(fill="both", expand=True)
        top = ttk.Frame(inner, style="Card.TFrame")
        top.pack(fill="x")
        if dot:
            mark = tk.Canvas(top, width=8, height=8, bg=C["card"], highlightthickness=0, bd=0)
            mark.create_oval(0, 0, 7, 7, fill=dot, outline=dot)
            mark.pack(side="left", padx=(0, 6))
        ttk.Label(top, text=label, style="Muted.TLabel").pack(side="left")
        if callable(value):
            value(inner).pack(anchor="w", pady=(4, 0))
        else:
            ttk.Label(inner, textvariable=value, style=value_style).pack(anchor="w", pady=(4, 0))
        if footnote:
            ttk.Label(inner, text=footnote, style="Hint.TLabel").pack(anchor="w")
        self.body = inner


def kpi_row(parent, tiles):
    """Lays `tiles` - (label, StringVar, kwargs) - out in equal columns."""
    row = ttk.Frame(parent)
    for i, (label, var, options) in enumerate(tiles):
        tile = KpiTile(row, label, var, **options)
        tile.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 6, 0 if i == len(tiles) - 1 else 6))
        row.columnconfigure(i, weight=1, uniform="kpi")
    return row


class Sidebar(tk.Frame):
    """The kit's app sidebar: brand mark, one navigation group, footer.
    `items` are (key, number, label); `on_select(key)` is called on click."""

    WIDTH = 256

    def __init__(self, parent, fonts, title, subtitle, items, on_select, footer_var=None):
        super().__init__(parent, bg=C["sidebar"], width=self.WIDTH, bd=0)
        self.pack_propagate(False)
        tk.Frame(self, bg=C["sidebar_border"], width=1).pack(side="right", fill="y")
        self._fonts = fonts
        self._on_select = on_select
        self._rows = {}
        self._active = None

        brand = tk.Frame(self, bg=C["sidebar"])
        brand.pack(fill="x", padx=16, pady=(16, 14))
        tk.Label(brand, text="✉", bg=C["primary"], fg=C["primary_foreground"], font=fonts.body_medium,
                 width=2, pady=2).pack(side="left")
        names = tk.Frame(brand, bg=C["sidebar"])
        names.pack(side="left", padx=(10, 0))
        tk.Label(names, text=title, bg=C["sidebar"], fg=C["foreground"], font=fonts.body_medium,
                 anchor="w").pack(anchor="w")
        tk.Label(names, text=subtitle, bg=C["sidebar"], fg=C["muted_foreground"], font=fonts.small,
                 anchor="w").pack(anchor="w")
        tk.Frame(self, bg=C["sidebar_border"], height=1).pack(fill="x")

        nav = tk.Frame(self, bg=C["sidebar"])
        nav.pack(fill="x", padx=12, pady=(14, 0))
        tk.Label(nav, text="Workflow", bg=C["sidebar"], fg=C["muted_foreground"], font=fonts.small_medium,
                 anchor="w").pack(fill="x", padx=12, pady=(0, 6))
        for key, number, label in items:
            self._rows[key] = self._nav_item(nav, key, number, label)

        if footer_var is not None:
            footer = tk.Frame(self, bg=C["sidebar"])
            footer.pack(side="bottom", fill="x")
            tk.Frame(footer, bg=C["sidebar_border"], height=1).pack(fill="x")
            tk.Label(footer, textvariable=footer_var, bg=C["sidebar"], fg=C["muted_foreground"],
                     font=fonts.small, anchor="w", justify="left", wraplength=self.WIDTH - 32
                     ).pack(fill="x", padx=16, pady=10)


    def _nav_item(self, parent, key, number, label):
        row = tk.Frame(parent, bg=C["sidebar"], cursor="hand2")
        row.pack(fill="x", pady=1)
        num = tk.Label(row, text=str(number), width=2, bg=C["sidebar"], fg=C["muted_foreground"],
                       font=self._fonts.small_medium)
        num.pack(side="left", padx=(10, 4), pady=8)
        text = tk.Label(row, text=label, bg=C["sidebar"], fg=C["foreground"],
                        font=self._fonts.body, anchor="w")
        text.pack(side="left", fill="x", expand=True, pady=8)
        widgets = (row, num, text)
        for w in widgets:
            w.bind("<Button-1>", lambda e, k=key: self._on_select(k))
            w.bind("<Enter>", lambda e, k=key: self._paint(k, hover=True))
            w.bind("<Leave>", lambda e, k=key: self._paint(k))
        return widgets

    def _paint(self, key, hover=False):
        row, num, text = self._rows[key]
        if key == self._active:
            bg, fg, num_fg, font = C["primary"], C["primary_foreground"], tint(C["primary_foreground"], 0.7, C["primary"]), self._fonts.body_medium
        elif hover:
            bg, fg, num_fg, font = C["sidebar_accent"], C["foreground"], C["muted_foreground"], self._fonts.body
        else:
            bg, fg, num_fg, font = C["sidebar"], tint(C["foreground"], 0.8, C["sidebar"]), C["muted_foreground"], self._fonts.body
        for w in (row, num, text):
            w.configure(bg=bg)
        num.configure(fg=num_fg)
        text.configure(fg=fg, font=font)

    def set_active(self, key):
        previous, self._active = self._active, key
        if previous in self._rows:
            self._paint(previous)
        if key in self._rows:
            self._paint(key)


class PageHeader(ttk.Frame):
    """Page title and description on the left; a context badge on the right."""

    def __init__(self, parent, fonts, badge_var=None):
        super().__init__(parent, padding=(28, 20, 28, 16))
        if badge_var is not None:
            badge = bordered(self)
            badge.pack(side="right", anchor="n")
            tk.Label(badge, textvariable=badge_var, bg=C["card"], fg=C["foreground"],
                     font=fonts.small_medium, padx=10, pady=4).pack()
        text = ttk.Frame(self)
        text.pack(side="left", fill="x", expand=True)
        self._title = ttk.Label(text, style="PageTitle.TLabel")
        self._title.pack(anchor="w")
        self._description = WrapLabel(text, style="Muted.TLabel")
        self._description.pack(anchor="w", fill="x", pady=(2, 0))

    def set(self, title, description):
        self._title.configure(text=title)
        self._description.configure(text=description)

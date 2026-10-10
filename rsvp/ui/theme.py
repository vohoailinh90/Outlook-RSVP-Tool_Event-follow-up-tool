"""ttk styles that give the app the Automation UI Kit's look.

The kit is shadcn/ui on the web: a white page, white cards with a hairline
border, near-black primary buttons, outline secondary buttons, quiet grey
labels, and colour reserved for status. Tk cannot draw rounded corners or
shadows, so this keeps everything else: the tokens, the type scale, the
spacing, flat controls, and the rule that colour means something.

`apply(root)` must run once, right after the Tk root exists and before any
widget is built. It returns the fonts the app builds its widgets with.
"""
from __future__ import annotations

import tkinter as tk
from dataclasses import dataclass
from tkinter import font as tkfont
from tkinter import ttk

from .tokens import COLORS

C = COLORS

# Space scale, in pixels (Tailwind's 4px grid, as the kit uses it).
SPACE = {"xs": 4, "sm": 8, "md": 12, "lg": 16, "xl": 24}

SCROLLBAR_THUMB = "#d4d4d4"   # neutral-300, the kit's scrollbar thumb


@dataclass(frozen=True)
class Fonts:
    body: str
    body_medium: str
    small: str
    small_medium: str
    title: str
    page_title: str
    kpi: str
    mono: str


def _families(root):
    return set(tkfont.families(root))


def _make_fonts(root):
    families = _families(root)
    family = next((f for f in ("Segoe UI", "Inter", "Helvetica Neue", "Arial") if f in families),
                  tkfont.nametofont("TkDefaultFont").actual("family"))
    # Windows ships the semibold cut as its own family; elsewhere use bold.
    semibold = "Segoe UI Semibold" if family == "Segoe UI" and "Segoe UI Semibold" in families else None
    mono = next((f for f in ("Cascadia Mono", "Consolas", "DejaVu Sans Mono") if f in families), "TkFixedFont")

    # tkinter deletes a named font when the Font object that created it is
    # garbage-collected, and every style naming it silently falls back to
    # the default; so the objects live as long as the root window.
    keep = root.__dict__.setdefault("_rsvp_fonts", {})

    def font(name, size, strong=False, fam=None):
        options = {"family": fam or (semibold if strong and semibold else family), "size": size,
                   "weight": "bold" if strong and not semibold else "normal"}
        if name in keep:
            keep[name].configure(**options)
        else:
            keep[name] = tkfont.Font(root, name=name, **options)
        return name

    fonts = Fonts(
        body=font("RsvpBody", 10),
        body_medium=font("RsvpBodyMedium", 10, strong=True),
        small=font("RsvpSmall", 9),
        small_medium=font("RsvpSmallMedium", 9, strong=True),
        title=font("RsvpTitle", 11, strong=True),
        page_title=font("RsvpPageTitle", 15, strong=True),
        kpi=font("RsvpKpi", 18, strong=True),
        mono=font("RsvpMono", 10, fam=mono),
    )
    # Every widget that does not name a font - dialogs, menus, tkcalendar -
    # picks up the body font too.
    for named in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
        tkfont.nametofont(named).configure(family=family, size=10)
    return fonts


def apply(root):
    """Styles every ttk widget class the app uses and sets defaults for the
    plain tk ones (Text, Listbox, Entry). Returns the Fonts."""
    fonts = _make_fonts(root)
    style = ttk.Style(root)
    style.theme_use("clam")   # a pure-Tk theme: honours every colour below
    root.configure(background=C["background"])

    linespace = tkfont.Font(root, font=fonts.body).metrics("linespace")

    style.configure(
        ".", background=C["background"], foreground=C["foreground"], font=fonts.body,
        bordercolor=C["border"], lightcolor=C["background"], darkcolor=C["background"],
        troughcolor=C["muted"], focuscolor=C["ring"], selectbackground=C["selection"],
        selectforeground=C["foreground"], insertcolor=C["foreground"])

    # ── surfaces and text ──
    style.configure("TFrame", background=C["background"])
    style.configure("Card.TFrame", background=C["card"])
    style.configure("TLabel", background=C["background"], foreground=C["foreground"])
    style.configure("Muted.TLabel", foreground=C["muted_foreground"])
    style.configure("Hint.TLabel", foreground=C["muted_foreground"], font=fonts.small)
    style.configure("Strong.TLabel", font=fonts.body_medium)
    style.configure("Field.TLabel", font=fonts.small_medium)
    style.configure("CardTitle.TLabel", font=fonts.title)
    style.configure("PageTitle.TLabel", font=fonts.page_title)
    style.configure("KpiValue.TLabel", font=fonts.kpi)
    style.configure("Mono.TLabel", font=fonts.mono)
    for tone, color in (("Success", C["success"]), ("Danger", C["destructive"]), ("Info", C["info"])):
        style.configure(f"{tone}.TLabel", foreground=color)
        style.configure(f"{tone}Kpi.TLabel", foreground=color, font=fonts.kpi)
    style.configure("TSeparator", background=C["border"])

    # ── buttons: default is the kit's "outline", plus primary / destructive / ghost ──
    def button(name, bg, fg, border, hover, pressed=None):
        # clam draws the 1px border only for a raised relief; with the bevel
        # colours equal to the fill it is a plain hairline, like the kit's.
        style.configure(name, background=bg, foreground=fg, bordercolor=border,
                        lightcolor=bg, darkcolor=bg, padding=(14, 6), relief="raised",
                        borderwidth=1, font=fonts.body_medium, focusthickness=0, anchor="center")
        pressed = pressed or hover
        style.map(name,
                  relief=[("pressed", "raised")],
                  background=[("disabled", bg), ("pressed", pressed), ("active", hover)],
                  lightcolor=[("pressed", pressed), ("active", hover)],
                  darkcolor=[("pressed", pressed), ("active", hover)],
                  bordercolor=[("focus", C["ring"])],
                  foreground=[("disabled", C["muted_foreground"])])

    button("TButton", C["card"], C["foreground"], C["border"], C["muted"], C["border"])
    button("Primary.TButton", C["primary"], C["primary_foreground"], C["primary"], C["primary_hover"])
    button("Destructive.TButton", C["destructive"], "#ffffff", C["destructive"], C["destructive_hover"])
    button("Ghost.TButton", C["background"], C["foreground"], C["background"], C["muted"], C["border"])
    style.configure("Small.TButton", padding=(10, 3), font=fonts.small_medium)

    # ── inputs ──
    field = dict(fieldbackground=C["card"], foreground=C["foreground"], bordercolor=C["border"],
                 lightcolor=C["card"], darkcolor=C["card"], insertcolor=C["foreground"])
    style.configure("TEntry", padding=(8, 5), **field)
    style.map("TEntry", bordercolor=[("focus", C["ring"])],
              fieldbackground=[("readonly", C["muted"]), ("disabled", C["muted"])])
    style.configure("TCombobox", padding=(8, 4), background=C["card"], arrowcolor=C["muted_foreground"],
                    arrowsize=14, **field)
    style.map("TCombobox",
              fieldbackground=[("readonly", C["card"]), ("disabled", C["muted"])],
              selectbackground=[("readonly", C["card"])],
              selectforeground=[("readonly", C["foreground"])],
              background=[("active", C["card"]), ("pressed", C["card"])],
              bordercolor=[("focus", C["ring"])])
    root.option_add("*TCombobox*Listbox.background", C["card"])
    root.option_add("*TCombobox*Listbox.foreground", C["foreground"])
    root.option_add("*TCombobox*Listbox.selectBackground", C["selection"])
    root.option_add("*TCombobox*Listbox.selectForeground", C["foreground"])
    root.option_add("*TCombobox*Listbox.font", fonts.body)
    root.option_add("*TCombobox*Listbox.relief", "flat")

    for name in ("TCheckbutton", "TRadiobutton"):
        style.configure(name, background=C["background"], foreground=C["foreground"],
                        indicatorbackground=C["card"], indicatorforeground=C["primary_foreground"],
                        upperbordercolor=C["border"], lowerbordercolor=C["border"],
                        indicatormargin=(0, 0, 8, 0), focusthickness=0)
        style.map(name,
                  background=[("active", C["background"])],
                  indicatorbackground=[("selected", C["primary"]), ("pressed", C["muted"])],
                  upperbordercolor=[("selected", C["primary"])],
                  lowerbordercolor=[("selected", C["primary"])])

    # ── tables ──
    style.layout("Treeview", [("Treeview.treearea", {"sticky": "nswe"})])   # the card draws the border
    style.configure("Treeview", background=C["card"], fieldbackground=C["card"],
                    foreground=C["foreground"], rowheight=linespace + 14, font=fonts.body,
                    borderwidth=0)
    style.map("Treeview", background=[("selected", C["selection"])],
              foreground=[("selected", C["foreground"])])
    style.configure("Treeview.Heading", background=C["card"], foreground=C["muted_foreground"],
                    font=fonts.small_medium, relief="flat", padding=(8, 7),
                    bordercolor=C["border"], lightcolor=C["card"], darkcolor=C["card"])
    style.map("Treeview.Heading", background=[("active", C["muted"])],
              lightcolor=[("active", C["muted"])], darkcolor=[("active", C["muted"])])

    # ── thin scrollbars without arrows ──
    for orient, sticky in (("Vertical", "ns"), ("Horizontal", "ew")):
        style.layout(f"{orient}.TScrollbar", [(f"{orient}.Scrollbar.trough", {
            "sticky": sticky,
            "children": [(f"{orient}.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})]})])
    style.configure("TScrollbar", troughcolor=C["background"], background=SCROLLBAR_THUMB,
                    bordercolor=C["background"], lightcolor=SCROLLBAR_THUMB,
                    darkcolor=SCROLLBAR_THUMB, arrowsize=9, gripcount=0, relief="flat")
    style.map("TScrollbar", background=[("active", C["muted_foreground"])],
              lightcolor=[("active", C["muted_foreground"])],
              darkcolor=[("active", C["muted_foreground"])])

    # ── the page stack: a notebook whose tab strip is replaced by the sidebar ──
    style.layout("Pages.TNotebook.Tab", [])
    style.layout("Pages.TNotebook", [("Notebook.client", {"sticky": "nswe"})])
    style.configure("Pages.TNotebook", background=C["background"], borderwidth=0, padding=0,
                    tabmargins=0, bordercolor=C["background"], lightcolor=C["background"],
                    darkcolor=C["background"])

    # ── plain tk widgets ──
    for widget in ("Text", "Listbox", "Entry"):
        root.option_add(f"*{widget}.background", C["card"])
        root.option_add(f"*{widget}.foreground", C["foreground"])
        root.option_add(f"*{widget}.font", fonts.body)
        root.option_add(f"*{widget}.relief", "flat")
        root.option_add(f"*{widget}.borderWidth", 0)
        root.option_add(f"*{widget}.selectBackground", C["selection"])
        root.option_add(f"*{widget}.selectForeground", C["foreground"])
        root.option_add(f"*{widget}.highlightThickness", 0)
    root.option_add("*Text.insertBackground", C["foreground"])
    root.option_add("*Entry.insertBackground", C["foreground"])
    root.option_add("*Text.padX", 10)
    root.option_add("*Text.padY", 8)
    root.option_add("*Listbox.activeStyle", "none")
    return fonts

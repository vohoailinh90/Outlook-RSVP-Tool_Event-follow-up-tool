"""A broad corpus of budget strings a human might actually type.

Money parsing decides what colleagues are asked to contribute. Hand-picked
examples in test_domain_money.py document the RULES; this corpus catches the
cases nobody thought to pick, by rendering every combination and comparing the
whole result set against a golden file.

Built from parts rather than listed, so the combinations that break things are
present without anyone having had to imagine them.
"""
from __future__ import annotations

import itertools

NUMBERS = ["3000", "3,000", "3.000", "3000.50", "3,000.50", "1.234.567",
           "1,234,567", "12.5", "12,5", "0", "0.500", "500", "250000",
           "2026", "3.5", "1.50", "10.000", "100.000"]

# "$", "US$" and " EUR" were added with the USD decimal rule: it turns on
# exactly these markers ("US$" in, bare "$" and EUR deliberately out), so the
# corpus must exercise them.
CURRENCIES = ["", " JPY", " VND", "¥", "円", "đ", " USD", " yen", " dong",
              "$", "US$", " EUR"]

CONTEXTS = [
    "{n}{c}",
    "{c}{n}",
    "Budget: {n}{c}",
    "{n}{c} / person",
    "{n}{c}/người",
    "2026 year-end party, {n}{c}",
    "5 people x {n}{c}",
    "approx {n}{c} each",
    "予算 {n}{c}",
    "khoảng {n}{c} mỗi người",
    "Event 2026 - {n}{c}",
    "{n}{c} (tentative)",
    "deposit {n}{c} due 2026-12-10",
    "Room 205, budget {n}{c}",
]

EDGE_CASES = ["", "   ", "no digits", "-500", "abc", "2026", "0.000",
              "1.000.000,50", "3000-4000", "12:30", "2026-12-20"]


def corpus() -> list[str]:
    cases = [t.format(n=n, c=c)
             for n, c, t in itertools.product(NUMBERS, CURRENCIES, CONTEXTS)]
    return cases + EDGE_CASES


def build_snapshot() -> dict[str, str]:
    from rsvp.domain import parse_amount_from_text
    return {case: repr(parse_amount_from_text(case)) for case in corpus()}

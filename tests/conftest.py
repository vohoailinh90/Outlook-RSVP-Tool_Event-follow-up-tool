"""Import the app modules, or fail loudly where they are supposed to work.

The characterization tests target the platform-neutral functions in
rsvp_app.py - subject/body builders, amount parsing, paste cleanup. Those need
no Outlook and no display, but they live in a module whose top-level imports
pull in tkinter and outlook_com, so on a machine without them the whole test
file is unimportable.

Skipping is the right behavior on a developer's Linux box. It is the WRONG
behavior in CI, where a silent skip turns "all tests pass" into "no tests
ran" - the exact false-confidence failure CLAUDE.md warns about. So CI sets
RSVP_REQUIRE_APP_IMPORT=1 and a failed import becomes an error instead.
"""
import os

import pytest

REQUIRE = os.environ.get("RSVP_REQUIRE_APP_IMPORT") == "1"

try:
    import rsvp_app
    IMPORT_ERROR = None
except Exception as exc:  # pragma: no cover - environment-dependent
    rsvp_app = None
    IMPORT_ERROR = exc


@pytest.fixture(scope="session")
def monolith():
    """The rsvp_app module itself, or skip/fail depending on the environment.

    Named `monolith` rather than `app` because most tests no longer need it:
    after phase 1 the message builders come from rsvp.i18n, which imports
    anywhere. This fixture is only for code still inside rsvp_app.py.
    """
    if rsvp_app is not None:
        return rsvp_app
    msg = f"rsvp_app is not importable here: {IMPORT_ERROR!r}"
    if REQUIRE:
        pytest.fail(
            msg + "\n\nRSVP_REQUIRE_APP_IMPORT=1 is set, which means this "
            "environment is supposed to be able to import it. A skip here "
            "would report a green suite that tested nothing."
        )
    pytest.skip(msg + " (set RSVP_REQUIRE_APP_IMPORT=1 to make this an error)")

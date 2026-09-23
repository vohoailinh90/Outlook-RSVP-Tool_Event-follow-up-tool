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

# Fail the whole run here, at collection, not in the fixture below: no test
# requests `monolith`, so a failure deferred to it never fired, and an import
# of a nonexistent module added to rsvp_app.py still left CI green (Codex
# review of PR #1). The import of rsvp_app is itself the check - it is how a
# name dropped from the compatibility shim at module level shows up.
if REQUIRE and IMPORT_ERROR is not None:
    pytest.exit(
        f"rsvp_app is not importable here: {IMPORT_ERROR!r}\n\n"
        "RSVP_REQUIRE_APP_IMPORT=1 is set, which means this environment is "
        "supposed to be able to import it. A skip here would report a green "
        "suite that tested nothing.",
        returncode=1)


@pytest.fixture(scope="session")
def monolith():
    """The rsvp_app module itself, or skip/fail depending on the environment.

    Named `monolith` rather than `app` because most tests no longer need it:
    after phase 1 the message builders come from rsvp.i18n, which imports
    anywhere. This fixture is only for code still inside rsvp_app.py.
    """
    if rsvp_app is not None:
        return rsvp_app
    pytest.skip(f"rsvp_app is not importable here: {IMPORT_ERROR!r} "
                "(set RSVP_REQUIRE_APP_IMPORT=1 to make this an error)")

"""Pytest configuration and shared fixtures."""

from typing import TYPE_CHECKING

import pytest

if TYPE_CHECKING:
    import wx


@pytest.fixture(scope="package")
def wx_app() -> "wx.App":
    """Initialize wx application for test packages.

    This shared fixture ensures wx is initialized before any wx-dependent imports
    or operations. It uses delayed import to avoid issues on macOS where code
    formatters may trigger Python icon display.

    The fixture is package-scoped, so a single wx.App instance is created per test
    package and reused across all tests in that package.

    :returns: wx.App instance (creates new if none exists, otherwise returns existing)
    """
    # Delayed import of wx - only load when needed for tests
    import wx

    # Ensure wx is initialized (required for CommandUI which inherits from wx.EvtHandler)
    if not wx.GetApp():
        return wx.App(False)
    return wx.GetApp()

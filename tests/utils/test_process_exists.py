"""Tests for process_exists module.

This module contains unit tests for process existence checking.
"""

import sys

import pytest

from service_ppt.utils.process_exists import process_exists


class TestProcessExists:
    """Test process_exists function."""

    @pytest.mark.skipif(sys.platform == "win32", reason="Unix-specific test")
    def test_unix_returns_boolean(self):
        """Test that process_exists returns a boolean value on Unix."""
        # On Unix, the function uses pgrep which may raise FileNotFoundError
        # if pgrep is not available. We'll test with a process that might exist.
        try:
            result = process_exists("python3")
            assert isinstance(result, bool)
        except FileNotFoundError:
            # pgrep not available, skip this test
            pytest.skip("pgrep command not available")

    @pytest.mark.skipif(sys.platform != "win32", reason="Windows-specific test")
    def test_windows_returns_boolean(self):
        """Test that process_exists returns a boolean value on Windows."""
        # On Windows, check for a common process
        result = process_exists("python.exe")
        assert isinstance(result, bool)

    @pytest.mark.skipif(sys.platform == "win32", reason="Unix-specific test")
    def test_unix_handles_nonexistent_process(self):
        """Test that process_exists handles nonexistent process on Unix."""
        try:
            # Use a very unlikely process name
            result = process_exists("definitely_nonexistent_process_xyz789")
            assert result is False
        except FileNotFoundError:
            # pgrep not available, skip this test
            pytest.skip("pgrep command not available")

    @pytest.mark.skipif(sys.platform != "win32", reason="Windows-specific test")
    def test_windows_handles_nonexistent_process(self):
        """Test that process_exists handles nonexistent process on Windows."""
        result = process_exists("definitely_nonexistent_process_xyz789.exe")
        assert result is False

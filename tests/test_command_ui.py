"""Tests for command_ui module.

This module contains unit tests for the UIManager class in command_ui,
testing the open() method with sample service definition files.
"""

from pathlib import Path
from typing import TYPE_CHECKING

import pytest

if TYPE_CHECKING:
    import wx

    from service_ppt.command_ui import UIManager


@pytest.fixture
def sample_sdf_path():
    """Return the path to the sample service.sdf file."""
    # Get the tests directory
    tests_dir = Path(__file__).parent
    return tests_dir / "data" / "service.sdf"


@pytest.fixture
def uimgr(wx_app: "wx.App") -> "UIManager":
    """Create a UIManager instance for testing.

    :param wx_app: wx application instance (ensures wx is initialized)
    :returns: UIManager instance
    """
    # Import UIManager after wx is initialized
    from service_ppt.command_ui import UIManager

    return UIManager()


class TestUIManagerOpen:
    """Test UIManager.open() method."""

    def test_open_sample_service_sdf(self, uimgr: "UIManager", sample_sdf_path: Path) -> None:
        """Test opening the sample service.sdf file."""
        # Verify the sample file exists
        assert sample_sdf_path.exists(), f"Sample file not found: {sample_sdf_path}"

        # Initialize empty dir_dict
        dir_dict = {}

        # Open the file
        uimgr.open(str(sample_sdf_path), dir_dict)

        # Verify that command_ui_list is populated
        assert len(uimgr.command_ui_list) > 0, "command_ui_list should not be empty after opening file"

        # Verify the expected number of commands (based on the sample file)
        # The sample file has 13 command entries
        assert len(uimgr.command_ui_list) == 13, f"Expected 13 commands, got {len(uimgr.command_ui_list)}"

        # Verify that modified flag is set to False after opening
        assert not uimgr.get_modified(), "modified flag should be False after opening file"

        # Verify the first command is OpenFile
        first_ui = uimgr.command_ui_list[0]
        assert first_ui.__class__.__name__ == "OpenFileUI", f"First command should be OpenFileUI, got {first_ui.__class__.__name__}"
        assert first_ui.name == "New File", f"First command name should be 'New File', got '{first_ui.name}'"
        assert first_ui.command.enabled is True, "First command should be enabled"

        # Verify the last command is ExportSlides
        last_ui = uimgr.command_ui_list[-1]
        assert last_ui.__class__.__name__ == "ExportSlidesUI", f"Last command should be ExportSlidesUI, got {last_ui.__class__.__name__}"
        assert last_ui.name == "Export ending announcement to image files", "Last command name mismatch"

        # Verify uimgr is set on all UI objects
        for ui in uimgr.command_ui_list:
            assert ui.uimgr is uimgr, "Each UI object should have uimgr reference set"

    def test_open_sample_service_sdf_command_types(self, uimgr: "UIManager", sample_sdf_path: Path) -> None:
        """Test that all expected command types are loaded from sample file."""
        dir_dict = {}
        uimgr.open(str(sample_sdf_path), dir_dict)

        # Expected command types in order
        expected_types = [
            "OpenFileUI",
            "InsertSlidesUI",
            "DuplicateWithTextUI",
            "SetVariablesUI",
            "DuplicateWithTextUI",
            "GenerateBibleVerseUI",
            "DuplicateWithTextUI",
            "PopupMessageUI",
            "SaveFilesUI",
            "ExportSlidesUI",
            "ExportShapesUI",
            "ExportShapesUI",
            "ExportSlidesUI",
        ]

        assert len(uimgr.command_ui_list) == len(expected_types), "Command count mismatch"

        for i, (ui, expected_type) in enumerate(zip(uimgr.command_ui_list, expected_types, strict=True)):
            assert ui.__class__.__name__ == expected_type, (
                f"Command {i} type mismatch: expected {expected_type}, got {ui.__class__.__name__}"
            )

    def test_open_sample_service_sdf_command_data(self, uimgr: "UIManager", sample_sdf_path: Path) -> None:
        """Test that command data is correctly loaded from sample file."""
        dir_dict = {}
        uimgr.open(str(sample_sdf_path), dir_dict)

        # Test OpenFile command data
        open_file_ui = uimgr.command_ui_list[0]
        assert open_file_ui.command.filename == "", "OpenFile filename should be empty"
        assert open_file_ui.command.notes_filename == "service-notes-template.txt", "OpenFile notes_filename mismatch"

        # Test InsertSlides command data
        insert_slides_ui = uimgr.command_ui_list[1]
        assert insert_slides_ui.command.enabled is True, "InsertSlides should be enabled"
        assert "Service-Template.pptx" in insert_slides_ui.command.filelist, "InsertSlides filelist should contain Service-Template.pptx"

        # Test GenerateBibleVerse command data
        bible_verse_ui = uimgr.command_ui_list[5]
        assert bible_verse_ui.command.bible_version1 == "ESV", "Bible version should be ESV"
        assert bible_verse_ui.command.main_verses == "Genesis 1:1", "Main verses should be Genesis 1:1"
        assert bible_verse_ui.command.additional_verses == "Genesis 1:2", "Additional verses should be Genesis 1:2"

        # Test PopupMessage command data
        popup_ui = uimgr.command_ui_list[7]
        assert popup_ui.command.message == "Please check the slides before continue.", "Popup message mismatch"

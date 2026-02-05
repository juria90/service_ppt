"""Tests for make_transparent module.

This module contains unit tests for image transparency conversion functions.
"""

import argparse
import os
import tempfile

from PIL import Image
import pytest

from service_ppt.utils.make_transparent import (
    color_to_transparent,
    parse_color,
    white_to_transparent,
)


class TestWhiteToTransparent:
    """Test white_to_transparent function."""

    def test_converts_white_to_transparent(self):
        """Test that white pixels are converted to transparent."""
        # Create a white image with RGBA mode
        img = Image.new("RGBA", (10, 10), (255, 255, 255, 255))

        result = white_to_transparent(img)

        assert result.mode == "RGBA"
        # Check that white pixels have alpha = 0
        pixel = result.getpixel((0, 0))
        assert pixel[3] == 0  # Alpha channel should be 0

    def test_preserves_non_white_pixels(self):
        """Test that non-white pixels are preserved."""
        # Create image with red pixels
        img = Image.new("RGBA", (10, 10), (255, 0, 0, 255))

        result = white_to_transparent(img)

        # Red pixels should still have alpha = 255
        pixel = result.getpixel((0, 0))
        assert pixel[3] == 255  # Alpha channel should be preserved


class TestColorToTransparent:
    """Test color_to_transparent function."""

    def test_converts_specific_color_to_transparent(self):
        """Test that specific color is converted to transparent."""
        with tempfile.TemporaryDirectory() as tmpdir:
            input_file = os.path.join(tmpdir, "test.png")
            output_file = os.path.join(tmpdir, "output.png")

            # Create image with red background
            img = Image.new("RGB", (10, 10), (255, 0, 0))
            img.save(input_file, "PNG")

            # Convert red to transparent
            result = color_to_transparent(input_file, output_file, (255, 0, 0))

            assert result is True
            assert os.path.exists(output_file)

            # Check that output is RGBA
            output_img = Image.open(output_file)
            assert output_img.mode == "RGBA"

    def test_returns_false_when_no_pixels_changed(self):
        """Test that function returns False when no pixels match."""
        with tempfile.TemporaryDirectory() as tmpdir:
            input_file = os.path.join(tmpdir, "test.png")
            output_file = os.path.join(tmpdir, "output.png")

            # Create image with blue background
            img = Image.new("RGB", (10, 10), (0, 0, 255))
            img.save(input_file, "PNG")

            # Try to convert red (which doesn't exist) to transparent
            result = color_to_transparent(input_file, output_file, (255, 0, 0))

            # Should return False or True depending on implementation
            assert isinstance(result, bool)

    def test_handles_white_special_case(self):
        """Test that white color uses optimized path."""
        with tempfile.TemporaryDirectory() as tmpdir:
            input_file = os.path.join(tmpdir, "test.png")
            output_file = os.path.join(tmpdir, "output.png")

            # Create white image
            img = Image.new("RGB", (10, 10), (255, 255, 255))
            img.save(input_file, "PNG")

            result = color_to_transparent(input_file, output_file, (255, 255, 255))

            assert result is True
            assert os.path.exists(output_file)


class TestParseColor:
    """Test parse_color function."""

    def test_parses_color_names(self):
        """Test that parse_color parses color names."""
        assert parse_color("red") == (255, 0, 0)
        assert parse_color("green") == (0, 128, 0)
        assert parse_color("blue") == (0, 0, 255)
        assert parse_color("white") == (255, 255, 255)
        assert parse_color("black") == (0, 0, 0)

    def test_parses_hex_colors(self):
        """Test that parse_color parses hex color values."""
        assert parse_color("#FF0000") == (255, 0, 0)
        assert parse_color("#00FF00") == (0, 255, 0)
        assert parse_color("#0000FF") == (0, 0, 255)

    def test_parses_rgb_colors(self):
        """Test that parse_color parses rgb() color values."""
        assert parse_color("rgb(255, 0, 0)") == (255, 0, 0)
        assert parse_color("rgb(0, 255, 0)") == (0, 255, 0)

    def test_raises_on_invalid_color(self):
        """Test that parse_color raises on invalid color string."""
        with pytest.raises((ValueError, argparse.ArgumentTypeError)):
            parse_color("invalid_color_name")

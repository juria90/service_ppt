#!/usr/bin/env python3
"""Update and compile translation files.

This script merges .pot template files with existing .po files and
compiles them to .mo binary format for use by the application.

Requires GNU gettext package to be installed and available in PATH.
"""

import os
import shutil
import subprocess
import sys
from pathlib import Path


def find_command(cmd):
    """Find command executable in PATH."""
    found = shutil.which(cmd)
    if not found:
        print(f"Error: {cmd} not found in PATH. Please install GNU gettext.", file=sys.stderr)
        sys.exit(1)
    return found


def main():
    """Main function to update and compile translations."""
    # Get the project root directory (parent of scripts/)
    project_root = Path(__file__).parent.parent
    os.chdir(project_root)

    msgmerge = find_command("msgmerge")
    msgfmt = find_command("msgfmt")

    # Locales to process
    locales = ["ko"]

    # Process bible translations
    print("Processing bible translations...")
    for locale in locales:
        po_file = f"service_ppt/locale/{locale}/LC_MESSAGES/bible.po"
        pot_file = "service_ppt/locale/bible.pot"
        mo_file = f"service_ppt/locale/{locale}/LC_MESSAGES/bible.mo"

        # Ensure directory exists
        Path(po_file).parent.mkdir(parents=True, exist_ok=True)

        # Merge .pot with .po
        if Path(pot_file).exists():
            print(f"  Merging {pot_file} into {po_file}...")
            result = subprocess.run(
                [msgmerge, "-s", "-U", po_file, pot_file],
                check=False,
            )
            if result.returncode != 0:
                print(f"  Warning: msgmerge for {po_file} returned {result.returncode}", file=sys.stderr)
        else:
            print(f"  Warning: {pot_file} not found, skipping merge", file=sys.stderr)

        # Compile .po to .mo
        if Path(po_file).exists():
            print(f"  Compiling {po_file} to {mo_file}...")
            result = subprocess.run(
                [msgfmt, po_file, "-o", mo_file],
                check=False,
            )
            if result.returncode != 0:
                print(f"  Warning: msgfmt for {po_file} returned {result.returncode}", file=sys.stderr)
        else:
            print(f"  Warning: {po_file} not found, skipping compilation", file=sys.stderr)

    # Process service_ppt translations
    print("Processing service_ppt translations...")
    for locale in locales:
        po_file = f"service_ppt/locale/{locale}/LC_MESSAGES/service_ppt.po"
        pot_file = "service_ppt/locale/service_ppt.pot"
        mo_file = f"service_ppt/locale/{locale}/LC_MESSAGES/service_ppt.mo"

        # Ensure directory exists
        Path(po_file).parent.mkdir(parents=True, exist_ok=True)

        # Merge .pot with .po
        if Path(pot_file).exists():
            print(f"  Merging {pot_file} into {po_file}...")
            result = subprocess.run(
                [msgmerge, "-s", "-U", po_file, pot_file],
                check=False,
            )
            if result.returncode != 0:
                print(f"  Warning: msgmerge for {po_file} returned {result.returncode}", file=sys.stderr)
        else:
            print(f"  Warning: {pot_file} not found, skipping merge", file=sys.stderr)

        # Compile .po to .mo
        if Path(po_file).exists():
            print(f"  Compiling {po_file} to {mo_file}...")
            result = subprocess.run(
                [msgfmt, po_file, "-o", mo_file],
                check=False,
            )
            if result.returncode != 0:
                print(f"  Warning: msgfmt for {po_file} returned {result.returncode}", file=sys.stderr)
        else:
            print(f"  Warning: {po_file} not found, skipping compilation", file=sys.stderr)

    print("Done!")


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Extract translatable strings using xgettext.

This script extracts translatable strings from Python source files and
generates .pot template files for translation.

xgettext for Python does not parse pgettext by default, so we add it.
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
    """Main function to extract translatable strings."""
    # Get the project root directory (parent of scripts/)
    project_root = Path(__file__).parent.parent
    os.chdir(project_root)

    xgettext = find_command("xgettext")

    # Extract strings for bible module
    bible_files = [
        "service_ppt/bible/bibcore.py",
        "service_ppt/bible/biblang.py",
    ]
    bible_pot = "service_ppt/locale/bible.pot"

    print(f"Extracting strings for bible module...")
    cmd = [
        xgettext,
        "-kpgettext:1c,2",  # xgettext for Python does not parse pgettext by default
        "-d",
        "bible",
        "-o",
        bible_pot,
    ] + bible_files

    result = subprocess.run(cmd, check=False)
    if result.returncode != 0:
        print(f"Warning: xgettext for bible module returned {result.returncode}", file=sys.stderr)
    else:
        print(f"Created {bible_pot}")

    # Extract strings for service_ppt module
    service_ppt_files = [
        "service_ppt/mainframe.py",
        "service_ppt/preferences_dialog.py",
        "service_ppt/command.py",
        "service_ppt/command_ui.py",
    ]
    service_ppt_pot = "service_ppt/locale/service_ppt.pot"

    print(f"Extracting strings for service_ppt module...")
    cmd = [
        xgettext,
        "-d",
        "service_ppt",
        "-o",
        service_ppt_pot,
    ] + service_ppt_files

    result = subprocess.run(cmd, check=False)
    if result.returncode != 0:
        print(f"Warning: xgettext for service_ppt module returned {result.returncode}", file=sys.stderr)
        sys.exit(1)
    else:
        print(f"Created {service_ppt_pot}")

    print("Done!")


if __name__ == "__main__":
    main()

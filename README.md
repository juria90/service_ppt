# What is service_ppt.
The service_ppt is an application generating PowerPoint slides for the Church's service.<br>
It produces Bible Verse, lyric and announcement slides based on template with find and replace text.<br>
The find and replace text can be applied to a text template file and can be saved as a text file for other purposes.<br>
It also can insert slides that are pre-made such as song slides.<br>
After it generates slide, it can export slides into images or shapes in slides into transparent images.<br>
<br>
PowerPoint automation is handled by platform-specific backends:<br>
- **Windows**: Uses COM automation (requires PowerPoint installed)<br>
- **macOS/Linux**: Uses python-pptx library by default (no PowerPoint installation required), with AppleScript fallback on macOS if python-pptx is unavailable<br>
Note: Image export features require platform-specific backends (COM on Windows, AppleScript on macOS).<br>

# Installation
This document describes how to set up the environment for the service_ppt application.

## Install Python 3.12 or higher
Install Python 3.12 or higher from https://www.python.org/downloads/.<br>
Add python3 to the PATH environment.

## Install the package

### Using a virtual environment (recommended)
```bash
# Create a virtual environment
python -m venv venv

# Activate the virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate

# Install the package in editable mode
pip install -e .
```

### Install with development dependencies
To install with development tools (ruff, pytest):
```bash
pip install -e ".[dev]"
```

### Code Style (PEP 8)
The project follows PEP 8 style guidelines. Code is automatically checked and formatted using `ruff`. Run `ruff check --fix service_ppt/` to check for violations and auto-fix issues, or `ruff format service_ppt/` to format code.

### What gets installed
The `pyproject.toml` file automatically handles all dependencies:
- **Core dependencies**: wxPython, iso639-lang, langdetect, numpy, Pillow, pysword, python-pptx, six
- **Platform-specific dependencies** (installed automatically based on your OS):
  - `pywin32` on Windows
  - `py-applescript` on macOS

### Running the application
After installation, you can run the application using:
```bash
# Using the console script (if installed in PATH)
service-ppt

# Or using Python module syntax
python -m service_ppt
```

## Set up environment for language translation

### Install GNU gettext
Install GNU gettext package and ensure the bin directory is in your PATH environment.<br>
The `xgettext`, `msgfmt` and `msgmerge` commands will be used for translation strings.

- **Windows**: Download from https://mlocati.github.io/articles/gettext-iconv-windows.html or http://gnuwin32.sourceforge.net/packages/gettext.htm
- **macOS**: Install via Homebrew: `brew install gettext`
- **Linux**: Install via package manager (e.g., `apt-get install gettext` on Debian/Ubuntu)

### Using the translation scripts

The project includes cross-platform Python scripts in the `scripts/` directory for managing translations:

#### Extract translatable strings

Run `scripts/xgettext.py` to extract translatable strings from the source code and generate `.pot` template files:

```bash
python scripts/xgettext.py
```

This script:
- Extracts strings from `service_ppt/bible/bibcore.py` and `service_ppt/bible/biblang.py` → `service_ppt/locale/bible.pot`
- Extracts strings from main application files → `service_ppt/locale/service_ppt.pot`
- Handles `pgettext` function calls (used for context-aware translations)

#### Update and compile translations

Run `scripts/msgupdate.py` to merge template files with existing translations and compile them:

```bash
python scripts/msgupdate.py
```

This script:
- Merges `.pot` template files with existing `.po` translation files
- Compiles `.po` files to `.mo` binary format (required at runtime)
- Processes all locales defined in the script (currently: `ko`)

### Translation workflow

1. **Extract strings**: Run `python scripts/xgettext.py` after adding new translatable strings
2. **Translate**: Edit the `.po` files in `service_ppt/locale/<locale>/LC_MESSAGES/` (e.g., `service_ppt/locale/ko/LC_MESSAGES/service_ppt.po`)
3. **Update and compile**: Run `python scripts/msgupdate.py` to merge changes and compile translations
4. **Test**: Restart the application to see translated strings

The compiled `.mo` files are automatically included in the package distribution.

# Set up data

# # Set up bible text.
The Bible text can be retrieved from MyBible(Application's proprietary format),<br>
MySword(https://www.mysword.info/modules-format),<br>
Sword(https://pypi.org/project/pysword/0),<br>
and Zefania(https://sourceforge.net/projects/zefania-sharp/files/Bibles/) format.<br>
After installing required bible package and Bible text,<br>
you need to set up the Active Bible from Application's preference.<br>

## Set up hymn lyrics.
The supported Hymn and lyric format is OpenLyric(https://docs.openlyrics.org/en/latest/).<br>
But, it can be extended easily.

# Directory structure

## Project root
```
applescript/          - Test AppleScript files for macOS PowerPoint automation
scripts/              - Utility scripts (cross-platform Python scripts)
  ├── bibconv.py      - Bible format conversion utility
  ├── msgupdate.py    - Update and compile translation files
  ├── osx_sample.py   - Sample script for macOS PowerPoint automation
  └── xgettext.py     - Extract translatable strings from source code
sample/               - Sample files and templates
  ├── service.sdf     - Sample service definition file (update pathnames before use)
  ├── Service-Template.pptx - Sample PowerPoint template
  └── service-notes-template.txt - Sample notes template
screenshot/           - Application screenshots
service_ppt/          - Main application package
tests/                - Unit tests
```

## service_ppt package structure
```
service_ppt/
  ├── bible/          - Bible text data structure and format readers
  │   ├── bibcore.py  - Core Bible data structures
  │   ├── biblang.py  - Bible language and translation support
  │   └── [format readers] - Support for MyBible, MySword, Sword, Zefania, etc.
  ├── hymn/           - Hymn and lyric data structure and readers
  │   └── [OpenLyrics support]
  ├── utils/          - General utility modules
  ├── wx_utils/       - wxPython-specific utility modules
  ├── image24/        - 24x24 pixel icons for toolbar and commands
  ├── image32/        - 32x32 pixel icons for toolbar and commands
  ├── locale/         - Language translation files (.mo, .po, .pot)
  │   └── [locale]/LC_MESSAGES/ - Compiled translations per locale
  ├── __main__.py     - Application entry point
  ├── mainframe.py    - Main window UI class
  ├── command.py      - Command classes for slide generation
  ├── command_ui.py   - UI classes for command configuration
  ├── powerpoint_base.py - Base classes for PowerPoint automation
  ├── powerpoint_win32.py - Windows COM automation backend
  ├── powerpoint_osx.py - macOS AppleScript automation backend
  └── powerpoint_pptx.py - Cross-platform python-pptx backend
```

## The structure of the application.
```
service_ppt/__main__.py - main application entry point (can be run as `python -m service_ppt` or `service-ppt` command).
service_ppt/mainframe.py - main window UI class.
service_ppt/command.py - The several **Command** classes to generate slides.
service_ppt/command_ui.py - The **CommandUI** classes that display UI for each Command classes.
service_ppt/powerpoint_base.py - The PresentationBase and PPTAppBase base classes that contain the common platform independent PowerPoint automation code
service_ppt/powerpoint_osx.py - The OS X specific classes using AppleScript to automate PowerPoint.
service_ppt/powerpoint_win32.py - The Windows specific classes using COM to automate PowerPoint.
service_ppt/powerpoint_pptx.py - Cross-platform implementation using python-pptx library.
```

# Service Definition File
The `.sdf` file contains the command instructions saved from each Command.<br>
It is a JSON file format internally.

# Building an Executable

To create a standalone executable, use PyInstaller (version 6.0.0 or higher for Python 3.12+ support).

## Install PyInstaller
```
pip install pyinstaller>=6.0.0
```

## Build the executable
```
pyinstaller service_ppt.spec
```

The executable will be created in the `dist/service_ppt/` directory.

**Note:** The `service_ppt.spec` file has been updated to work with the new package structure:
- Entry point: `service_ppt/__main__.py`
- Data directories: `service_ppt/image24`, `service_ppt/image32`, `service_ppt/locale`

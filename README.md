# What is service_ppt.
The service_ppt is an application generating PowerPoint slides for the Church's service.<br>
It produces Bible Verse, lyric and announcement slides based on template with find and replace text.<br>
The find and replace text can be applied to a text template file and can be saved as a text file for other purposes.<br>
It also can insert slides that are pre-made such as song slides.<br>
After it generates slide, it can export slides into images or shapes in slides into transparent images.<br>
<br>
All these operations are done using Windows Powerpoint COM service and that means you need to have PowerPoint installed on your machine.<br>
For OS X, it is using AppleScript but the feature is less mature than that of Windows.<br>

# Installation
This document describes how to set up the environment for the service_ppt application.

## Install Python 3.
Install Python 3.9 from https://www.python.org/downloads/.<br>
Add python3 to the PATH environment.

## Install packages using requirement.txt
All the below packages can be install by following commands.
Or each package can be installed one by one.
```
pip install -r requirement.txt  # For Python 3.9

pip install -r requirement-3.10.txt  # For Python 3.10
```

## Install wxPython
Install wxPython that is used for main GUI application.
```
pip install wxpython
```

## Install 3rd party packages.
These are common to Windows and Mac OS.
```
pip install iso639-lang

pip install langdetect

pip install numpy

pip install Pillow

pip install pysword
```

## Set up environment for language translation.
Install GNU gettext package and add the bin directory to the PATH environment.<br>
The 'xgettext', 'msgfmt' and 'msgmerge' will be used for translation strings.<br>
For Windows you can get the installer from http://gnuwin32.sourceforge.net/packages/gettext.htm.

## Windows platform
Windows specific packages.
```
pip install pywin32
```

## OS X platform
OS X specific packages.
```
pip install py-applescript
```

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
```
applescript - Test AppleScript files for OS X PowerPoint

bible - Bible text data structure and format reader.

hymn - Hymn and lyric data structure and reader.

image24, image32 - Icons/Images for Toolbar and command.

locale - Language translation.

sample - The sample service.sdf, service definition file that contains command instructions to generate slides file.
You need to update the pathnames in the file to be full path names before run it.
```

## The structure of the application.
```
service_ppt.pyw - main application start up code.
mainframe.py - main window UI class.
command.py - The several **Command** classes to generate slides.
command_ui.py - The **CommandUI** classes that display UI for each Command classes.
powerpoint_base.py - The PresentationBase and PPTAppBase base classes that contain the common platform independent PowerPoint automation code
powerpoint_osx.py - The OS X specific classes using AppleScript to automate PowerPoint.
powerpoint_win32.py - The Windows specific classes using COM to automate PowerPoint.
```

# Service Definition File
The `.sdf` file contains the command instructions saved from each Command.<br>
It is a JSON file format internally.

# Installer

Use `pyinstaller==5.13.2` to make an executable.
```
pyinstaller service_ppt.spec
```

@echo off

SET XGETTEXT="C:\Program Files\gettext-iconv\bin\xgettext.exe"

REM xgettext for Python does not parse pgettext by default, so add it.
%XGETTEXT% -kpgettext:1c,2  -d bible -o locale\bible.pot bible\bibcore.py bible\biblang.py

%XGETTEXT% -d service_ppt -o locale\service_ppt.pot mainframe.py preferences_dialog.py command.py command_ui.py

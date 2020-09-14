@ECHO OFF

REM This batch file requres GNU gettext package install on the path.

FOR %%D IN (ko) DO (
  msgmerge.exe -s -U locale\%%D\LC_MESSAGES\bible.po locale\bible.pot
  msgfmt.exe locale\%%D\LC_MESSAGES\bible.po -o locale\%%D\LC_MESSAGES\bible.mo
)


FOR %%D IN (ko) DO (
  msgmerge.exe -s -U locale\%%D\LC_MESSAGES\service_ppt.po locale\service_ppt.pot
  msgfmt.exe locale\%%D\LC_MESSAGES\service_ppt.po -o locale\%%D\LC_MESSAGES\service_ppt.mo
)

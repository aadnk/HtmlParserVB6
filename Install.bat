rem BUGFIX
chdir /D %~dp0

regsvr32 HTMLParser.dll
regsvr32 BrowserControl.ocx
pause
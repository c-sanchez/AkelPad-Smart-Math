@echo off
setlocal
set "FB_HOME=C:\tools\FreeBASIC"
if exist "%FB_HOME%\bin\win64" (
  set "PATH=%PATH%;%FB_HOME%;%FB_HOME%\bin\win64"
  set "FBC_ARCH=-arch x86_64"
) else (
  set "PATH=%PATH%;%FB_HOME%;%FB_HOME%\bin\win32"
  set "FBC_ARCH="
)
cd /d "%~dp0.."
fbc %FBC_ARCH% -s console -strip -O 2 %* "tools\calc.bas" "MathParser.bas" -x "tools\calc.exe"
exit /b %ERRORLEVEL%

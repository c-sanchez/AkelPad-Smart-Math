@echo off
setlocal EnableDelayedExpansion

rem Isolated factorint tests with a 20-second wall-clock limit per case.
rem Stops at the first case that exceeds the limit or returns non-zero.

set "FB_HOME=C:\tools\FreeBASIC"
if exist "%FB_HOME%\bin\win64" (
  set "PATH=%PATH%;%FB_HOME%;%FB_HOME%\bin\win64"
  set "FBC_ARCH=-arch x86_64"
) else (
  set "PATH=%PATH%;%FB_HOME%;%FB_HOME%\bin\win32"
  set "FBC_ARCH="
)

set "TIMEOUT_SEC=20"
set "CASE_COUNT=29"
set "EXE=SmokeTest_Factorint.exe"

echo ==========================================
echo factorint timeout tests (%TIMEOUT_SEC%s per case)
echo ==========================================
echo.

echo [1/2] Compiling %EXE% ...
fbc %FBC_ARCH% -strip -O 2 -Wc -O2 "tests\SmokeTest_Factorint.bas" "MathParser.bas" -x "%EXE%"
if errorlevel 1 goto :compile_failed

echo.
echo [2/2] Running cases with %TIMEOUT_SEC%s timeout ...
echo.

for /L %%I in (1,1,%CASE_COUNT%) do (
  echo --- Case %%I/%CASE_COUNT% ---
  powershell -NoProfile -Command ^
    "$p = Start-Process -FilePath '%EXE%' -ArgumentList '%%I' -PassThru -NoNewWindow; " ^
    "if (-not $p.WaitForExit(%TIMEOUT_SEC% * 1000)) { Stop-Process -Id $p.Id -Force; exit 124 }; " ^
    "exit $p.ExitCode"
  set "RC=!ERRORLEVEL!"
  if "!RC!"=="124" (
    echo.
    echo TIMEOUT: case %%I exceeded %TIMEOUT_SEC% seconds.
    echo Expression: see case list in tests\SmokeTest_Factorint.bas
    exit /b 124
  )
  if not "!RC!"=="0" (
    echo.
    echo FAILED: case %%I exited with code !RC!.
    exit /b !RC!
  )
  echo.
)

echo All %CASE_COUNT% factorint cases finished within %TIMEOUT_SEC%s each.
exit /b 0

:compile_failed
echo Compilation failed.
exit /b 1

@echo off
setlocal

rem Clipboard copy normalization tests (SmartMath_CopyNormalize.bas). Run from repo root.
set "FB_HOME=C:\tools\FreeBASIC"
set "PATH=%PATH%;%FB_HOME%"

echo ==========================================
echo SmartMath copy normalization tests
echo ==========================================
echo.
echo [1/2] Compiling CopyRegressionTests.exe ...
fbc "tests\FormatterTest_Globals.bas" "tests\CopyRegressionTests.bas" "SmartMath_CopyNormalize.bas" -x "CopyRegressionTests.exe"
if errorlevel 1 goto :compile_failed

echo.
echo [2/2] Running CopyRegressionTests.exe ...
echo.
"CopyRegressionTests.exe"
set "TEST_EXIT=%ERRORLEVEL%"

echo.
if not "%TEST_EXIT%"=="0" goto :tests_failed

echo Copy normalization tests PASSED.
exit /b 0

:compile_failed
echo.
echo Compilation failed.
exit /b 1

:tests_failed
echo Copy normalization tests FAILED (exit code %TEST_EXIT%).
exit /b %TEST_EXIT%

@echo off
setlocal

rem Formatter regression tests (SmartMath_Format.bas). Run from repo root.
set "FB_HOME=C:\tools\FreeBASIC"
set "PATH=%PATH%;%FB_HOME%"

echo ==========================================
echo SmartMath formatter regression tests
echo ==========================================
echo.
echo [1/4] Compiling FormatterRegressionTests.exe (unit) ...
fbc -d FORMATTER_TEST_BUILD -strip -O 2 -Wc -O2 "tests\FormatterTest_Globals.bas" "tests\FormatterRegressionTests.bas" "SmartMath_Format.bas" "MathParserRawResult.bas" -x "FormatterRegressionTests.exe"
if errorlevel 1 goto :compile_failed

echo.
echo [2/4] Running FormatterRegressionTests.exe ...
echo.
"FormatterRegressionTests.exe"
set "UNIT_EXIT=%ERRORLEVEL%"
if not "%UNIT_EXIT%"=="0" goto :tests_failed

echo.
echo [3/4] Compiling FormatterParserIntegrationTests.exe ...
fbc -d PARSER_FORMATTER_INTEGRATION -strip -O 2 -Wc -O2 "tests\FormatterTest_Globals.bas" "tests\FormatterRegressionTests.bas" "SmartMath_Format.bas" "MathParser.bas" -x "FormatterParserIntegrationTests.exe"
if errorlevel 1 goto :compile_failed

echo.
echo [4/4] Running FormatterParserIntegrationTests.exe ...
echo.
"FormatterParserIntegrationTests.exe"
set "TEST_EXIT=%ERRORLEVEL%"

echo.
if not "%TEST_EXIT%"=="0" goto :tests_failed

echo Formatter tests PASSED (unit + parser integration).
exit /b 0

:compile_failed
echo.
echo Compilation failed.
exit /b 1

:tests_failed
echo Formatter tests FAILED (exit code %TEST_EXIT%).
exit /b %TEST_EXIT%

@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"

set "TIMEOUT_SEC=20"
set "EXE=FactorintBench32.exe"

set VC_ROOT=%ProgramFiles%\Microsoft Visual Studio\2022
if exist "%VC_ROOT%\Professional\VC\Auxiliary\Build\vcvarsall.bat" (
  call "%VC_ROOT%\Professional\VC\Auxiliary\Build\vcvarsall.bat" x86 >nul
) else if exist "%VC_ROOT%\Community\VC\Auxiliary\Build\vcvarsall.bat" (
  call "%VC_ROOT%\Community\VC\Auxiliary\Build\vcvarsall.bat" x86 >nul
) else (
  echo ERROR: vcvarsall.bat not found
  exit /b 1
)

if exist %EXE% del /f %EXE% 2>nul
echo Compiling %EXE% ...
cl /nologo /O2 /EHsc FactorintBench.cpp MathParser.cpp MathParserFactorInt.cpp MathParserRawResult.cpp /Fe:%EXE%
if errorlevel 1 exit /b 1

echo Running %EXE% with %TIMEOUT_SEC%s timeout ...
powershell -NoProfile -Command ^
  "$p = Start-Process -FilePath '%EXE%' -WorkingDirectory '%CD%' -PassThru -NoNewWindow; " ^
  "if (-not $p.WaitForExit(%TIMEOUT_SEC% * 1000)) { Stop-Process -Id $p.Id -Force; Write-Host 'TIMEOUT after %TIMEOUT_SEC%s'; exit 124 }; " ^
  "exit $p.ExitCode"
set "RC=!ERRORLEVEL!"
if "!RC!"=="124" exit /b 124
exit /b !RC!

@echo off
rem C++-only combined verification (x64 + x86). Keep C++ actions inside cpp\ scripts.
cd /d "%~dp0"

echo [1/2] Running x64 C++ checks...
call "%~dp0RunTests_x64.bat"
if errorlevel 1 exit /b 1

echo [2/2] Running x86 C++ checks...
call "%~dp0RunTests_win32.bat"
if errorlevel 1 exit /b 1

echo All C++ checks PASSED.
exit /b 0

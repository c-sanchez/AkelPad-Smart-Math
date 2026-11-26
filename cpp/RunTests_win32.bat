@echo off
rem C++-only verification entrypoint (x86). Keep C++ actions inside cpp\ scripts.
cd /d "%~dp0"

call "%~dp0BuildTests_vc2022_win32.bat"
if errorlevel 1 exit /b 1

"%~dp0Tests_MathParser.exe"
if errorlevel 1 exit /b 1

echo C++ x86 checks PASSED.
exit /b 0

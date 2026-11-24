@echo off
rem C++-only verification entrypoint. Keep C++ actions inside cpp\ scripts.
cd /d "%~dp0"

call "%~dp0RunAnalyzeAndTests_vc2022_x64.bat"
if errorlevel 1 exit /b 1

echo C++ checks PASSED.
exit /b 0

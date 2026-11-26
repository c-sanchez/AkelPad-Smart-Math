@echo off
cd /d "%~dp0"
set VC_ROOT=%ProgramFiles%\Microsoft Visual Studio\2022
if exist "%VC_ROOT%\Community\VC\Auxiliary\Build\vcvarsall.bat" (
  call "%VC_ROOT%\Community\VC\Auxiliary\Build\vcvarsall.bat" x86 >nul
) else (
  call "%VC_ROOT%\Professional\VC\Auxiliary\Build\vcvarsall.bat" x86 >nul
)
if exist FactorintModTest.exe del FactorintModTest.exe
cl /nologo /O2 tests\FactorintModTest.cpp /Fe:FactorintModTest.exe
if errorlevel 1 exit /b 1
del /q FactorintModTest.obj >nul 2>&1
FactorintModTest.exe

@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set VC_ROOT=%ProgramFiles%\Microsoft Visual Studio\2022
set VCVARS_ARG=x86

if exist "%VC_ROOT%\Professional\VC\Auxiliary\Build\vcvarsall.bat" goto UseVcProfessional
if exist "%VC_ROOT%\Community\VC\Auxiliary\Build\vcvarsall.bat" goto UseVcCommunity

goto ErrorNoVcVarsAll

:UseVcProfessional
call "%VC_ROOT%\Professional\VC\Auxiliary\Build\vcvarsall.bat" %VCVARS_ARG%
goto Building

:UseVcCommunity
call "%VC_ROOT%\Community\VC\Auxiliary\Build\vcvarsall.bat" %VCVARS_ARG%
goto Building

:Building
cl /O2 /EHsc %* calc.cpp MathParser.cpp MathParserFactorInt.cpp MathParserRawResult.cpp /Fe:calc.exe
if errorlevel 1 exit /b 1
goto End

:ErrorNoVcVarsAll
echo ERROR: Could not find "vcvarsall.bat"
exit /b 1

:End
endlocal

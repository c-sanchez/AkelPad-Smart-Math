@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set VC_ROOT=%ProgramFiles%\Microsoft Visual Studio\2022
set VCVARS_ARG=amd64

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
set "CL=/DSMARTMATH_COMPLEX_NUMBERS=0 /DSMARTMATH_TIME_VALUES=0 /DSMARTMATH_LAMBDA_FUNCTIONS=0 /DSMARTMATH_FACTORINT=0"
cl /O2 /EHsc %* calc.cpp MathParser.cpp MathParserRawResult.cpp /Fe:calc_mini.exe
if errorlevel 1 exit /b 1
goto End

:ErrorNoVcVarsAll
echo ERROR: Could not find "vcvarsall.bat"
exit /b 1

:End
endlocal

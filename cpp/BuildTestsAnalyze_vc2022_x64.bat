@echo off
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
cl /std:c++17 /EHsc /W4 /analyze /wd4100 Tests_MathParser.cpp MathParser.cpp MathParserFactorInt.cpp MathParserRawResult.cpp /Fe:Tests_MathParser_analyze.exe
goto End

:ErrorNoVcVarsAll
echo ERROR: Could not find "vcvarsall.bat"
goto End

:End

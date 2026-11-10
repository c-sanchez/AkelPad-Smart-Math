@echo off
rem C++-only perf runner. Keep C++ actions inside cpp\ scripts.
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
set SMARTMATH_FLAGS=/DSMARTMATH_COMPLEX_NUMBERS=0 /DSMARTMATH_TIME_VALUES=0 /DSMARTMATH_LAMBDA_FUNCTIONS=0 /DSMARTMATH_FACTORINT=0
cl %SMARTMATH_FLAGS% /O2 /EHsc PerfHotPath_MathParser.cpp MathParser.cpp MathParserFactorInt.cpp MathParserRawResult.cpp /Fe:PerfHotPath_MathParser.exe
if errorlevel 1 exit /b 1

set ITERATIONS=%1
if "%ITERATIONS%"=="" set ITERATIONS=200000

set REPEATS=%2
if "%REPEATS%"=="" set REPEATS=3

set WARMUP=%3
if "%WARMUP%"=="" set WARMUP=1

set MODE=%4
if "%MODE%"=="" (
  echo Running perf with %ITERATIONS% iterations, %REPEATS% repeats, %WARMUP% warmup runs...
  echo Modes: lambda-stress ^| profile ^(compile/eval/raw split^)
  PerfHotPath_MathParser.exe %ITERATIONS% %REPEATS% %WARMUP%
) else (
  echo Running perf mode "%MODE%" with %ITERATIONS% iterations...
  PerfHotPath_MathParser.exe %ITERATIONS% %REPEATS% %WARMUP% %MODE%
)
if errorlevel 1 exit /b 1
exit /b 0

:ErrorNoVcVarsAll
echo ERROR: Could not find "vcvarsall.bat"
exit /b 1

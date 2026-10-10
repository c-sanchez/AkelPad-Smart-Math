@echo off

rem Specify the correct path to FreeBASIC here:
set "FB_HOME=C:\tools\FreeBASIC"

if exist "%FB_HOME%" goto UseWin32
echo ERROR: FreeBASIC not found in "%FB_HOME%". Please specify the correct path to FreeBASIC in "%~nx0".
exit /b 1

:UseWin32
if not exist "%FB_HOME%\bin\win32" goto UseWin64
set "PATH=%PATH%;%FB_HOME%;%FB_HOME%\bin\win32"
set GCC_FLAGS=
goto Compile

:UseWin64
if not exist "%FB_HOME%\bin\win64" goto ErrorNoBinWin32Win64
set "PATH=%PATH%;%FB_HOME%;%FB_HOME%\bin\win64"
set GCC_FLAGS=-arch x86_64
goto Compile

:ErrorNoBinWin32Win64
echo ERROR: FreeBASIC neither "%FB_HOME%\bin\win32" nor "%FB_HOME%\bin\win64" found.
exit /b 1

:Compile
echo Compiling...
fbc -dll -gen gcc %GCC_FLAGS% "SmartMath.bas" "SmartMath_Config.bas" "SmartMath_Format.bas" "SmartMath_CopyNormalize.bas" "SmartMath_About.bas" "SmartMath_Menu.bas" "MathParser.bas" "SmartMath.rc"
if exist "*.a" del "*.a"
if exist "*.o" del "*.o"
REM pause

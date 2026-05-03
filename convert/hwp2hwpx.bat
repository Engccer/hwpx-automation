@echo off
setlocal enabledelayedexpansion
set "JAVA_HOME=C:\Program Files\Eclipse Adoptium\jdk-21.0.10.7-hotspot"
set "TOOL_DIR=%~dp0"
if "%TOOL_DIR:~-1%"=="\" set "TOOL_DIR=%TOOL_DIR:~0,-1%"
set "JAR1=%TOOL_DIR%\hwp2hwpx-1.0.0.jar"
set "JAR2=%TOOL_DIR%\lib\hwplib-1.1.10.jar"
set "JAR3=%TOOL_DIR%\lib\hwpxlib-1.0.8.jar"
set "CP=%JAR1%;%JAR2%;%JAR3%;%TOOL_DIR%"

set "INPUT=%~1"
set "OUTPUT=%~2"

if "%INPUT%"=="" (
    echo Usage: %~nx0 ^<input.hwp^> [output.hwpx]
    exit /B 2
)

if "%OUTPUT%"=="" (
    set "INPUT_DIR=%~dp1"
    set "INPUT_NAME=%~n1"
    set "OUTPUT_DIR=%~dp1_output"
    if not exist "!OUTPUT_DIR!" mkdir "!OUTPUT_DIR!"
    set "OUTPUT=!OUTPUT_DIR!\!INPUT_NAME!.hwpx"
)

REM JVM decodes argv as cp949 on Windows; en-dash, em-dash, etc. break it.
REM Stage input/output to ASCII-safe %TEMP% to dodge FileNotFoundException.
set "STAGE=%TEMP%\hwp2hwpx_%RANDOM%_%RANDOM%"
mkdir "%STAGE%" >nul 2>&1
if not exist "%STAGE%" (
    echo ERROR: failed to create staging dir %STAGE% 1>&2
    exit /B 1
)

copy /Y "%INPUT%" "%STAGE%\input.hwp" >nul
if errorlevel 1 (
    echo ERROR: failed to stage input file 1>&2
    rmdir /S /Q "%STAGE%" >nul 2>&1
    exit /B 1
)

"%JAVA_HOME%\bin\java.exe" -cp "%CP%" Hwp2HwpxCLI "%STAGE%\input.hwp" "%STAGE%\output.hwpx"
set "RC=!ERRORLEVEL!"

if !RC! EQU 0 (
    if exist "%STAGE%\output.hwpx" (
        copy /Y "%STAGE%\output.hwpx" "!OUTPUT!" >nul
        if errorlevel 1 (
            echo ERROR: failed to copy output to !OUTPUT! 1>&2
            set "RC=1"
        )
    ) else (
        echo ERROR: Java succeeded but output file missing 1>&2
        set "RC=1"
    )
)

rmdir /S /Q "%STAGE%" >nul 2>&1
exit /B !RC!

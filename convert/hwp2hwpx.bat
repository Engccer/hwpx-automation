@echo off
setlocal enabledelayedexpansion
set JAVA_HOME=C:\Program Files\Eclipse Adoptium\jdk-21.0.10.7-hotspot
set TOOL_DIR=C:\Users\pc\Converters\hwpx-tool\convert
set JAR1=%TOOL_DIR%\hwp2hwpx-1.0.0.jar
set JAR2=%TOOL_DIR%\lib\hwplib-1.1.10.jar
set JAR3=%TOOL_DIR%\lib\hwpxlib-1.0.8.jar
set CP=%JAR1%;%JAR2%;%JAR3%;%TOOL_DIR%

set INPUT=%~1
set OUTPUT=%~2

if "%OUTPUT%"=="" (
    set "INPUT_DIR=%~dp1"
    set "INPUT_NAME=%~n1"
    set "OUTPUT_DIR=%~dp1_output"
    if not exist "!OUTPUT_DIR!" mkdir "!OUTPUT_DIR!"
    set "OUTPUT=!OUTPUT_DIR!\!INPUT_NAME!.hwpx"
)

"%JAVA_HOME%\bin\java.exe" -cp "%CP%" Hwp2HwpxCLI "%INPUT%" "%OUTPUT%"

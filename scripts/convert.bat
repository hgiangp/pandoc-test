@echo off
rem Rasterize Word drawings in a .docx so pandoc can extract them.
rem Usage:   convert.bat input.docx [options]
rem Example: convert.bat input.docx -DryRun
rem          convert.bat input.docx -Dpi 300 -IncludeTextBoxes
rem Drag-and-drop a .docx onto this file also works.
setlocal EnableExtensions

if "%~1"=="" goto usage
if not exist "%~f1" goto notfound

set "INPUT=%~f1"
shift
set "EXTRA="
:collect
if "%~1"=="" goto run
set "EXTRA=%EXTRA% %1"
shift
goto collect

:run
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Convert-ShapesToPictures.ps1" -InputPath "%INPUT%"%EXTRA%
set "RC=%ERRORLEVEL%"
if "%RC%"=="0" echo [OK] Done.
if "%RC%"=="2" echo [WARN] Some drawings failed to convert - see manifest.csv, column Error.
if not "%RC%"=="0" if not "%RC%"=="2" echo [ERROR] Script failed with exit code %RC%.
if not defined NOPAUSE pause
exit /b %RC%

:notfound
echo [ERROR] File not found: "%~f1"
goto fail

:usage
echo Usage: %~nx0 input.docx [-DryRun] [-Dpi 300] [-IncludeTextBoxes] [-KeepMetafiles] [-NoCluster] [-NoTrim] [-Visible]
:fail
if not defined NOPAUSE pause
exit /b 1

@echo off
rem Full pipeline: preflight -> data cleanup -> drawings to PNG -> pandoc -> gate.
rem All logic lives in Invoke-Pipeline.ps1; this file only makes it easy to launch.
rem
rem Usage:   run-all.bat input.docx [options]
rem Options: -Profile ns              data cleanup profile (profiles\ns.json); default: no cleanup
rem          -CleanupMode Report      cleanup only reports, changes nothing
rem          -NoCleanup               skip data cleanup
rem          -NoFigureFilter          run pandoc without pandoc\figures.lua
rem          -OutputFormat markdown   Pandoc Markdown instead of gfm (default)
rem          -Dpi 300 -IncludeTextBoxes -KeepMetafiles -NoCluster -NoTrim -Visible
rem Output next to input.docx: input.clean.docx, input.shapes.docx, input.md, images\,
rem          input.cleanup-manifest.csv, input.shapes\manifest.csv, input.pipeline.log
rem Exit code: 0 = pass, 3 = finished with issues, 1 = error. Set NOPAUSE=1 for unattended runs.
setlocal EnableExtensions
rem Save the script folder now: SHIFT also shifts %0, so %~dp0 is wrong after it
set "SCRIPTS=%~dp0"

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
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPTS%Invoke-Pipeline.ps1" -InputPath "%INPUT%"%EXTRA%
set "RC=%ERRORLEVEL%"
if not defined NOPAUSE pause
exit /b %RC%

:notfound
echo [ERROR] File not found: "%~f1"
goto fail

:usage
echo Usage: %~nx0 input.docx [-Profile ns] [-CleanupMode Report] [-NoCleanup] [-NoFigureFilter] [-OutputFormat markdown] [-Dpi 300] [-IncludeTextBoxes] [-KeepMetafiles] [-NoCluster] [-NoTrim] [-Visible]
:fail
if not defined NOPAUSE pause
exit /b 1

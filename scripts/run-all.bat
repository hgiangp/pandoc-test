@echo off
rem Full pipeline: preflight -> cleanup -> drawings to PNG -> pandoc prep -> pandoc -> media -> gate.
rem All logic lives in Invoke-Pipeline.ps1; this file only makes it easy to launch.
rem
rem Usage:   run-all.bat input.docx [options]
rem Options: -Profile ns              data cleanup profile (profiles\ns.json); default: no cleanup
rem          -CleanupMode Report      cleanup only reports, changes nothing
rem          -NoCleanup               skip data cleanup
rem          -NoPandocPrep            skip the pandoc compatibility fixes (profiles\pandoc.json)
rem          -NoFigureFilter          run pandoc without pandoc\figures.lua
rem          -NoMediaConvert          keep extracted EMF/WMF instead of converting them to PNG
rem          -OutputFormat markdown   Pandoc Markdown instead of gfm (default)
rem          -UnlinkShapeFields       fields in shapes to plain text (if images show "Error! Reference source...")
rem          -Dpi 300 -IncludeTextBoxes -KeepMetafiles -NoCluster -NoTrim -Visible
rem Output next to input.docx: input.clean.docx, input.shapes.docx, input.pandoc.docx, input.md, images\,
rem          input.cleanup-manifest.csv, input.shapes\manifest.csv, input.pandoc-manifest.csv, input.pipeline.log
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
echo Usage: %~nx0 input.docx [-Profile ns] [-CleanupMode Report] [-NoCleanup] [-NoPandocPrep] [-NoFigureFilter] [-NoMediaConvert] [-OutputFormat markdown] [-Dpi 300] [-IncludeTextBoxes] [-KeepMetafiles] [-NoCluster] [-NoTrim] [-Visible] [-UnlinkShapeFields]
:fail
if not defined NOPAUSE pause
exit /b 1

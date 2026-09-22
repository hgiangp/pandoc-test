@echo off
rem Full pipeline: preflight -> convert drawings -> pandoc -> gate.
rem Usage:   run-all.bat input.docx [options for Convert-ShapesToPictures.ps1]
rem Output (next to input.docx):
rem   input.shapes.docx   docx with drawings replaced by PNG
rem   input.shapes\       rendered PNG/EMF + manifest.csv
rem   input.md            pandoc markdown
rem   images\media\       images extracted by pandoc
rem Set NOPAUSE=1 to run unattended (CI).
setlocal EnableExtensions

if "%~1"=="" goto usage
if not exist "%~f1" goto notfound

set "SCRIPTS=%~dp0"
set "INPUT=%~f1"
set "WORKDIR=%~dp1"
set "NAME=%~n1"
set "SHAPES=%WORKDIR%%NAME%.shapes.docx"
set "MD=%NAME%.md"
set "PS=powershell -NoProfile -ExecutionPolicy Bypass -File"

shift
set "EXTRA="
:collect
if "%~1"=="" goto check
set "EXTRA=%EXTRA% %1"
shift
goto collect

:check
where pandoc >nul 2>nul
if errorlevel 1 (
    echo [ERROR] pandoc not found in PATH. Install: winget install --id JohnMacFarlane.Pandoc
    goto fail
)

echo.
echo ===== [1/4] Preflight: drawings pandoc would drop =====
%PS% "%SCRIPTS%Test-DocxDrawings.ps1" -Docx "%INPUT%"

echo.
echo ===== [2/4] Convert drawings to PNG with Word =====
%PS% "%SCRIPTS%Convert-ShapesToPictures.ps1" -InputPath "%INPUT%" -OutputPath "%SHAPES%"%EXTRA%
set "RC=%ERRORLEVEL%"
if "%RC%"=="2" echo [WARN] Some drawings failed to convert - see %NAME%.shapes\manifest.csv
if not "%RC%"=="0" if not "%RC%"=="2" (
    echo [ERROR] Conversion failed with exit code %RC%.
    goto fail
)
if not exist "%SHAPES%" goto noshapes

echo.
echo ===== [3/4] pandoc =====
rem Run inside the input folder so image links in the markdown are relative (images/media/...)
pushd "%WORKDIR%"
pandoc -f docx -t markdown --extract-media=./images "%SHAPES%" -o "%MD%"
set "PRC=%ERRORLEVEL%"
popd
if not "%PRC%"=="0" (
    echo [ERROR] pandoc failed with exit code %PRC%.
    goto fail
)
echo Written: %WORKDIR%%MD%

echo.
echo ===== [4/4] Gate: converted docx + markdown =====
%PS% "%SCRIPTS%Test-DocxDrawings.ps1" -Docx "%SHAPES%" -Markdown "%WORKDIR%%MD%"
set "GRC=%ERRORLEVEL%"

echo.
if "%GRC%"=="0" if "%RC%"=="0" (
    echo [PASS] All drawings converted, no drawing objects left for pandoc to drop.
    if not defined NOPAUSE pause
    exit /b 0
)
echo [CHECK] Pipeline finished with issues: convert=%RC%, gate=%GRC%. Review the output above and manifest.csv.
if not defined NOPAUSE pause
exit /b 3

:noshapes
echo [ERROR] Converted file was not created: "%SHAPES%"
echo         Note: -DryRun cannot be used with run-all.bat, use convert.bat instead.
goto fail

:notfound
echo [ERROR] File not found: "%~f1"
goto fail

:usage
echo Usage: %~nx0 input.docx [-Dpi 300] [-IncludeTextBoxes] [-KeepMetafiles] [-NoCluster] [-NoTrim] [-Visible]
:fail
if not defined NOPAUSE pause
exit /b 1

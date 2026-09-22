@echo off
rem Data cleanup only. Output: input.clean.docx + input.clean.cleanup-manifest.csv
rem Usage:   cleanup.bat input.docx [--profile ns] [--mode Report] [-o out.docx] [--list-rules]
rem Needs uv (winget install --id astral-sh.uv -e); the first run installs Python + dependencies.
setlocal EnableExtensions

where uv >nul 2>nul || goto nouv
uv run --project "%~dp0cleanup" --locked --no-dev docx-cleanup %*
set "RC=%ERRORLEVEL%"
if not defined NOPAUSE pause
exit /b %RC%

:nouv
echo [ERROR] uv not found in PATH. Install: winget install --id astral-sh.uv -e
if not defined NOPAUSE pause
exit /b 1

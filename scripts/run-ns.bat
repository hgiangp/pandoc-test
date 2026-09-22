@echo off
rem Full pipeline with the NS data cleanup profile. Drag-and-drop a .docx onto this file.
rem Usage: run-ns.bat input.docx [more options for run-all.bat]
call "%~dp0run-all.bat" %* -Profile ns
exit /b %ERRORLEVEL%

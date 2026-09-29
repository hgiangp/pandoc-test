@echo off
rem Full pipeline with the review-markup data cleanup profile (reviewer boxes, layout tables). Drag-and-drop a .docx onto this file.
rem Usage: run-review-markup.bat input.docx [more options for run-all.bat]
call "%~dp0run-all.bat" %* -Profile review-markup
exit /b %ERRORLEVEL%

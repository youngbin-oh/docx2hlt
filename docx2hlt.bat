@echo off
rem docx2hlt.bat input.docx [output.hlt] [more.docx ...]
rem   output omitted -> <input>.hlt written next to each input
if "%~1"=="" (echo usage: docx2hlt.bat input.docx [output.hlt] [more.docx ...] & exit /b 2)
setlocal
rem the bundled embeddable Python (python\) first, so no install is needed
set "PY=python"
if exist "%~dp0python\python.exe" set "PY=%~dp0python\python.exe"
"%PY%" "%~dp0docx2hlt.py" %*

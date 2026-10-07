@echo off
rem docx2hlt.bat input.docx [output.hlt] [more.docx ...]
rem   output omitted -> <input>.hlt written next to each input
if "%~1"=="" (echo usage: docx2hlt.bat input.docx [output.hlt] [more.docx ...] & exit /b 2)
python "%~dp0docx2hlt.py" %*

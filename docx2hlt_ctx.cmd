@echo off
rem Right-click context-menu wrapper for docx2hlt.
rem Receives one or more .docx paths (Explorer passes every selected file
rem when the verb has MultiSelectModel=Document). Each writes <name>.hlt
rem next to its input, sequentially.
setlocal
title docx2hlt
echo Converting:
for %%A in (%*) do echo   %%~A
echo.
rem --pause-on-review-error: python itself waits for Enter when the .hlt was
rem made but its .rtf / .compare.docx review files were not
python "%~dp0docx2hlt.py" --pause-on-review-error %*
if errorlevel 1 (
  echo.
  echo *** CONVERSION FAILED ***
  pause
  exit /b 1
)
echo.
echo Done.
rem keep window briefly so the result is visible, then auto-close
ping -n 4 127.0.0.1 >nul
endlocal

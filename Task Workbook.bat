@echo off
setlocal

set "ROOT=%~dp0"
set "EXE=%ROOT%dist\Task Workbook.exe"

if exist "%EXE%" (
    start "" "%EXE%"
    endlocal
    exit /b 0
)

set "PYTHON=%ROOT%.venv\Scripts\pythonw.exe"

if exist "%PYTHON%" (
    start "" "%PYTHON%" "%ROOT%app.py"
) else (
    pythonw "%ROOT%app.py"
)

endlocal
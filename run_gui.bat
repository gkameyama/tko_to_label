@echo off
setlocal
cd /d "%~dp0"

where python.exe >nul 2>nul
if "%ERRORLEVEL%"=="0" (
    python.exe "%~dp0tko_to_label_xlsx.py" --gui
) else (
    where py.exe >nul 2>nul
    if "%ERRORLEVEL%"=="0" (
        py.exe -3 "%~dp0tko_to_label_xlsx.py" --gui
    ) else (
        echo Python was not found.
        set "EXITCODE=1"
        goto end
    )
)

set "EXITCODE=%ERRORLEVEL%"
:end
echo.
if "%EXITCODE%"=="0" (
    echo Done.
) else (
    echo Failed. Exit code: %EXITCODE%
)
pause
exit /b %EXITCODE%

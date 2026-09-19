@echo off
rem Export all lab pre-exam questions to .private\exam\. Double-click to run.
rem
rem ASCII-only on purpose: this machine's console code page is 936 (GBK), and
rem mixing UTF-8 text with "chcp 65001" makes cmd mis-parse later lines
rem (observed: "set /p OVERWRITE=" got truncated to "TE", then "%OVERWRITE%"
rem never expanded).  The Chinese instructions live in the exported .txt files,
rem which Python reads as UTF-8 and cmd never touches.
setlocal
cd /d "%~dp0.."

echo.
echo === Lab pre-exam: export questions ===
echo.
echo Starts the project's own Chromium (already signed in) and writes every
echo optional subject's questions into .private\exam\.
echo Read-only: nothing is submitted by this step.
echo.

if exist ".private\exam" (
    set /p OVERWRITE=".private\exam\ already exists; files will be overwritten. Continue? [y/N] "
    if /i not "%OVERWRITE%"=="y" (
        echo Cancelled.
        goto :done
    )
    echo.
)

uv run course-selection lab-exam --center dxwl --all-subjects --out-dir .private\exam
if errorlevel 1 (
    echo.
    echo [ERROR] Export failed.
    echo         If the browser session expired, start the workbench once and
    echo         sign in to openlab again, then retry.
    goto :done
)

echo.
echo Done. Open any ^<ID^>.txt inside .private\exam\ and follow the rules at
echo the top of the file, then double-click lab-exam-submit.cmd.

:done
echo.
pause

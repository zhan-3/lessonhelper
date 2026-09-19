@echo off
rem Check and submit answers found in .private\exam\. Double-click to run.
rem
rem Deliberately two-stage: a dry run first (validates, writes nothing) so you
rem can see exactly what would be submitted, then a confirmation prompt.  The
rem submit stage sends one request per subject and stops at the first outcome
rem that is not a confirmed success.
rem
rem ASCII-only on purpose; see lab-exam-export.cmd for why.
setlocal
cd /d "%~dp0.."

echo.
echo === Lab pre-exam: submit answers ===
echo.

if not exist ".private\exam" (
    echo [ERROR] .private\exam\ not found.
    echo         Run lab-exam-export.cmd first, then fill in the answers.
    goto :done
)

echo [1/2] Dry run - validates only, submits nothing...
echo.
uv run course-selection lab-exam --center dxwl --all-subjects --answers-dir .private\exam
if errorlevel 1 (
    echo.
    echo [ERROR] Validation failed; nothing was submitted.
    goto :done
)

echo.
echo Only the subjects listed above as validated will be submitted.
echo Anything reported as not ready needs its answers completed first.
echo.
set /p CONFIRM="Submit those subjects now? [y/N] "
if /i not "%CONFIRM%"=="y" (
    echo.
    echo Cancelled; nothing was submitted.
    goto :done
)

echo.
echo [2/2] Submitting - one request per subject, no retry...
echo.
uv run course-selection lab-exam --center dxwl --all-subjects --answers-dir .private\exam --confirm all
if errorlevel 1 (
    echo.
    echo [NOTE] A result was not confirmed. If it reports an unknown outcome,
    echo        verify it on the openlab page before running this again.
    goto :done
)

echo.
echo All subjects submitted successfully.

:done
echo.
pause

@echo off
rem Launch an isolated Chrome with the CDP debugging port open, for observing
rem real openlab requests.
rem
rem Why a dedicated user-data-dir is mandatory: if Chrome is already running,
rem a new process hands its arguments to the existing instance and exits, so
rem --remote-debugging-port would silently do nothing.  A separate profile also
rem keeps your normal browser session untouched.
rem
rem Output is deliberately ASCII-only: this machine's console code page is 936
rem (GBK), and UTF-8 text in a .cmd file corrupts the following %VAR% expansions.
rem
rem Usage: double-click this file, sign in to openlab in the window that opens
rem (captcha included), then run the lab-exam / lab-contract commands.
setlocal

set "PROFILE=%LOCALAPPDATA%\openlab-observe"
set "CHROME=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME%" set "CHROME=%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"

if not exist "%CHROME%" (
    echo [ERROR] Chrome not found. Edit CHROME in this file and retry.
    echo         Expected: %%ProgramFiles%%\Google\Chrome\Application\chrome.exe
    pause
    exit /b 1
)

echo Browser : %CHROME%
echo Profile : %PROFILE%
echo CDP     : http://127.0.0.1:9222
echo.
echo Sign in to openlab in the window that opens, then leave it running.
echo.

start "" "%CHROME%" --remote-debugging-port=9222 --user-data-dir="%PROFILE%" "http://openlab.hitwh.edu.cn/dxwl/booking/"
exit /b 0

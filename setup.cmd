@echo off
setlocal

where uv >nul 2>nul
if errorlevel 1 (
    echo 未找到 uv，请先安装：https://docs.astral.sh/uv/getting-started/installation/
    exit /b 1
)

echo 正在安装项目依赖...
uv sync
if errorlevel 1 exit /b 1

echo 正在安装 Playwright Chromium 浏览器内核...
uv run playwright install chromium
if errorlevel 1 exit /b 1

echo.
echo 安装完成。首次使用请运行：
echo   uv run course-selection configure-login
echo   uv run course-selection configure-profile --grade 2025
pause

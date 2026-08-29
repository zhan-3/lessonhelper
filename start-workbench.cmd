@echo off
setlocal
cd /d "%~dp0"

set "PYTHONW=.venv\Scripts\pythonw.exe"

if not exist "%PYTHONW%" (
    where uv >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] 未找到 uv，无法创建运行环境。
        echo 请先安装 uv：https://docs.astral.sh/uv/getting-started/installation/
        pause
        exit /b 1
    )

    echo 首次启动，正在安装工作台依赖...
    uv sync
    if errorlevel 1 (
        echo [ERROR] 工作台依赖安装失败。
        pause
        exit /b 1
    )
)

start "选课工作台" "%PYTHONW%" -m course_selection workbench --port 5000
exit /b 0

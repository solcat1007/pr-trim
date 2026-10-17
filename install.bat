@echo off
chcp 65001 >nul
title PR瘦身 v1.0 - 安装脚本
echo ==========================================
echo   PR瘦身 v1.0 — 安装脚本
echo ==========================================
echo.
echo 检查 Python 环境...

where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [错误] 未检测到 Python，请先安装 Python 3.10+
    echo 下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
)

for /f "tokens=2" %%V in ('python --version 2^>^&1') do set PYVER=%%V
echo [信息] 检测到 Python %PYVER%

echo.
echo 本工具仅使用 Python 标准库，无需安装第三方依赖。
echo 可直接运行：python pr_trim.py
echo.

echo 正在启动 PR瘦身 v1.0...
start pythonw pr_trim.py
echo 已启动！
pause

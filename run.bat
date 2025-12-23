@echo off
chcp 65001 >nul
echo ========================================
echo    送货单对账单生成工具
echo ========================================
echo.
echo 正在启动程序...
echo.

uv run simple_gui.py

pause

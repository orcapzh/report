@echo off
chcp 65001 >nul
echo ========================================
echo   送货单对账单工具 - 一键制作安装程序
echo ========================================
echo.

REM 检查Inno Setup是否安装
if not exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" (
    echo [错误] 未检测到Inno Setup
    echo.
    echo 请先安装Inno Setup:
    echo 下载地址: https://jrsoftware.org/isdl.php
    echo.
    pause
    exit /b 1
)

echo [步骤 1/3] 安装依赖...
uv sync
if errorlevel 1 (
    echo [错误] 依赖安装失败
    pause
    exit /b 1
)

echo.
echo [步骤 2/3] 打包程序为EXE...
uv run pyinstaller --name="送货单对账单工具" --onefile --add-data "merge_delivery_orders.py;." --icon=NONE --clean simple_gui.py
if errorlevel 1 (
    echo [错误] 打包失败
    pause
    exit /b 1
)

echo.
echo [步骤 3/3] 创建安装程序...
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
if errorlevel 1 (
    echo [错误] 安装程序创建失败
    pause
    exit /b 1
)

echo.
echo ========================================
echo   制作完成！
echo ========================================
echo.
echo 安装程序位置: installer_output\送货单对账单工具_Setup_v1.0.exe
echo.
echo 你可以将此安装程序分发给其他人使用
echo.

REM 打开输出文件夹
if exist "installer_output" (
    explorer installer_output
)

pause

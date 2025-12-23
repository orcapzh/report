#!/bin/bash
# 送货单对账单生成工具 - 打包脚本

echo "开始打包送货单对账单生成工具..."

# 安装依赖
echo "安装依赖..."
uv sync

# 使用PyInstaller打包GUI版本
echo "使用PyInstaller打包GUI版本..."
uv run pyinstaller --name="送货单对账单工具-GUI" \
    --windowed \
    --onefile \
    --add-data "merge_delivery_orders.py:." \
    --icon=NONE \
    --clean \
    gui_app.py

echo "打包完成！"
echo "可执行文件位置: dist/送货单对账单工具"

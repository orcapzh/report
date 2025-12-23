#!/bin/bash
# 送货单对账单生成工具 - 打包脚本（简化版）

echo "开始打包送货单对账单生成工具（简化版）..."

# 安装依赖
echo "安装依赖..."
uv sync

# 使用PyInstaller打包
echo "使用PyInstaller打包..."
uv run pyinstaller --name="送货单对账单工具" \
    --onefile \
    --add-data "merge_delivery_orders.py:." \
    --icon=NONE \
    --clean \
    simple_gui.py

echo "打包完成！"
echo "可执行文件位置: dist/送货单对账单工具"
echo ""
echo "说明：这是命令行交互版本，无需图形界面支持"

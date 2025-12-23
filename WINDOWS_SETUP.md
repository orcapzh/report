# Windows用户使用说明

## 方式一：直接运行（推荐，最简单）

### 步骤1：安装uv包管理器

1. 打开PowerShell（管理员权限）
2. 运行以下命令：
```powershell
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

### 步骤2：运行程序

1. 将整个项目文件夹解压到电脑上
2. 双击运行 `run.bat`（如果没有这个文件，见下方）
3. 或在项目文件夹中打开命令提示符，运行：
```batch
uv run simple_gui.py
```

### 步骤3：使用程序

1. 程序启动后，按提示输入原始数据文件夹路径（例如：`C:\Users\YourName\Desktop\raw-data`）
2. 输入输出文件夹路径（例如：`C:\Users\YourName\Desktop\output`）
3. 按 `y` 确认开始处理
4. 等待处理完成，查看生成的文件

---

## 方式二：打包成EXE（可分发给其他人）

如果需要分发给没有安装Python的电脑，可以打包成EXE：

### 步骤1：安装uv（同上）

### 步骤2：打包

双击运行 `build_simple.bat`

### 步骤3：获取可执行文件

打包完成后，在 `dist` 文件夹找到 `送货单对账单工具.exe`

这个EXE文件可以复制到任何Windows电脑上直接运行，无需安装Python。

---

## 方式三：图形界面版（需要额外设置）

如果需要完整的图形界面：

1. 确保安装了Python 3.13
2. 安装时确保勾选了"tcl/tk"组件
3. 运行 `uv run gui_app.py`
4. 或使用 `build.bat` 打包

---

## 文件说明

- `simple_gui.py` - 简化版程序（推荐）
- `gui_app.py` - 图形界面版
- `merge_delivery_orders.py` - 核心处理逻辑
- `build_simple.bat` - 打包脚本（简化版）
- `build.bat` - 打包脚本（GUI版）
- `run.bat` - 快速运行脚本
- `raw-data/` - 放置原始送货单Excel文件的文件夹
- `output/` - 生成的对账单保存位置

---

## 常见问题

### 问题1：提示找不到uv命令

**解决方案**：
1. 重新打开命令提示符（确保之前安装uv后重新打开）
2. 或者手动设置环境变量，将uv安装路径添加到PATH

### 问题2：打包失败

**解决方案**：
1. 确保已安装uv：`uv --version`
2. 运行 `uv sync` 安装所有依赖
3. 再次运行打包脚本

### 问题3：程序运行出错

**解决方案**：
1. 检查原始数据文件夹路径是否正确
2. 确保Excel文件格式正确
3. 查看错误信息，按提示操作

---

## 技术支持

如有问题，请查看程序运行时的错误提示，或联系技术支持。

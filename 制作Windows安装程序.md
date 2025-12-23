# 制作Windows安装程序完整指南

本指南将帮助你在Windows电脑上创建一个专业的安装程序（.exe），用户可以像安装普通软件一样安装使用。

## 准备工作（在Windows电脑上操作）

### 第一步：安装必要工具

1. **安装uv包管理器**
   - 打开PowerShell（管理员权限）
   - 运行：
     ```powershell
     powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
     ```
   - 重新打开命令提示符

2. **安装Inno Setup**（安装程序制作工具）
   - 下载地址：https://jrsoftware.org/isdl.php
   - 下载 `innosetup-6.x.x.exe`（选择最新版本）
   - 双击安装，一路默认即可
   - **重要**：安装时选择"中文语言包"

## 制作步骤

### 步骤1：准备项目文件

1. 将整个项目文件夹复制到Windows电脑
2. 在项目文件夹中打开命令提示符
3. 确保有以下文件：
   - `simple_gui.py`
   - `merge_delivery_orders.py`
   - `build_simple.bat`
   - `installer.iss`（安装程序脚本）

### 步骤2：打包成EXE

1. 在项目文件夹打开命令提示符
2. 运行打包命令：
   ```batch
   build_simple.bat
   ```
3. 等待打包完成
4. 在 `dist` 文件夹会生成 `送货单对账单工具.exe`

### 步骤3：创建安装程序

1. 确保项目根目录有以下内容：
   - `dist\送货单对账单工具.exe` （第2步生成的）
   - `installer.iss` （安装程序脚本）
   - `使用说明.txt`
   - `README.md`
   - `raw-data\` 文件夹（可以是空的）

2. 右键点击 `installer.iss` 文件
3. 选择"Compile"（编译）
4. Inno Setup会自动编译生成安装程序

### 步骤4：获取安装程序

编译完成后，在 `installer_output` 文件夹找到：
- `送货单对账单工具_Setup_v1.0.exe`

这就是最终的安装程序！

## 安装程序功能特性

用户运行安装程序后会：

✅ **安装到程序文件夹**
- 默认路径：`C:\Program Files\送货单对账单工具`
- 用户可自定义安装路径

✅ **创建开始菜单快捷方式**
- 送货单对账单工具
- 使用说明
- 卸载程序

✅ **可选桌面快捷方式**
- 安装时可选择是否创建桌面图标

✅ **完整卸载功能**
- 通过"添加或删除程序"卸载
- 或通过开始菜单卸载
- 自动删除生成的临时文件

## 安装程序文件结构

安装后的文件夹结构：
```
C:\Program Files\送货单对账单工具\
├── 送货单对账单工具.exe    # 主程序
├── 使用说明.txt             # 使用说明
├── README.md                # 详细文档
├── raw-data\                # 示例文件夹
└── output\                  # 自动创建的输出文件夹
```

## 用户使用流程

1. **安装**
   - 双击 `送货单对账单工具_Setup_v1.0.exe`
   - 按提示选择安装路径
   - 选择是否创建桌面快捷方式
   - 等待安装完成

2. **使用**
   - 从开始菜单或桌面启动程序
   - 按提示输入文件夹路径
   - 查看生成的对账单

3. **卸载**
   - 通过"设置 > 应用 > 应用和功能"卸载
   - 或通过开始菜单的卸载快捷方式

## 分发安装程序

### 方式1：直接分发

将 `送货单对账单工具_Setup_v1.0.exe` 通过以下方式分发：
- 邮件附件
- U盘
- 网盘分享
- 内部文件服务器

### 方式2：制作压缩包

创建一个压缩包包含：
```
送货单对账单工具_安装包.zip
├── 送货单对账单工具_Setup_v1.0.exe
└── 安装说明.txt
```

`安装说明.txt` 内容：
```
送货单对账单工具 - 安装说明

1. 双击运行"送货单对账单工具_Setup_v1.0.exe"
2. 按提示完成安装
3. 从开始菜单启动程序
4. 详细使用方法请查看程序中的使用说明

系统要求：
- Windows 10/11
- 无需安装Python或其他依赖
```

## 常见问题

### Q: 安装时提示"Windows已保护你的电脑"？
**A:** 这是Windows SmartScreen的安全提示。
- 点击"更多信息"
- 点击"仍要运行"
- 原因：程序没有数字签名（需要购买证书）

### Q: 如何添加数字签名？
**A:** 需要购买代码签名证书（约$100-300/年），然后使用签名工具：
```batch
signtool sign /f "证书.pfx" /p "密码" /t http://timestamp.digicert.com "安装程序.exe"
```

### Q: 安装程序太大？
**A:** PyInstaller打包的EXE通常20-50MB，这是正常的。可以使用UPX压缩：
```batch
pyinstaller --onefile --upx-dir=upx-path simple_gui.py
```

### Q: 想自定义安装程序图标？
**A:**
1. 准备一个 `.ico` 图标文件
2. 修改 `installer.iss` 中的配置
3. 修改 `build_simple.bat`，添加 `--icon=你的图标.ico`

## 更新版本

发布新版本时：

1. 修改代码后重新打包EXE
2. 修改 `installer.iss` 中的版本号：
   ```
   #define MyAppVersion "1.1"
   ```
3. 重新编译安装程序
4. 分发新的安装程序

用户可以直接安装新版本覆盖旧版本。

## 技术支持

如果在制作过程中遇到问题：
1. 检查所有文件是否齐全
2. 确保已正确打包EXE
3. 查看Inno Setup的编译日志
4. 参考Inno Setup官方文档：https://jrsoftware.org/ishelp/

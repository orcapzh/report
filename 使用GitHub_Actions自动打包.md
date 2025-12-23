# 使用GitHub Actions自动打包多平台程序

## 什么是GitHub Actions？

GitHub Actions是GitHub提供的**免费CI/CD服务**，可以在云端的虚拟机上自动执行任务。

### 简单理解

想象GitHub有这样一个机房：

```
┌─────────────────────────────────────┐
│     GitHub数据中心（云端）          │
│                                     │
│  ┌─────────┐  ┌─────────┐  ┌──────┐│
│  │Windows  │  │ macOS   │  │Linux ││
│  │虚拟机   │  │虚拟机   │  │虚拟机││
│  │         │  │         │  │      ││
│  │运行你的 │  │运行你的 │  │运行你││
│  │构建脚本 │  │构建脚本 │  │的脚本││
│  └─────────┘  └─────────┘  └──────┘│
└─────────────────────────────────────┘
         ↓           ↓          ↓
    生成.exe    生成.app   生成Linux版
```

当你推送代码，GitHub会：
1. 同时启动3个真实的操作系统虚拟机
2. 在每个系统上**原生编译**你的代码
3. 自动打包发布

## 为什么GitHub能做到这个？

### 原因1：真实的物理机器

GitHub在全球有数据中心，里面有大量服务器：

```
GitHub的服务器配置：
├─ Windows Server 2022虚拟机
├─ macOS虚拟机（运行在Mac硬件上）
└─ Ubuntu Linux虚拟机
```

### 原因2：虚拟化技术

每次构建时，GitHub会：
1. 从虚拟机池中分配一个干净的系统
2. 按照你的配置安装软件
3. 执行构建任务
4. 保存结果
5. 销毁虚拟机（下次重新分配）

### 原因3：并行执行

三个虚拟机**同时**工作，所以速度快：

```
时间轴：
0分钟  → 触发构建
1分钟  → 启动3个虚拟机（并行）
2-5分钟→ 安装依赖、编译（并行）
6分钟  → 3个平台都完成！
```

如果手动操作需要：
- 在Windows上打包：5分钟
- 切换到macOS打包：5分钟
- 切换到Linux打包：5分钟
- **总共：15分钟+切换系统时间**

## 使用GitHub Actions的步骤

### 第1步：创建GitHub仓库

1. 访问 https://github.com
2. 注册/登录账号
3. 创建新仓库（可以是私有的）

### 第2步：推送代码

```bash
cd /Users/bytedance/Repositories/monthlyreport

# 初始化git（如果还没有）
git init

# 添加所有文件
git add .

# 提交
git commit -m "初始提交"

# 添加远程仓库
git remote add origin https://github.com/你的用户名/monthlyreport.git

# 推送
git push -u origin master
```

### 第3步：触发构建

两种方式触发：

**方式1：打标签发布**
```bash
git tag v1.0
git push origin v1.0
```

**方式2：手动触发**
1. 打开GitHub仓库页面
2. 点击"Actions"标签
3. 选择"构建多平台可执行文件"
4. 点击"Run workflow"

### 第4步：等待构建完成

大约5-10分钟后，GitHub会：
1. 自动创建Release
2. 上传3个平台的可执行文件
3. 发送邮件通知你

### 第5步：下载发布文件

在GitHub仓库的"Releases"页面下载：
- 送货单对账单工具-Windows.exe
- 送货单对账单工具-macOS
- 送货单对账单工具-Linux

## 配置文件说明

已创建的配置文件：`.github/workflows/build.yml`

### 关键配置解释

```yaml
strategy:
  matrix:
    os: [windows-latest, macos-latest, ubuntu-latest]
```
这行代码告诉GitHub：
- 同时在3个系统上运行
- `matrix.os` 会自动循环使用这3个值

```yaml
runs-on: ${{ matrix.os }}
```
指定在哪个系统上运行：
- `windows-latest` = Windows Server 2022
- `macos-latest` = macOS 13 (Ventura)
- `ubuntu-latest` = Ubuntu 22.04

## 优势对比

### 传统方式
```
❌ 需要3台不同系统的电脑
❌ 手动在每个系统上编译
❌ 耗时：15-30分钟
❌ 容易出错
❌ 环境不一致
```

### GitHub Actions方式
```
✅ 只需一台电脑（任何系统）
✅ 自动编译所有平台
✅ 耗时：5-10分钟（并行）
✅ 完全自动化
✅ 环境一致
✅ 完全免费（公开仓库）
```

## 免费额度

GitHub Actions免费额度：

| 仓库类型 | 每月免费时长 |
|---------|------------|
| 公开仓库 | 无限制 ✅ |
| 私有仓库 | 2000分钟 |

**计算示例**：
- 每次构建：~10分钟（3个系统并行）
- 私有仓库：每月可构建200次
- 公开仓库：无限次

## 实际使用示例

### 场景1：发布新版本

```bash
# 1. 修改代码
vim merge_delivery_orders.py

# 2. 提交更改
git add .
git commit -m "修复bug"

# 3. 打标签
git tag v1.1

# 4. 推送
git push origin v1.1

# 5. 等待5-10分钟，GitHub自动：
#    - 构建Windows/macOS/Linux版本
#    - 创建Release
#    - 上传所有文件
```

### 场景2：测试构建

不想发布，只想测试编译：

```bash
# 推送到特定分支
git checkout -b test-build
git push origin test-build

# 然后在GitHub Actions页面手动触发
```

## 查看构建日志

如果构建失败，可以查看详细日志：

1. 打开GitHub仓库
2. 点击"Actions"标签
3. 点击失败的构建
4. 查看每个步骤的输出

## 高级技巧

### 自定义构建触发条件

只在main分支触发：
```yaml
on:
  push:
    branches:
      - main
```

每天自动构建：
```yaml
on:
  schedule:
    - cron: '0 0 * * *'  # 每天UTC 0点
```

### 添加Windows签名

如果有代码签名证书：
```yaml
- name: 签名Windows程序
  run: |
    signtool sign /f cert.pfx /p ${{ secrets.CERT_PASSWORD }} dist/*.exe
```

## 常见问题

**Q: 构建失败怎么办？**
A: 查看Actions日志，通常是依赖问题。确保pyproject.toml正确。

**Q: 可以构建其他系统吗？**
A: 可以，GitHub支持：
- Windows (多个版本)
- macOS (多个版本)
- Ubuntu (多个版本)

**Q: 私有仓库会泄露代码吗？**
A: 不会，构建过程完全私密，生成的文件可以选择公开或私有。

**Q: 构建时间能更快吗？**
A: 可以启用缓存：
```yaml
- uses: actions/cache@v3
  with:
    path: ~/.cache/uv
    key: ${{ runner.os }}-uv-${{ hashFiles('**/uv.lock') }}
```

## 总结

GitHub Actions就像是**在云端租用了3台电脑**：
- 你只需写一次配置
- GitHub自动在3个系统上编译
- 完全免费（公开项目）
- 每次更新自动构建

比起传统方式（需要3台电脑手动操作），这简直是**开发者的福音**！

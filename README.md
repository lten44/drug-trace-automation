# 药品批发企业追朔码自动处理软件

[![Python](https://img.shields.io/badge/Python-3.12+-blue.svg)](https://www.python.org/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Version](https://img.shields.io/badge/Version-v3.0-orange.svg)]()

> 一个用于自动化处理药品追溯码的桌面应用程序，通过模拟鼠标操作与"码上放心"客户端交互，实现追溯码的自动查询、排重和1级码提取，生成标准化输出表格。

**当前版本**：v3.0 | **运行平台**：Windows 7/8/10/11 | **打包方式**：PyInstaller

---

## ✨ v3.0 核心亮点

| 特性 | 说明 |
|:-----|:------|
| **统一排重模式** | 移除全1级/混装两模式，所有追溯码走同一流程，自动排重 |
| **输入表匹配** | 1级码查到同箱码后，只保留输入表中也存在的，不同批号互不干扰 |
| **零重复输出** | covered_set 覆盖排重 + output_1level_set 防重复 |
| **5秒倒计时** | 操作前倒计时提示，充分准备 |
| **失败弹窗** | 3次重试失败后弹出 跳过/重试/停止 三选一 |
| **断点续跑** | 崩溃后自动从上次进度继续 |
| **UI美化** | 卡片式玻璃设计 + 左侧彩色装饰条 + 全平台兼容 |

---

## 🚀 快速开始

### 系统要求

- Windows 7/8/10/11
- "码上放心"客户端（已安装并登录）

### 使用预编译版本（推荐）

1. 下载 `Setup-药品批发企业追朔码自动处理软件-v3.0.exe`
2. 双击安装，一路下一步
3. 从桌面快捷方式启动

### 从源码运行

```bash
git clone https://github.com/lten44/drug-trace-automation.git
cd drug-trace-automation
pip install -r requirements.txt
python gui.py
```

### 首次使用

1. **校准位置** → 点击"校准位置"，按 F5 依次记录5个按钮坐标（窗口失焦也能按）
2. **选择文件** → 点击"浏览"选择要处理的 Excel 文件
3. **开始处理** → 自动排重，无需选择模式

---

## 🏗️ 核心算法

### 统一排重模式（v3.0）

```
输入表追溯码集合 = 输入文件所有码

排序：非1级码(3级→2级) → 1级码

FOR 每条记录:
    IF 码 ∈ covered_set → 跳过（已被父级覆盖）

    IF 包装数 ≠ "1" (2级/3级码):
        query_single → 一键复制全部码
        输出全部下级1级码
        covered_set += 整箱全部码（子码全跳过）

    IF 包装数 == "1" (1级码):
        query_single → 查出同箱全部码
        仅保留 ∩ 输入表追溯码集合 的项
        covered_set += 仅实际输出的码（不覆盖整箱，不同批号不误跳）
```

详见 [运行逻辑文档](药品批发企业追朔码自动处理软件_运行逻辑文档_v2.4.md)

---

## 📁 项目结构

```
drug-trace-automation/
├── gui.py                        # 主程序 v3.0（Tkinter）
├── build-v3.spec                 # PyInstaller 构建配置
├── build.bat                     # 构建脚本
├── publish.ps1                   # 一键发布脚本
├── installer.iss                 # Inno Setup 安装包配置
├── requirements.txt              # Python 依赖
├── version.json                  # 版本信息（自动更新用）
├── icon.ico                      # 程序图标
├── README.md                     # 本文件
├── LICENSE                       # MIT 许可证
├── DEVELOPMENT_LOG.md            # 开发历程
├── RELEASE_NOTE_v3.0.md          # Release 说明模板
├── 药品批发企业追朔码自动处理软件_开发文档.md        # 完整开发文档
├── 药品批发企业追朔码自动处理软件_运行逻辑文档_v2.4.md  # 运行逻辑文档
└── v2.4更新日志.md               # 更新日志
```

---

## 🛠️ 技术栈

| 库 | 用途 | 版本 |
|:----|:------|:------|
| pyautogui | 鼠标/键盘模拟操作 | ≥0.9.54 |
| pyperclip | 剪贴板读写 | ≥1.8.2 |
| openpyxl | Excel 文件读写 | ≥3.1.2 |
| pynput | 全局键盘监听（F1/F5） | ≥1.7.6 |
| tkinter | GUI 界面（Python 内置） | — |
| PyInstaller | 打包为 exe | 6.20.0 |
| Inno Setup | 制作安装包 | 6.x |

---

## 🔄 版本历史

| 版本 | 更新内容 |
|:----:|:---------|
| v1.0 | 基础框架：鼠标模拟 + 3点校准 + Excel 读写 |
| v1.1~v1.9 | 逐步优化过滤规则、UI、自动更新、软件更名 |
| v2.0 | 双位置校准（复制按钮查询前后位置不同） |
| v2.1 | 智能跳过：码包装数为1的行跳过查询 |
| v2.2 | 自定义图标，更新优化 |
| **v3.0** | **统一排重模式 + 输入表匹配 + UI全面美化 + 5项优化** |

详细开发历程请查看 [DEVELOPMENT_LOG.md](DEVELOPMENT_LOG.md)

---

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

---

## 📝 许可证

本项目采用 [MIT](LICENSE) 许可证开源。

> **免责声明**：本软件仅供学习交流使用，请确保使用符合相关法律法规。使用本软件产生的任何后果由用户自行承担。
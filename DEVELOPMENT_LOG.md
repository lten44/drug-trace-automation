# 国瑞新特药追朔码自动处理软件 - 开发历程

> 一个用于自动化处理药品追溯码的桌面应用程序，通过模拟鼠标操作与"码上放心"客户端交互，实现追溯码的自动查询和1级码提取。

## 📋 项目背景

### 需求来源
用户日常工作需要将Excel表格中的追溯码转换为文本格式，然后逐个复制到"码上放心"客户端查询关联的所有追溯码，再整理成"标2"格式输出。由于每天处理量可能达到6000条，手工操作效率极低，因此需要开发自动化工具。

### 技术挑战
1. **API方案不可行**：用户企业为批发企业，无法获取"码上放心"开放平台的零售药店接口权限
2. **UI自动化困难**："码上放心"客户端基于WebView（Chromium）开发，传统UI自动化工具（如pywinauto）无法定位控件
3. **权限限制**：WES接口需要年费2800元，成本过高

## 🚀 开发迭代历程

### v1.0 - 基础框架搭建
**核心功能**：
- 使用 `pyautogui` 进行鼠标/键盘模拟操作
- 屏幕坐标校准系统（3个校准点：输入框、查询按钮、复制按钮）
- Excel文件读取与写入
- 1级码过滤逻辑（从复制结果中提取1级码）

**技术实现**：
```python
# 核心流程
1. 读取Excel获取追溯码列表
2. 对每个追溯码：
   - 点击输入框并粘贴追溯码
   - 点击查询按钮
   - 等待查询结果
   - 点击"一键复制所有码"
   - 读取剪贴板并过滤1级码
3. 生成输出Excel文件
```

---

### v1.1 - 过滤规则优化
**问题**：2级码、3级码也被包含在输出中

**解决方案**：
- 修改 `filter_level_one()` 函数
- 改为只提取"1级码"标记后的所有纯数字行
- 遇到其他级别标记（2级码、3级码等）时停止收集

---

### v1.2 - 权限问题修复
**问题**：默认输出到程序目录导致权限被拒绝

**解决方案**：
- 默认输出路径改为桌面下的"追溯码输出"文件夹
- 添加日志文件输出（`debug.log`），解决 `--windowed` 模式下看不到控制台输出的问题

---

### v1.3 - 表格样式优化
**问题**：用户需要保留原始表格样式

**解决方案**：
- 改为直接复制原始Excel文件
- 只修改追溯码列和码包装数列
- 保留所有原始样式（字体、颜色、边框等）

**回退**：后续版本改回新建表格模式，因为复制模式难以处理行数变化

---

### v1.4 - 产品批号自动获取
**新增功能**：
- 第4个校准点：产品批号位置
- 查询后自动双击选中并复制批号
- 删除"零头数"、"生产信息"、"验证信息"三列
- 文件名包含所有不重复药品名（用&连接）

**技术细节**：
```python
# 获取产品批号
batch_pos = self.cal.get("batch_no_pos")
pyautogui.doubleClick(batch_pos['x'], batch_pos['y'])
pyautogui.hotkey('ctrl', 'c')
batch_no = pyperclip.paste().strip()
```

---

### v1.5 - UI美化与默认路径
**改进**：
- 卡片式UI布局
- 渐变色标题栏
- 图标标识（💊🔧📂⚙️💡）
- 默认路径设置功能（可修改默认输入/输出目录）
- 文件名前缀统一加"汕头国瑞新特药"

---

### v1.6 - 扫描时间修复
**问题**：扫描时间列读取不到

**原因**：按固定列索引读取，输入表格列顺序可能不同

**解决方案**：
- 改为按表头名称查找列索引
- 兼容任意列顺序的输入表格

---

### v1.7 - 样式调整
**改进**：
- 所有列左对齐
- 等线字体
- 列宽自适应（根据内容自动计算）

---

### v1.8 - 自动更新功能
**新增功能**：
- 启动时自动检查 `E:\国瑞新特药追朔码自动处理软件\` 目录
- 发现新版本自动复制覆盖并重启
- 版本号比较逻辑（支持 v1.9 < v1.10 < v2.0 等）

**更新机制**：
```python
def check_and_update():
    # 检查 version.json 中的版本号
    # 如果新版本 > 当前版本：
    #   1. 复制新exe覆盖当前exe
    #   2. 启动新版本
    #   3. 退出当前程序
```

---

### v2.0 - 双位置校准
**问题**："一键复制所有码"按钮在首次查询前后位置不同

**解决方案**：
- 增加第5个校准点：`copy_btn_after`（查询后的按钮位置）
- 校准步骤变为5步：
  1. 追溯码输入框
  2. 查询按钮
  3. 一键复制所有码（首次查询前位置）
  4. 一键复制所有码（查询结果后的位置）
  5. 产品批号文字位置

---

## 🏗️ 技术架构

### 核心模块

```
gui.py
├── 配置与常量
│   ├── VERSION: 版本号
│   ├── UPDATE_DIR: 自动更新目录
│   └── OUTPUT_FOLDER: 默认输出目录
│
├── 数据模型
│   ├── MetaInfo: 单据元信息（扫码日期、单据类型等）
│   ├── TraceRecord: 追溯码记录
│   └── ExcelData: Excel数据容器
│
├── 核心功能
│   ├── read_sales_excel(): 读取输入Excel
│   ├── filter_level_one(): 过滤1级码
│   ├── write_result_excel(): 生成输出Excel
│   └── generate_filename(): 生成文件名
│
├── UI自动化
│   ├── ScreenCalibrator: 屏幕坐标校准器
│   ├── CodeQueryUI: 码上放心查询UI
│   └── CalibrateWindow: 校准窗口
│
└── 主应用
    └── DrugTraceApp: 主窗口类
```

### 依赖库

```
pyautogui      # 鼠标/键盘模拟
pyperclip      # 剪贴板操作
openpyxl       # Excel处理
tkinter        # GUI界面
```

### 文件结构

```
drug_trace_app/
├── gui.py              # 主程序（单文件）
├── calibration.json    # 校准坐标数据
├── config.json         # 用户配置（默认路径等）
├── debug.log           # 运行日志
└── dist/
    └── 国瑞新特药追朔码自动处理软件.exe

更新目录（E:\国瑞新特药追朔码自动处理软件\）:
├── 国瑞新特药追朔码自动处理软件.exe
└── version.json        # 版本信息
```

---

## 🔧 关键技术点

### 1. 屏幕坐标校准系统

由于"码上放心"是WebView应用，无法通过控件ID定位，采用坐标校准方案：

```python
class CalibrateWindow:
    def __init__(self, ...):
        self.steps = [
            ("input_box", "追溯码输入框"),
            ("query_btn", "查询按钮"),
            ("copy_btn", "一键复制所有码（首次）"),
            ("copy_btn_after", "一键复制所有码（查询后）"),
            ("batch_no_pos", "产品批号位置"),
        ]
    
    def _track_mouse(self):
        # 后台线程实时跟踪鼠标位置
        while self.running:
            x, y = pyautogui.position()
            # 更新显示...
    
    def _capture_position(self):
        # 按F5记录当前鼠标位置
        self.cal.set(name, self.current_x, self.current_y)
```

### 2. 1级码过滤算法

```python
def filter_level_one(raw_text):
    codes = []
    lines = raw_text.strip().split("\n")
    found_level_one = False
    
    for line in lines:
        line = line.strip()
        if not line: continue
        
        # 检测级别标记
        if re.match(r"^\d级码", line):
            if re.match(r"^1级码", line):
                found_level_one = True  # 开始收集
            else:
                found_level_one = False  # 遇到其他级别，停止收集
            continue
        
        # 收集1级码后的所有追溯码
        if found_level_one and re.match(r"^\d{16,24}$", line):
            codes.append(line)
    
    return codes
```

### 3. 自动更新机制

```python
def check_and_update():
    update_dir = r'E:\国瑞新特药追朔码自动处理软件'
    version_file = os.path.join(update_dir, "version.json")
    new_exe = os.path.join(update_dir, "国瑞新特药追朔码自动处理软件.exe")
    
    # 读取最新版本号
    with open(version_file, 'r') as f:
        latest_version = json.load(f)['version']
    
    # 版本号比较
    if compare_versions(latest_version, VERSION) > 0:
        # 复制新exe覆盖当前exe
        shutil.copy2(new_exe, sys.executable)
        # 启动新版本
        os.startfile(sys.executable)
        return True  # 需要退出
    
    return False
```

---

## 📊 开发数据统计

| 指标 | 数据 |
|------|------|
| 开发周期 | 约2周 |
| 版本迭代 | 20+ 次 |
| 代码行数 | ~1000 行 |
| 校准点数量 | 5 个 |
| 核心功能模块 | 8 个 |

---

## 🎯 核心问题解决

### 问题1：API权限不足
**方案**：放弃API方案，改用UI自动化

### 问题2：WebView无法定位控件
**方案**：屏幕坐标校准系统 + 实时鼠标跟踪

### 问题3：按钮位置变化
**方案**：双位置校准（首次/查询后）

### 问题4：大量数据处理
**方案**：
- 后台线程处理，不阻塞UI
- 进度条实时显示
- 可随时停止

### 问题5：软件分发更新
**方案**：
- 单文件exe打包
- 自动更新机制
- 版本号管理

---

## 💡 经验总结

### 技术选型经验
1. **pyautogui vs pywinauto**：对于WebView应用，pyautogui的坐标模拟比pywinauto的控件查找更可靠
2. **单文件 vs 模块化**：使用PyInstaller打包为单文件，便于分发
3. **配置持久化**：使用JSON文件保存用户配置和校准数据

### 用户体验优化
1. **校准系统**：F5热键 + 实时鼠标坐标显示，操作直观
2. **进度反馈**：进度条 + 状态文字，让用户了解处理进度
3. **错误处理**：详细的日志记录，便于排查问题
4. **自动更新**：无需手动下载，打开即是最新版本

### 迭代开发方法
1. **快速原型**：先实现核心功能，再逐步优化
2. **用户反馈驱动**：每个版本都针对用户反馈的具体问题
3. **向后兼容**：配置和校准数据格式保持稳定

---

## 🔮 未来展望

### 可能的改进方向
1. **OCR识别**：如果"码上放心"界面变化，可考虑使用OCR识别按钮位置
2. **多线程查询**：同时查询多个追溯码，提高效率
3. **批量导入**：支持拖拽文件夹自动处理多个Excel
4. **数据校验**：增加追溯码格式校验，提前发现错误

---

## 📄 开源协议

本项目采用 MIT 协议开源。

---

## 👨‍💻 开发者

- 开发周期：2025年1月
- 主要技术：Python + Tkinter + PyAutoGUI
- 目标平台：Windows 10/11

---

> **提示**：使用本软件需要配合"码上放心"客户端，请确保已安装并登录。

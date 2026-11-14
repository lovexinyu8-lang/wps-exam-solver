# wps-exam-solver

专用于全国计算机等级考试（NCRE）二级 WPS Office、WPS 认证考试及办公自动化操作题求解的 Agent Skill。通过调度 Windows 本机原生 WPS COM 接口（`Kwpp.Application` / `Kwps.Application` / `Ket.Application`）单轮脚本完成作答，并强制执行只读模式二次布尔断言自检，保证生成结果与阅卷系统的机器判分标准一致。

## 为什么采用原生 COM 而不是拆包改 XML

直接用 `zipfile` 拆包修改 OOXML 在考试阅卷场景极易踩坑（例如主题色槽位、文字特效退化、SmartArt 节点结构、内嵌图表数据源行数不一致等）。本 Skill 规定：

- 统一由 Windows 原生 WPS COM 完成核心操作与最终 `Save()` 落盘；
- 仅对极少数无 COM 写入接口的专属节点（如 Word `w14:textFill` 多色标渐变、WPS 原生双斜线表头、主题底纹 `w:themeFill`）允许在局部通过 `python-docx` / XML 补充，且**最后一次打开与保存仍必须由 `Kwps.Application` 执行**。

## 涵盖能力与避坑范围

- **WPS 演示（`Kwpp.Application`）**
  - 16:9 宽屏（`960×540pt`）与“确保适合”时的母版/版式占位符同步缩放与内容平移
  - 首页强制跟随母版背景（`FollowMasterBackground = -1`）
  - `TextFrame2` 垂直与水平锚定（避开缺少 `anchorCtr="0"` 扣分）
  - 小型大写字母、图案填充、阴影与三维相机参数
  - SmartArt 节点增删需 `WithWindow=True`、图表数据区动态扩容、无头状态下禁用 `ShapeRange.Align` / `Distribute` 改用几何坐标计算
- **WPS 文字（`Kwps.Application`）**
  - 页面边距、装订线与主题背景色（先调 `Solid()` 再着色）
  - 页眉日期域、底边框与水印添加顺序及 `Anchor=header.Range` 绑定
  - 固定/多倍行距常量、文本边框（排除末尾 `\r` 设在 `Font` 上）、文字阴影/发光防退化
  - 显式 `while` 循环查找替换、图片禁纵横比调尺寸与椭圆裁剪
  - 表格先设宽高再 `Merge()`、纵向合并后改用 `Cell(r, c)` 访问、内嵌图表插行扩容与垂直文字方向旋转页面顺序
- **WPS 表格（`Ket.Application`）**
  - 英文函数名与英文逗号规范、动态定位数据末行防误算汇总行
  - 正则拆分混合列替代无头 `TextToColumns` 静默失效
  - 数据透视表创建、经典行标签布局与刷新、会计专用格式与边框
- **运行安全与准出铁律**
  - 打开文件前先做 `1ms` 写锁预检（`r+b`），防弹窗死锁
  - 严禁全局 `Stop-Process` 误杀前台考试或预览窗口
  - 严禁在 `ReadOnly=True` 下激活内嵌图表 `ChartData.Activate()`
  - 提供标准 `solve.py` 与只读 `check.py` 双阶段骨架，全部断言 `PASS` 方可交付

## 环境要求

- **操作系统**：Windows 10 / 11
- **办公软件**：WPS Office（建议 12.0+，已注册 `Kwpp.Application` / `Kwps.Application` / `Ket.Application` COM 组件）
- **Python**：3.9+

安装依赖：

```powershell
pip install -r requirements.txt
```

## 安装为 Agent Skill

适用于支持 `.agents/skills/` 规范的 Agent 平台（如 Antigravity / Codex / Claude Code 等）：

### 全局安装

```powershell
git clone https://github.com/lovexinyu8-lang/wps-exam-solver.git "$env:USERPROFILE\.agents\skills\wps-exam-solver"
```

### 仅当前项目安装

在项目根目录下执行：

```powershell
git clone https://github.com/lovexinyu8-lang/wps-exam-solver.git ".agents\skills\wps-exam-solver"
```

## 仓库结构

```text
wps-exam-solver/
├── SKILL.md           # Skill 核心指令、COM 避坑字典与标准作答/自检骨架
├── README.md          # 项目说明与安装指南
├── requirements.txt   # Python 依赖（pywin32 / python-docx / openpyxl）
└── .gitignore         # 忽略缓存、虚拟环境与 WPS 临时锁文件（~$*）
```

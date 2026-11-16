# wps-exam-solver

面向 **NCRE 二级 WPS Office / KOS 认证 / 办公自动化课程实操** 的极速求解与判分级自检引擎（Agent Skill + 独立 Python CLI）。

核心目标：**对齐真实考场机器阅卷的判分标准**——采用 **Windows 原生 WPS 动态 COM（`Kwpp` / `Kwps` / `Ket`）+ OOXML 底层外科补丁与双轨审计**，既可挂载到 AI Agent 中实现单轮闭环极速作答，也可脱离 Agent 独立作为 **“作答文件 vs 参考答案 / JSON 分值规则”** 的自动化判分系统，平替商业培训学校模考软件。

---

## 为什么采用“动态 COM + OOXML 双轨”

只用 `zipfile` 裸拆 XML 改包在考试阅卷中极易丢分（如主题色槽位、文字特效退化、SmartArt 节点关系、内嵌图表数据区行数不匹配等）；反过来，单纯调用 WPS COM 也会遇到少数内核写盘盲区（如 `Kwpp` 不持久化 `<p:prnPr>` 打印设置、合并单元格剥离隐藏 `hMerge` 居中、Word `w14:textFill` 多色标渐变与 `w14:glow` 饱和度、`w:themeFill` 槽位、`wpsCustomData:diagonals` 双斜线表头等）。

本项目采取**三段式判分级闭环**：
1. **原生动态 COM 执行**：统一通过 `win32com.client.dynamic.Dispatch` 调度 WPS（绕过双装 Office 的 `gen_py` 缓存冲突），执行前做 `1ms` 写锁预检与 `.bak` 自动备份回滚。
2. **OOXML 外科补丁 + WPS 终保存**：仅在 COM 写不出的节点做局部 XML 注入，之后强制由 WPS 重新打开并 `Save()` 一次。
3. **COM + OOXML 双轨审计**：只读打开核验渲染/对象树属性，同时直读底层 XML（`anchorCtr="0"`、`behindDoc="1"`、`chineseCounting`、`<c:ptCount>`、`<sortState>`、条件格式 `<cfRule><formula>` 相对行号等）逐项断言。

---

## 一、 环境要求与安装

- **操作系统**：Windows 10 / 11
- **办公软件**：WPS Office 完整版（建议 12.0+；已注册 `Kwpp.Application` / `Kwps.Application` / `Ket.Application` 或旧版 `WPP` / `WPS` / `ET` COM 接口；若仅使用 `--no-com` 纯 XML 审计可免装 WPS）
- **Python**：3.9+

### 安装依赖

```powershell
pip install -r requirements.txt
```

### 方式 1：作为 Agent Skill 安装（Antigravity / Codex / Claude Code 等）

全局安装（所有工作区自动可用）：
```powershell
git clone https://github.com/lovexinyu8-lang/wps-exam-solver.git "$env:USERPROFILE\.agents\skills\wps-exam-solver"
```

仅当前项目安装（在项目根目录下执行）：
```powershell
git clone https://github.com/lovexinyu8-lang/wps-exam-solver.git ".agents\skills\wps-exam-solver"
```

### 方式 2：直接作为独立开源判分工具拉取使用

```powershell
git clone https://github.com/lovexinyu8-lang/wps-exam-solver.git
cd wps-exam-solver
pip install -r requirements.txt
```

---

## 二、 完整使用流程

### 流程 A：搭配 AI Agent 自动做题并验卷

1. **安装到 `.agents/skills/wps-exam-solver`** 后，先让环境体检确认本机 COM 注册正常：
   ```powershell
   python scripts/wps_env.py --probe
   ```
   若电脑同时装有 Microsoft Office，工具会自动提示注意 `gen_py` 碰撞（本库内置的 `wps_session` 已默认强制晚绑定与实例隔离，不会误杀前台已打开的 WPS/考试页面）。
2. **向 Agent 发出考题与作答文件目录**（例如附带题目截图或文本，指明 `.docx` / `.wps` / `.xlsx` / `.et` / `.pptx` / `.dps` 路径）：
   - Agent 会按 `SKILL.md` 先读取对应专项字典（`references/word-kwps.md` / `excel-ket.md` / `ppt-kwpp.md`）。
   - 最多做 **1 次**全量只读结构探查，随后生成单轮闭环 `solve.py` 调用 `wps_session(..., backup=True)` 完成作答与必要 XML 补丁。
3. **强制双轨自检准出**：
   - 有参考答案时，自动运行 `ref_diff.py` 逐项核对至全 `PASS`；
   - 无参考答案时，自动运行基于 `AuditSuite` + `audit_ooxml()` 的 `check.py`，全部硬性指标断言通过后才交付。

---

### 流程 B：独立用作开源阅卷/练题判分系统（培训学校软件平替）

#### 场景 1：检查本机 WPS COM 环境与目标文件锁状态
```powershell
python scripts/wps_env.py --probe path\to\作答文件.docx
```
会输出已注册的所有 WPS/Office ProgID、可执行文件路径、版本号、是否被锁定占用（`~$` 锁文件）以及文件真实容器类型（`ooxml` 或 `ole2`）。

#### 场景 2：纯只读查看底层 OOXML 隐含得分点（无需打开 WPS）
```powershell
python scripts/ooxml_patch.py path\to\作答文件.pptx
```
直接以 JSON 打印页边距 `pgSz/pgMar`、页眉图片 `behindDoc`、多级列表 `numbering.xml`、表格 `w:themeFill`、图表 `<c:ptCount>`、Excel 条件格式公式与 `<sortState>`、PPT `anchorCtr` 与 `<p:prnPr>`。

#### 场景 3：有参考答案 —— 一对一判分级对拍打分
```powershell
# COM + OOXML 双轨全维比对（最接近真实考场阅卷）
python scripts/ref_diff.py path\to\作答文件.pptx path\to\参考答案.pptx

# 输出结构化 JSON 供脚本或前端消费
python scripts/ref_diff.py path\to\作答文件.xlsx path\to\参考答案.xlsx --json

# 纯 OOXML 极速比对（免 WPS 进程，适合 CI/Linux 服务端批量预检）
python scripts/ref_diff.py path\to\作答文件.docx path\to\参考答案.docx --no-com
```
- 退出码：`0` 表示全部通过，`1` 表示存在差异。
- 输出会精确指出失分位置（例如单元格 `E2` 公式不一致、条件格式行号偏移 `$B3` vs `$B2`、缺少 `anchorCtr="0"`、图表数据点 `<c:ptCount>` 不符等）。

#### 场景 4：无参考答案（或需按分值权重算分）—— JSON 规则判分
1. **根据已有参考答案一键生成规则模板**（也可参考 [examples/ppt_sample.rubric.json](examples/ppt_sample.rubric.json) 手写）：
   ```powershell
   python scripts/rubric_grader.py init path\to\参考答案.pptx -o my_exam.rubric.json --total-points 20
   ```
2. **用规则文件对学生/Agent 作答文件按分值打分**：
   ```powershell
   python scripts/rubric_grader.py grade path\to\作答文件.pptx my_exam.rubric.json
   # 加 --json 输出每小题得分明细
   python scripts/rubric_grader.py grade path\to\作答文件.pptx examples/ppt_sample.rubric.json --json
   ```

#### 场景 5：整套模拟试卷多题目录批量自动批改
将所有考生的作答文件放在 `作答目录`，参考答案（或 `*.rubric.json` 规则文件）放在 `参考答案或规则目录`，一条命令输出整卷汇总得分表：
```powershell
python scripts/rubric_grader.py batch path\to\作答目录 path\to\参考答案或规则目录
```

---

## 三、 涵盖考点与避坑范围

- **WPS 文字（`Kwps`）**：页面边距/装订线/横向与垂直文字方向顺序、页面背景 `Solid()` 写盘、下一页分节符与仅首页 `behindDoc="1"` 图片水印、页眉日期域与双线边框顺序、内建样式枚举 ID（`-1`~`-4`）+ 中文多级列表（`chineseCounting` 自校准）、独立字母列表、`w14:textFill` 渐变/`w14:glow` 饱和度/阴影防退化、显式 `while` 循环替换、倒序 `PhoneticGuide` 拼音指南、表格合并顺序与 `wpsCustomData:diagonals` 双斜线表头、主题底纹 `w:themeFill`、内嵌图表锚点与 4 行扩容、`SaveAs2(path, 0)` 真 `.wps` 保存。
- **WPS 表格（`Ket`）**：`UsedRange` 批量元组读写、英文函数与末行汇总防误算、正则拆分混合列替代无头 `TextToColumns`、身份证 `DATE(MID/VLOOKUP)` 公式、**条件格式先 `Activate()+Select()` 防公式行号偏移**、`ws.Sort` 生成 `<sortState>`（含 `xlPinYin` 拼音排序）、数据透视表 `RowAxisLayout(1)` + 联动图表、会计专用格式。
- **WPS 演示（`Kwpp`）**：16:9 宽屏（`960×540pt`）与“确保适合”母版/版式/内容联动缩放平移、全页跟随母版背景、几何坐标替换无头失效的 `Align/Distribute`、`TextFrame2` 双锚定（写出 `anchorCtr="0"`）、艺术字图案填充与阴影、SmartArt `WithWindow=True` 护航、表格合并后恢复隐藏 `hMerge` 居中、图表 4 行扩容、矩形先建后转动作按钮、`<p:prnPr>` 打印设置注入。

---

## 四、 仓库结构

```text
wps-exam-solver/
├── SKILL.md                   # Agent Skill 主指令：执行铁律与标准求解/自检骨架
├── README.md                  # 项目说明、安装方式与完整 CLI/Agent 使用流程
├── requirements.txt           # Python 依赖（pywin32 / python-docx / openpyxl）
├── .gitignore                 # 忽略缓存、虚拟环境、临时锁文件（~$*）与备份文件
├── examples/
│   └── ppt_sample.rubric.json # 加权分值评分规则示例
├── scripts/
│   ├── wps_env.py             # 多版本 ProgID 探测、dynamic.Dispatch、写锁预检、自动备份与进程隔离
│   ├── ooxml_patch.py         # WPS COM 盲区 XML 外科补丁 + 纯只读 OOXML 审计器
│   ├── ref_diff.py            # 判分级“作答文件 vs 参考答案”全维比对 CLI（COM + OOXML）
│   └── rubric_grader.py       # JSON 分值规则判分引擎（grade / init / batch 批量阅卷）
└── references/
    ├── word-kwps.md           # WPS 文字 (Kwps) 考点与坑点字典
    ├── excel-ket.md           # WPS 表格 (Ket) 考点与坑点字典
    ├── ppt-kwpp.md            # WPS 演示 (Kwpp) 考点与坑点字典
    └── grading-standard.md    # 机器阅卷常见“肉眼正常但判 0 分”对照表与规则类型速查
```

---

## 五、 已知边界

1. **COM 层仅支持 Windows**：`pywin32` 与 WPS COM 调度要求 Windows 10/11 环境；部分绿色版/极简版若未注册 `Kwps`/`Ket`/`Kwpp`（或旧版 `WPS`/`ET`/`WPP`），将无法执行 COM 自动化，仅可使用 `--no-com` 进行 OOXML 审计。
2. **旧版 OLE2 二进制文件**：非 ZIP 容器的传统二进制 `.doc` / `.xls` / `.wps` / `.et` 文件会自动跳过 OOXML 检查，仅走 COM 层面读写与对比。
3. **当前适配范围为 WPS Office**：Microsoft Office（`Word.Application` / `Excel.Application` / `PowerPoint.Application`）在内建样式本地化名称、内核特性及二级 MS Office 考纲（如邮件合并、尾注/交叉引用、PPT 触发器时间轴等）上存在显著差异，暂未纳入默认适配。

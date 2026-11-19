# DESIGN.md — wps-exam-solver GUI 设计与架构蓝图

> 本文件锁定 `wps-exam-solver` 面向大众的桌面工作台（Phase 2）的视觉审美、信息架构与技术栈边界。Phase 1 仅完成底层引擎优化与离线测试；Phase 2 按本蓝图实施，不重复发散选型。

---

## 一、 审美原则与 Design Tokens

本界面遵循 `claude-aesthetic` 暖纸编辑风，把机械冰冷的「机房判分软件」转化为克制、有书卷质感、低视觉疲劳的桌面阅卷工作台。

### 1. 核心色板
| 用途 | Token | Hex | 说明 |
| :--- | :--- | :--- | :--- |
| 主画布 | `--bg-canvas` | `#FAF9F5` | 全局底色，禁用纯白 `#FFFFFF` |
| 侧栏 / 分区柔底 | `--bg-soft` | `#F5F0E8` | 左侧导航栏、次级面板、表格偶数行 |
| 卡片底面 | `--bg-card` | `#EFE9DE` | 总分概览卡、规则编辑器卡片 |
| 强暖态 | `--bg-cream-strong` | `#E8E0D2` | 激活导航项、高亮选区 |
| 发丝边框 | `--border-hairline` | `#E6DFD8` | `1px` 分隔线（取代阴影） |
| 主墨色 | `--ink-primary` | `#141413` | 标题、得分数字、核心字段 |
| 正文墨色 | `--ink-body` | `#3D3D3A` | 诊断描述、修复建议文案 |
| 次要 / 弱提示 | `--ink-muted` / `--ink-soft` | `#6C6A64` / `#8E8B82` | 路径、耗时、次级元信息 |
| 主操作强调 | `--accent-primary` | `#CC785C` | 主 CTA（Hover `#B8674B`，Active `#A9583E`） |
| 语义通过 (PASS) | `--semantic-pass` | `#788C5D` | 全项通过 / 得分达标 |
| 语义提示 (COM/运行) | `--semantic-info` | `#6A9BCC` | COM 双轨激活、运行状态 |
| 语义警告 | `--semantic-warn` | `#D4A017` | 降级运行、缺失缓存值提示 |
| 语义失分 (FAIL) | `--semantic-fail` | `#C64545` | 扣分项、锁文件冲突 |
| 深炭证据窗 | `--bg-dark` / `--on-dark` | `#181715` / `#FAF9F5` | 原始 OOXML/JSON/公式对比折叠代码块 |

### 2. 字体与公开分发版权红线
> **禁止随公开发行包分发 Anthropic 专有字体（Copernicus / StyreneB / AnthropicSerif）**。

- **标题与大号得分读数**（`font-weight: 400`，负字距 `-0.025em`，靠字号不靠粗体）：
  `"AnthropicSerifCJK", "Source Serif 4", "Noto Serif SC", "Songti SC", "STSong", Georgia, serif`
- **界面操作与诊断正文**（`font-weight: 400/500`，行高 `1.55`，正文宽度 `< 75ch`）：
  `"Inter", "PingFang SC", "Microsoft YaHei UI", -apple-system, BlinkMacSystemFont, sans-serif`
- **XML / 公式 / 路径 / JSON**：
  `"JetBrains Mono", "Cascadia Code", ui-monospace, Consolas, monospace`

### 3. 克制法则
- **色块分级先于投影**：页面层次由 `#FAF9F5` → `#F5F0E8` → `#EFE9DE` + `1px solid #E6DFD8` 建立，不使用重阴影或 SaaS 卡片墙。
- **失分优先 + 证据分层**：普通考生/老师先看「错在哪 + 怎么改」的人话描述；点击展开后才在 `#181715` 深炭代码窗内查看原始 XML 节点（如 `<a:bodyPr anchor="t" anchorCtr="0"/>`）。

---

## 二、 锁定的桌面技术栈（Phase 2）

| 层级 | 选型 | 说明 |
| :--- | :--- | :--- |
| **桌面壳** | `pywebview` (Windows Edge WebView2) | 复用 Win10/11 系统自带 WebView2（免打包 Chromium）；原生文件/文件夹选择框直接返回真实绝对路径，原地锁检、`.bak` 备份与 COM 读写无缝闭环 |
| **本地后端** | `FastAPI` + `uvicorn` (本机 `127.0.0.1` 随机空闲端口) | 提供 REST + SSE 实时进度流；所有 WPS COM 任务统一经 `wps_env.COMWorkerQueue` 单线程 STA 队列串行执行，杜绝并发抢占与 `-2147221008` |
| **前端** | `React` + `TypeScript` + `Vite` | 构建为纯静态 `dist/` 由 FastAPI 托管；开发与 Playwright 回归测试支持 `--web` 无壳浏览器模式 |
| **分发打包** | `PyInstaller` 绿色目录包 | 免装 Python 环境，解压双击运行 |

---

## 三、 信息架构：4 个克制工作区视图

采用**左侧固定窄导航栏（`#F5F0E8`，220px）+ 右侧暖纸主画布（`#FAF9F5`）**的桌面布局：

1. **对拍与判分**
   - 拖入/选择「作答文件」+「参考答案（对拍）或 `.rubric.json` 规则（分值判分）」。
   - 模式切换：**双轨判分级（WPS COM + OOXML）** / **极速静态模式（`--no-com`）**。
   - 结果展示：大号 Regular 衬线得分数字（如 `18.5 / 20`）+ 耗时与引擎徽章；下方按「失分项优先」列表展示定位、期望值、实测值与人话修复提示，可按需展开深炭色 XML/公式对比证据。

2. **批量阅卷**
   - 选择「考生作答目录」+「单份参考答案 / 规则文件 / 套卷目录」（自动识别 N 生对 1 答案或镜像目录）。
   - 后台通过 `wps_batch_session` 复用隐藏 WPS 实例并经 SSE 实时推流进度条与当前批改文件名。
   - 展示班级均分、通过率、异常/未匹配文件清单，支持一键导出 **CSV (`utf-8-sig`) / 多工作表 Excel (`.xlsx`) 成绩单**。

3. **规则工坊**
   - 选择一份参考答案，一键调用 `scaffold_rubric` 自动抽取结构/样式/公式/XML 规则骨架（支持 COM 与 `--no-com`）。
   - 表单化增删检查项、调节每题分值权重、实时校验总分、一键保存/导出 `.rubric.json`。

4. **体检与补丁**
   - 本机环境体检：探测 `Kwps`/`Ket`/`Kwpp` 注册表、是否双装 MS Office、是否存在 `~$` 锁文件。
   - 隐形扣分项透视：只读扫描 `behindDoc`、`anchorCtr`、`prnPr`、`sortState`、`<cfRule>`、图表 `<c:ptCount>`。
   - 安全外科补丁：自动先创建 `.bak` 备份，再执行 `ooxml_patch` 注入并自动经 `wps_session` 原生重开 `Save()` 落盘。

---

## 四、 v1.0 产品边界（显式 Non-Goals）

- **v1.0 纯本地离线确定性判分工作台优先**：零外部 LLM API 费用、零网络上传隐私顾虑、零执行任意 LLM 生成代码的安全风险。
- 不在 v1.0 内置「上传截图调大模型自动代做生成 `solve.py` 执行」入口；若后续版本引入，须作为独立的 BYOK 可选模块，并强制经过人工代码预览确认门槛。

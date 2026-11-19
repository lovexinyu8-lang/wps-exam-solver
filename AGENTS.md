# AGENTS.md — wps-exam-solver

面向 Windows WPS Office (`Kwps`/`Ket`/`Kwpp`) 的考证求解 Skill 与开源双轨判分 CLI；远端仓库为 `lovexinyu8-lang/wps-exam-solver` (`main`)。

## 怎么跑与自检
- 离线回归单测（跨平台免 WPS）：`python -B -m unittest discover -s tests -v`
- 环境与文件锁体检：`python -B scripts/wps_env.py --probe [目标文件]`
- 只读 OOXML 检查：`python -B scripts/ooxml_patch.py <文件>`
- 对拍打分：`python -B scripts/ref_diff.py <作答文件> <参考答案> [--no-com] [--json]`
- 规则打分与导出：`python -B scripts/rubric_grader.py {grade|init|batch} ... [--no-com] [--export scores.xlsx]`

## 技术栈
- Python 3.9+、`pywin32`（强制 `win32com.client.dynamic.Dispatch` 晚绑定防双装 MS Office `gen_py` 碰撞；提供 `com_thread_scope`、`wps_batch_session` 长会话复用与 `COMWorkerQueue` 串行 STA 队列）。
- `python-docx`（Word `w14`/表格/文本框 OOXML 外科补丁）。
- `openpyxl>=3.1.2`（用于 `ooxml_patch.read_excel_static` 静态公式/缓存值读取与 `rubric_grader.export_batch_report` 多工作表 `.xlsx` 成绩单导出）。

## 目录与开发约定
- `SKILL.md` & `references/`：给“使用本 Skill 做题”的 Agent 看的主规则与考点字典，勿与开发者规范混写。
- `DESIGN.md`：Phase 2 面向大众桌面工作台的 Claude 暖纸审美与 `pywebview` + FastAPI + React/Vite 架构锁定文档。
- `scripts/`：`wps_env.py`（COM 隔离/线程队列/批量长会话/锁检/备份）、`ooxml_patch.py`（补丁与只读 XML/`openpyxl` 审计）、`ref_diff.py`（结构化 `AuditSuite` 与 1v1 比对）、`rubric_grader.py`（23 类加权规则与多模式批量阅卷/导出）。
- `tests/`：标准库 `unittest` 离线回归单测与脱敏样本构造器（`fixtures.py`、`test_ooxml_patch.py`、`test_rubric_and_diff.py`、`test_wps_env_unit.py`）。
- `examples/`：`ppt_sample.rubric.json`（JSON 规则示例）。
- 底层铁律：XML 外科补丁后必须经 `wps_session` 原生重开并 `Save()`；只读审计严禁调用 `ChartData.Activate()`；禁止空标题强杀系统里用户的 `wps`/`et`/`wpp` 进程；调试后不留 `__pycache__`、`*.bak.*`、`*.ooxml.tmp`。

## 当前状态与下一步计划（2026-11-19）
- 已完成（Phase 0 & Phase 1）：
  - 在本机沉淀全局 Claude 审美规范（`~/.agents/skills/claude-aesthetic/`）并在本仓写入 `DESIGN.md`；
  - 已切换为标准 Git 工作区（`scratch/wps-exam-solver`）与远端 `origin/main` 同步；
  - `wps_env.py` 新增 `com_thread_scope()`、`wps_batch_session()` 与 `COMWorkerQueue`；
  - `ooxml_patch.py` 修复 `audit_excel_xml` 按 `workbook.xml.rels` `r:id` 映射工作表、扩展 Word `w14` 补丁至表格与文本框、补齐静态 XML/`openpyxl` 提取；
  - `ref_diff.py` 输出结构化 `location/expected/actual/hint`、消除图表截断并支持长会话复用；
  - `rubric_grader.py` 修复 `word_numbering_xml` `lvlRestart=0` 判松问题、补齐 23 类规则 `--no-com` 静态实现、扩充 `init` 抽取与 `batch` 单答案/单规则/镜像匹配及 `.csv`/`.xlsx` 导出；
  - 补建 `tests/` 离线单测套组并实机验证 WPS COM 烟雾测试。
- 下一步（Phase 2 — 面向大众的桌面工作台）：
  1. 严格按 [DESIGN.md](DESIGN.md) 实现 `gui/`：FastAPI 本地服务（REST + SSE 实时进度流 + `COMWorkerQueue` 调度）+ React/TypeScript/Vite Claude 暖纸编辑风前端（4 大克制视图：对拍与判分 / 批量阅卷 / 规则工坊 / 体检与补丁）+ `pywebview` (Edge WebView2) 桌面外壳；
  2. 补充 `TestClient` API 契约测试与 Playwright `--web` 前端界面走查；
  3. 验证 `PyInstaller` 绿色包分发（严禁打包任何专有 Anthropic 字体）。

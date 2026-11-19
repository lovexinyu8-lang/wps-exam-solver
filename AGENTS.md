# AGENTS.md — wps-exam-solver

面向 Windows WPS Office (`Kwps`/`Ket`/`Kwpp`) 的考证求解 Skill 与开源双轨判分 CLI；远端仓库为 `lovexinyu8-lang/wps-exam-solver` (`main`)。

## 怎么跑与自检
- 环境与文件锁体检：`python -B scripts/wps_env.py --probe [目标文件]`
- 只读 OOXML 检查：`python -B scripts/ooxml_patch.py <文件>`
- 对拍打分：`python -B scripts/ref_diff.py <作答文件> <参考答案> [--no-com] [--json]`
- 规则打分：`python -B scripts/rubric_grader.py {grade|init|batch} ...`
- 现状提醒：当前目录**无 `.git`**（通过 GitHub MCP 推送同步），且尚无 `tests/` 回归单测。

## 技术栈
- Python 3.9+、`pywin32`（强制 `win32com.client.dynamic.Dispatch` 晚绑定防双装 MS Office `gen_py` 碰撞）、`python-docx`（仅用于 `ooxml_patch.py` Word 外科补丁）。
- `requirements.txt` 中的 `openpyxl>=3.1.2` 当前尚未在任何脚本导入，仅作预留。

## 目录与开发约定
- `SKILL.md` & `references/`：给“使用本 Skill 做题”的 Agent 看的主规则与考点字典，勿与开发者规范混写。
- `scripts/`：`wps_env.py`（COM 隔离/锁检/备份）、`ooxml_patch.py`（补丁与只读 XML 审计）、`ref_diff.py`（`AuditSuite` 与 1v1 比对）、`rubric_grader.py`（23 类加权规则）。
- `examples/`：`ppt_sample.rubric.json`（JSON 规则示例）。
- 底层铁律：XML 外科补丁后必须经 `wps_session` 原生重开并 `Save()`；只读审计严禁调用 `ChartData.Activate()`；禁止空标题强杀系统里用户的 `wps`/`et`/`wpp` 进程；调试后不留 `__pycache__`、`*.bak.*`、`*.ooxml.tmp`。

## 当前状态与下一步优化（2026-11-19）
- 已完成：多版本 COM 探查、`ooxml_patch`、`ref_diff`、`rubric_grader`（23 类规则、`init` 抽 10 类）、`README.md` 流程与 `--no-com` 边界勘误、`AGENTS.md` 规则文件（14 文件已同步 `main`）。
- 未闭合：GitHub 网页端 `About` 简介/Topics 仍为旧版文案（GitHub MCP 无修改仓库元数据接口）。
- 下一步优化候选：
  1. 建 `tests/` 离线单测与脱敏 `.docx/.xlsx/.pptx` fixture，先锁住 23 类规则与 `ooxml_patch` 回归；
  2. 引入 `openpyxl` + XML 解析补齐 `rubric_grader` / `ref_diff` 全量 `--no-com`（消除 9 类强依赖 COM 规则）；
  3. 扩充 `rubric_grader.py init` 规则自动提取（样式/公式/图形位置）与 `batch` 通用目录匹配；
  4. `ooxml_patch` 支持表格/文本框内 `w14` 特效补丁，以及 `SKILL.md` 兼容 `.claude/skills`/`.codex/skills` 路径；
  5. 为本地目录关联 `.git` 工作副本（或通过 `git clone` 接管），后续直接走标准 Git 提交。

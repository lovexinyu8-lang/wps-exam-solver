---
name: wps-exam-solver
description: "专用于全国计算机等级考试（NCRE）二级 WPS Office、KOS 认证及各类办公自动化考证操作题的判分级求解器与自检引擎。当用户提供 WPS 考试题目截图、要求完成 .docx/.wps/.xlsx/.et/.pptx/.dps 作答文件修改、或要求对照参考答案按考试评分标准验卷打分时强制触发。采用 Windows 原生动态 COM（win32com.client.dynamic 调度 Kwps/Ket/Kwpp）+ OOXML 底层外科补丁与双轨审计，严禁裸拆 XML 直接交卷。"
---

# WPS 考试判分级求解器 (WPS Exam Solver)

本 Skill 面向 Windows 本机 WPS Office（覆盖 `Kwps`/`Ket`/`Kwpp` 及旧版 `WPS`/`ET`/`WPP` ProgID），以**单轮闭环 + 自动备份回滚 + COM/OOXML 双轨自检**对齐考试机阅卷评分标准。

> **模块地图**
> - 运行时基础设施：[scripts/wps_env.py](scripts/wps_env.py) — 多版本 ProgID 探测、`dynamic.Dispatch`、写锁预检、自动备份与进程隔离关闭
> - OOXML 补丁与审计：[scripts/ooxml_patch.py](scripts/ooxml_patch.py) — WPS COM 盲区外科注入 + 只读 XML 结构提取
> - 判分级比对器：[scripts/ref_diff.py](scripts/ref_diff.py) — `AuditSuite` 断言器 + 作答文件 vs 参考答案双轨打分 CLI
> - 分值规则判分引擎：[scripts/rubric_grader.py](scripts/rubric_grader.py) — JSON 权重评分（`grade`）、参考答案自动出规则（`init`）、整套目录批量阅卷（`batch`）
> - 专项踩坑字典：[Word (Kwps)](references/word-kwps.md) | [Excel (Ket)](references/excel-ket.md) | [PPT (Kwpp)](references/ppt-kwpp.md) | [评分系统判分标准](references/grading-standard.md)

---

## 一、 核心执行铁律

1. **单轮闭环 + 强制前置备份**：
   - 整合所有小问生成单份完整 `solve.py` 一次性执行，禁止多回合试探。
   - 所有写入脚本必须通过 `wps_env.wps_session(kind, file_path, backup=True)` 打开，自动完成 `r+b` 写锁预检、创建 `.bak` 备份并在异常时回滚。
2. **严禁裸拆 XML 交卷（COM 为主、XML 为补、WPS 终保存）**：
   - 必须通过 WPS 原生 COM 落盘；极少数 COM 无法写出的专属节点（`w14:textFill`、`w14:glow` `satMod`、`w:themeFill`、`wpsCustomData:diagonals`、PPT `presProps/prnPr`、PPT `hMerge` 对齐），调用 `scripts/ooxml_patch.py` 补充后，**必须再用 `wps_session` 重新打开并 `Save()` 一次**。
   - 若文件为旧版 OLE2 二进制格式（`.doc`/`.xls`/`.wps` 非 ZIP 容器，`inspect_container` 返回 `ole2`），严禁调用 `zipfile` 改包。
3. **最多 1 次全量只读探查**：
   - 遇陌生模板或复杂表格结构，只允许执行 **1 个**只读探查脚本一次性导出完整段落/表结构/页码/XML 关键信息；严禁切片截断或因未加 `try/except` 中途中断而反复重探。
4. **多版本兼容与防卡死/防误杀铁律**：
   - **强制使用 `win32com.client.dynamic.Dispatch`**（已内置于 `wps_session`），绕过 `AppData\Local\Temp\gen_py` 缓存与双装 MS Office 的类型库碰撞。
   - **严禁按空标题全局 `Stop-Process` 杀 `wps`/`et`/`wpp`**：考试端或用户预览窗口的 `MainWindowTitle` 也常为空！`wps_session` 通过记录 `pre_count` 保证仅在自己新拉起空白实例时调用 `Quit()`，绝不误杀用户前台窗口。
   - **只读态严禁调用 `ChartData.Activate()`**：在 `ReadOnly=True` 下激活内嵌图表会死锁表格进程！校验图表数据点与系列一律通过 `ooxml_patch.audit_ooxml()` 读取 `chart*.xml`。
5. **双轨强制准出 & 真实耗时汇报**：
   - 若用户提供了**参考答案**，直接运行 `python scripts/ref_diff.py <作答文件> <参考答案>` 获取全维评分报告。
   - 若无参考答案，必须执行独立 `check.py`（结合 `wps_session(..., read_only=True)` + `audit_ooxml()` + `AuditSuite`），对题干全部硬性指标逐项断言全 `PASS` 后方可交付，并如实汇报包含探查、执行与复核在内的完整耗时。

---

## 二、 先读专项字典再写代码

在动手写 `solve.py` 之前，按当前题目类型**按需读取对应参考文档**（已包含完整参数常量与避坑代码）：
- Word（`.docx` / `.wps`）：读 [references/word-kwps.md](references/word-kwps.md)
- Excel（`.xlsx` / `.et`）：读 [references/excel-ket.md](references/excel-ket.md)
- PPT（`.pptx` / `.dps`）：读 [references/ppt-kwpp.md](references/ppt-kwpp.md)
- 核验或对答案时：读 [references/grading-standard.md](references/grading-standard.md)

首次在陌生电脑运行前，可先执行环境体检确认 COM 注册表与路径状态：
```bash
python scripts/wps_env.py --probe <可选:目标文件路径>
```

---

## 三、 标准求解与自检骨架

### 1. 求解骨架（`solve.py`）
```python
import os, sys, time
from pathlib import Path
sys.path.insert(0, str(next(
    p / "scripts" for p in (Path.cwd() / ".agents/skills/wps-exam-solver", Path.home() / ".agents/skills/wps-exam-solver")
    if (p / "scripts").is_dir()
)))
from wps_env import wps_session, infer_kind
from ooxml_patch import cm_to_pt, rgb_to_bgr_int  # 按需引入 XML 补丁函数

t0 = time.time()
file_path = os.path.abspath(r"...")
kind = infer_kind(file_path)  # 'word' | 'excel' | 'ppt'

with wps_session(kind, file_path, backup=True, with_window=False) as (app, doc_obj, meta):
    # --- 单轮执行全部小问（涉及 PPT SmartArt.AllNodes.Delete() 时改 with_window=True） ---
    if kind == "excel":
        app.Calculate()
    doc_obj.Save()

# 若使用了 ooxml_patch.* 注入底层节点，补一次 WPS 原生重新打开保存：
# with wps_session(kind, file_path) as (app, doc_obj, _):
#     doc_obj.Save()

print(f"SOLVE DONE in {time.time() - t0:.2f}s via {meta['progid']}")
```

### 2. 无参考答案时的双轨自检骨架（`check.py`）
```python
import os, sys
from pathlib import Path
sys.path.insert(0, str(next(
    p / "scripts" for p in (Path.cwd() / ".agents/skills/wps-exam-solver", Path.home() / ".agents/skills/wps-exam-solver")
    if (p / "scripts").is_dir()
)))
from wps_env import wps_session, infer_kind
from ooxml_patch import audit_ooxml
from ref_diff import AuditSuite

file_path = os.path.abspath(r"...")
suite = AuditSuite(os.path.basename(file_path))
xml_info = audit_ooxml(file_path)  # 只读审校图表 ptCount / bodyPr / numbering / cf / sortState

with wps_session(infer_kind(file_path), file_path, read_only=True) as (app, doc_obj, _):
    if infer_kind(file_path) == "word":
        doc_obj.Repaginate()
    # suite.check("指标名", 布尔条件, 实测详情)

print(suite.format_report())
sys.exit(0 if suite.all_passed else 1)
```

### 3. 有参考答案时的判分级比对（命令行一键验卷）
```bash
python scripts/ref_diff.py "<作答文件>" "<参考答案>"
```

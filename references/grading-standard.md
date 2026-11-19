# 考试评分系统判分标准与双轨自检规范

> 培训学校模考软件与 NCRE/KOS 官方阅卷系统的核心判分逻辑，是**“COM 渲染/对象树属性 + 底层 OOXML 节点属性”双重判定**。肉眼在界面上“看起来一样”但底层节点不对，机器阅卷仍会判 0 分。

---

## 1. 常见“看起来对、机器判 0 分”高危清单

| 模块 | 错误做法（肉眼正常但判 0 分） | 阅卷系统真实检查点 | 正确解法 |
| :--- | :--- | :--- | :--- |
| **PPT 文本框顶端对齐** | 只设 `TextFrame.VerticalAnchor = 1` | `<a:bodyPr anchor="t" anchorCtr="0"/>` | 必须用 `TextFrame2` 同时设 `VerticalAnchor=1` 与 `HorizontalAnchor=1` |
| **PPT 打印设置** | 只设 `prs.PrintOptions.*` 后 `Save()` | `ppt/presProps.xml` 中 `<p:prnPr>` 属性 | 必须用 `ooxml_patch.patch_ppt_print_options()` 写盘并 `Kwpp` 重存 |
| **PPT 表格合并居中** | 只对可见首格设居中 | 包括 `hMerge="1"` 隐藏格在内的所有 `<a:pPr algn="ctr"/>` | 合并前后居中 + `patch_ppt_hmerge_align()` |
| **PPT 宽屏“确保适合”** | 只改 `SlideWidth=960, SlideHeight=540` | 母版/版式占位符比例与内容页居中坐标 | 同步按 `960/720` 缩放母版与版式并平移内容页 |
| **Word 页眉图片水印** | 直接插进页眉或未设 `WrapFormat` | `header*.xml` 中 `behindDoc="1"` 且 `sect2` 断开链接 | `LinkToPrevious=False` + `WrapFormat.Type=5` + 传 `Anchor=header.Range` |
| **Word 发光/渐变特效** | 只设 `Font.Glow` 未补 XML | `<w14:glow>` 含 `<w14:satMod w14:val="175000"/>`、`<w14:textFill>` | `patch_word_glow()` / `patch_word_grad_fill()` + `Kwps` 重存 |
| **Word 表格主题底纹** | 只设 `Shading.BackgroundPatternColor` | `<w:shd w:themeFill="accent2" .../>` | COM 设颜色后补 `patch_word_table_theme_fill()` + `Kwps` 重存 |
| **Word 多级列表** | 手写“一、”“（一）”文本 | `numbering.xml` 中 `chineseCounting`、`lvlText`、`lvlRestart` | `ListTemplates.Add(True)` 绑定内建样式并实测校验 `ListString` |
| **Word/PPT 图表** | 直接填超 4 行数据不扩容 | `chart*.xml` 中 `<c:ptCount val="N"/>` 与 `<c:ser>` 数量 | `ChartData.Activate()` 后先 `Rows(5).Insert()` 扩行再填数与删多余列 |
| **Excel 条件格式** | 未选中区域直接 `FormatConditions.Add` | `sheet*.xml` `<cfRule><formula>` 中的行号是否偏移 | 必须先 `ws.Activate(); rng.Select()` 锚定 `ActiveCell` 到区域左上角 |
| **Excel 多字段排序** | Python 排序后回写值 | `sheet*.xml` 是否存在 `<sortState ref="...">` | 必须调用 `ws.Sort`（`SortMethod=1` 拼音排序） |
| **`.wps` 扩展名保存** | 调用 `doc.SaveAs2(path)` 不传格式码 | 文件头是 OLE2 还是 ZIP (`SaveFormat`) | 必须显式传 `doc.SaveAs2(path, 0)` |

---

## 2. 双轨自检与“参考答案比对器”使用方式

### 方式 A：单题自检（无参考答案时，由 `check.py` 实例化 `AuditSuite`）
```python
import os, sys
from pathlib import Path
_candidates = [
    *( [Path(os.environ["WPS_EXAM_SOLVER_HOME"])] if os.environ.get("WPS_EXAM_SOLVER_HOME") else [] ),
    *(b / r for b in (Path.cwd(), Path.home()) for r in (".agents/skills/wps-exam-solver", ".claude/skills/wps-exam-solver", ".codex/skills/wps-exam-solver")),
]
sys.path.insert(0, str(next(p / "scripts" for p in _candidates if (p / "scripts").is_dir())))
from wps_env import wps_session
from ooxml_patch import audit_ooxml
from ref_diff import AuditSuite

suite = AuditSuite("题目自检")
xml_info = audit_ooxml(file_path)  # 只读审校图表 ptCount / numbering / bodyPr / cf / sortState

with wps_session("word", file_path, read_only=True) as (app, doc, _):
    doc.Repaginate()
    # 对题干每一项硬性指标调用 suite.check(name, cond, detail)
print(suite.format_report())
assert suite.all_passed
```

### 方式 B：有参考答案时，直接做判分级全维对比（培训软件平替核心）
```bash
# 双轨对比（COM + OOXML），返回退出码 0（全通过）或 1（存在差异）
python scripts/ref_diff.py <作答文件> <参考答案>

# 输出结构化 JSON 供脚本/评测前端消费（含 location / expected / actual / hint）
python scripts/ref_diff.py <作答文件> <参考答案> --json

# 纯 OOXML + openpyxl 静态对比（无需安装 WPS，毫秒级）
python scripts/ref_diff.py <作答文件> <参考答案> --no-com
```

### 方式 C：JSON 分值规则判分、自动出规则与批量阅卷
```bash
# 1) 按 JSON 规则文件打分（支持单文件分点计分）
python scripts/rubric_grader.py grade <作答文件> examples/ppt_sample.rubric.json

# 2) 从已有参考答案提取 JSON 评分规则骨架（支持 COM 与 --no-com 静态抽取）
python scripts/rubric_grader.py init <参考答案> -o my_exam.rubric.json --total-points 20
python scripts/rubric_grader.py init <参考答案> -o my_exam.rubric.json --no-com

# 3) 整班/整套模考批量自动判分（支持：目录镜像、全班对单份参考答案/规则文件），并可导出 CSV/XLSX
python scripts/rubric_grader.py batch <作答目录> <参考答案或规则文件/目录> --export report.xlsx
```

#### 支持的 `rubric.json` 检查类型（`type`）速查（共 23 类，全部提供 OOXML/`openpyxl` `--no-com` 实现）
- **Word (9)**：`word_page_setup`、`word_style`、`word_paragraph_match`、`word_text_absent`、`word_header_watermark`、`word_numbering_xml`、`word_table`、`word_effects_xml`、`word_chart`
- **Excel (7)**：`excel_sheets`、`excel_range_formula`、`excel_range_values`、`excel_number_format`、`excel_conditional_formatting`、`excel_sort_state`、`excel_pivot_table`
- **PPT (7)**：`ppt_page_setup`、`ppt_print_options`、`ppt_master_background`、`ppt_slide_transition`、`ppt_shape_match`、`ppt_align_distribute`、`ppt_chart`

> **两条已知边界提示**：
> 1. 在 `--no-com` 纯静态模式下，`excel_range_values` 依赖工作簿内保存的 XML 缓存值（若公式未在 WPS/Excel 中保存计算缓存会提示 `cached value missing`），Word 动态分页/渲染编号文本（`page_number`、`ListString` 动态计算）仍以 WPS COM 模式为准。
> 2. 旧版二进制 OLE2 文件（`.doc` / `.xls` / `.wps` / `.et` 非 ZIP 容器）无 OOXML 结构，必须移除 `--no-com` 走 WPS COM 核验。

---
name: wps-exam-solver
description: "专用于全国计算机等级考试（NCRE）二级 WPS Office、WPS 认证考试及各类办公自动化考证操作题的秒级求解器。当用户提供 WPS/Office 考试题目截图、要求完成作答文件中的 .pptx、.docx、.wps、.xlsx、.et 文件修改、或要求做题/交卷时强制触发。该 Skill 采用 Windows 原生 COM 自动化（win32com.client 调度 kwpp/kwps/ket），单轮脚本直接完成所有小问，严禁拆包改 XML。"
---

# WPS 考试极速求解器 (WPS Exam Solver)

本 Skill 是全国计算机等级考试（NCRE）二级 WPS Office 及各类办公自动化考证题目的**唯一权威、免调试、高分求解器**。
通过统一调度 Windows 本机原生 WPS COM 接口（WPS 12.0 内核），实现单轮脚本 **2~6 秒极速通关**，100% 契合阅卷系统的机器判分标准。

---

## 一、 核心执行铁律

1. **单轮闭环（Single-Pass Execution）**：
   - 看到题目后，直接整合所有小问逻辑，生成单份完整 Python 脚本一次性执行完毕，禁止多回合拉扯。
2. **严禁解包改 XML（必须由 WPS COM 原生闭环）**：
   - 阅卷系统依赖 WPS COM 对象模型与标准 OOXML 判分，禁止使用 `zipfile`/`unpack.py` 改包后直接交付。必须调度原生 COM：
     - **WPS 演示 (PPT)**：`win32com.client.Dispatch("Kwpp.Application")`
     - **WPS 文字 (Word)**：`win32com.client.Dispatch("Kwps.Application")`
     - **WPS 表格 (Excel)**：`win32com.client.Dispatch("Ket.Application")`
   - **例外补偿原则**：极少数无 COM 写入接口的 WPS 专属节点（如 Word `w14:textFill` 多色标渐变、WPS 原生双斜线表头 `wpsCustomData:diagonals`），可由 `python-docx` `parse_xml` 局部补充，但**最后一次打开与 `Save()` 必须由 `Kwps.Application` 执行**。
3. **严格遵守执行规范，严禁模型擅自发挥与盲目试探**：
   - **最多 1 次全量只读探查**：遇陌生模板/复杂表结构，只允许用 **1 个** 无 COM（`python-docx`/`openpyxl`/`zipfile`）或轻量只读脚本一次性完整导出结构，禁止切片截断输出、禁止未加 `try/except` 导致中途中断而反复重探，杜绝把探查拆成多次往返拖垮总耗时。
   - **强力二次自检（强制准出）**：作答脚本执行后，必须立即触发 `ReadOnly=True` 模式二次检验脚本，对题干每一项硬性指标编写布尔断言；未达成全部 `PASS` 严禁向用户交付。
   - **真实耗时透明汇报**：如实汇报从接受指令到完成作答与校验的全部物理耗时（含探查与返工），杜绝统计口径造假。
4. **防卡死、防弹窗与进程保护铁律**：
   - **操作前 1ms 写锁预检**：调 COM 打开作答文件前，必须先执行 `with open(file_path, "r+b"): pass`；若抛 `PermissionError`，立刻停下报告文件被占用（检查 `~$...` 锁文件），**严禁强行调 COM 死等锁弹窗**。
   - **严禁全局 `Stop-Process` 误杀考试页面**：WPS Office 多进程架构下，用户前台开着的考试页面或预览进程在 `Get-Process` 中 `MainWindowTitle` 也为空！**绝对禁止对 `wps` / `et` / `wpp` / `WINWORD` 按空标题或全局执行 `Stop-Process`**。
   - **禁止只读态激活内嵌图表**：在 `ReadOnly=True` 且 `Visible=False` 下对 Word 图表调用 `ChartData.Activate()`（甚至遍历 `SeriesCollection`）会死锁内嵌表格进程！只读探查与校验图表数据时，直接读取 `word/charts/chart*.xml` 或只读 `Chart.ChartType` / `ChartTitle.Text`。
   - **防弹窗挂起与安全退出**：所有 COM 实例必设 `app.Visible = False` 与 `app.DisplayAlerts = False`（PPT/Excel 同理），且必须用 `try...finally`：`if doc_obj: doc_obj.Close(SaveChanges=False)` 后再 `app.Quit()`（切勿裸调用 `Close()` 触发“是否保存”弹窗）。
   - **PPT 动作按钮 & 图表数据源**：PPT 严禁直接 `AddShape(138)`（触发超链接弹窗锁死），须先建矩形再改 `AutoShapeType`；图表写入完毕务必 `try: wb.Close(False)`。

---

## 二、 WPS 演示 (PPT) 核心考点与防踩坑字典 (`Kwpp`)

### 1. 贴边与对齐计算
```python
W = prs.PageSetup.SlideWidth
H = prs.PageSetup.SlideHeight

pic.Left = W - pic.Width
pic.Top = 0

# 文本框对齐与锚定（关键坑：必须用 TextFrame2 并同时设 HorizontalAnchor = 1！）
tf2 = shape.TextFrame2
tf2.VerticalAnchor = 1    # 1=顶端对齐, 3=垂直居中, 4=底端对齐
tf2.HorizontalAnchor = 1  # 必须设 1 (msoAnchorNone -> anchorCtr="0")，否则单设 TextFrame.VerticalAnchor=1 缺 anchorCtr="0" 会扣分！
shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2  # 1=左, 2=居中, 3=右
```

### 2. 字体特效（小型大写字母、艺术字图案填充、阴影）
```python
shape.TextFrame2.TextRange.Font.Smallcaps = -1  # msoTrue
font2 = shape.TextFrame2.TextRange.Font
font2.Fill.Patterned(16)  # 16=深色上对角线 (dkUpDiag)
font2.Fill.ForeColor.ObjectThemeColor = 14; font2.Fill.ForeColor.Brightness = -0.5  # bg1 lumMod=50%
font2.Fill.BackColor.ObjectThemeColor = 13; font2.Fill.BackColor.Brightness = 0.25  # tx1 lumMod=75% lumOff=25%
sh = font2.Shadow
sh.Type = 21; sh.OffsetX = 1.06066; sh.OffsetY = 1.06066; sh.Transparency = 0.6
sh.ForeColor.ObjectThemeColor = 1; sh.ForeColor.Brightness = -0.5
shape.TextFrame2.ThreeD.SetPresetCamera(19)
```

### 3. 智能图形 (SmartArt) 稳定转换与节点控制
> **注意**：WPS COM 对 `AllNodes.Add` 有底层限制，严禁盲目调用 `Add()`；且**操作 `SmartArt.AllNodes.Delete()` 时必须以 `WithWindow=True` 打开演示文稿**，否则 `WithWindow=False` 会直接抛出 `-2147417851 服务器出现意外情况` 崩溃！删除原形状时若涉及多个形状，先记录对象或按名称定位，防止索引偏移。
```python
layout = next((app.SmartArtLayouts(i) for i in range(1, app.SmartArtLayouts.Count + 1)
               if "垂直项目符号列表" in app.SmartArtLayouts(i).Name or "Vertical Bullet" in app.SmartArtLayouts(i).Name), None)
sa = slide.Shapes.AddSmartArt(layout, rect.Left, rect.Top, rect.Width, rect.Height)
S = sa.SmartArt
if S.AllNodes.Count >= 4:
    S.AllNodes(4).Delete()
    S.AllNodes(2).Promote()
for i, item_text in enumerate(["产品范围", "发展趋势", "Q & A"], 1):
    if i <= S.AllNodes.Count:
        S.AllNodes(i).TextFrame2.TextRange.Text = item_text
rect.Delete()
```

### 4. 表格转图表 (Chart) 与数据区动态扩容
> **注意**：`AddChart2` 默认模板仅 4 行，第 5 行以上必须先 `Rows(5).Insert()` 撑开；`SetSourceData` 传字符串地址（勿跨进程直接传 Excel `Range` 对象）；关闭内部 Workbook 必传 `False`。
```python
ch = slide.Shapes.AddChart2(-1, 51, tb.Left, tb.Top, tb.Width, tb.Height)
c = ch.Chart
c.ChartData.Activate()
wb = c.ChartData.Workbook
ws = wb.Worksheets(1)
if len(data) > 5:
    for r in range(5, len(data)):
        ws.Rows(r).Insert()
for r_idx, row in enumerate(data, 1):
    for c_idx, val in enumerate(row, 1):
        ws.Cells(r_idx, c_idx).Value = val
c.SetSourceData(f"='Sheet1'!$A$1:$C${len(data)}")
try:
    wb.Close(False)
except Exception:
    pass
c.HasTitle = False
c.Axes(1).HasMajorGridlines = False
c.Axes(2).HasMajorGridlines = False
tb.Delete()
```

### 5. 母版/页面大小/对齐分布/无窗口坑（实测重灾区）
- **宽屏 16:9 + “确保适合”坑**：`SlideSize = 15` 在 `Kwpp` 中会缩成 `720×405pt`；现代 16:9 宽屏必须设 `SlideSize = 7; SlideWidth = 960.0; SlideHeight = 540.0`。**若题干要求“确保适合”，仅改宽高会丢“确保适合”分！必须同时把 `SlideMaster.Shapes` 与 `CustomLayouts` 占位符按 `ratio = 960/720` 缩放 `Left/Width`，并把非首页内容页图形水平平移 `+(960-720)/2`（`+120pt`）**。
- **母版背景 vs 第1页**：首页常默认 `FollowMasterBackground = 0`，只设 `SlideMaster.Background.Fill.UserPicture(...)` 会漏掉第 1 页！必须显式循环 `for sl in prs.Slides: sl.FollowMasterBackground = -1`。
- **严禁依赖无头 `ShapeRange.Align` / `Distribute`**：`WithWindow=False` 下 `sr.Align(0, 0)` 会静默失效、`sr.Distribute(0, 0)` 会把全部文本框叠在一起（间距成负数）！**对齐与均匀分布一律直接用几何公式赋值 `sp.Top` 与 `sp.Left`**。
- **删除/遍历与 `WithWindow`**：删除多个原形状先保存引用再删，防集合索引前移；涉及 `ActiveWindow` 视图操作时改为 `WithWindow=True`。

---

## 三、 WPS 文字 (Word) 核心考点与防踩坑字典 (`Kwps`)

### 1. 页面设置、页眉横线、日期域与水印（避坑重点）
```python
ps = doc.PageSetup
m = wps.CentimetersToPoints(2.5)
ps.TopMargin = ps.BottomMargin = ps.LeftMargin = ps.RightMargin = m
ps.GutterPos = 0  # 0=左装订线, 1=上装订线

# 背景色（主题着色4浅色80%）: 必须先调 Solid() 否则不写盘!
doc.Background.Fill.Solid()
doc.Background.Fill.Visible = -1
doc.Background.Fill.ForeColor.ObjectThemeColor = 7  # 7=accent4
doc.Background.Fill.ForeColor.Brightness = 0.8      # 浅色80%

# 页眉域与横线 + 水印顺序:
# 1) 必须先插域与边框，因为 Fields.Add(header.Range, ...) 会清空 header.Range 内已有形状!
# 2) Type=32 (wdFieldTime) 会自动注入 TIME，Code 只需传 '\@ "yyyy年M月d日"'
# 3) AddTextEffect 必须传第9个参数 Anchor=header.Range，否则水印会插进正文第1段冲掉标题边框!
header = doc.Sections(1).Headers(1)  # wdHeaderFooterPrimary
header.Range.Fields.Add(header.Range, 32, r'\@ "yyyy年M月d日"', False)
hp = header.Range.Paragraphs(1)
hp.Borders(-3).LineStyle = 7  # -3=底边框, 7=双实线
hp.Borders(-3).LineWidth = 8  # 1磅

wm = header.Shapes.AddTextEffect(0, "水印文字", "微软雅黑", 36, False, False, 0, 0, header.Range)
wm.Name = "PowerPlusWaterMarkObject"
wm.Rotation = 315  # 倾斜版式
wm.Fill.Visible = True; wm.Fill.Solid(); wm.Fill.ForeColor.RGB = 12632256; wm.Fill.Transparency = 0.5
wm.Line.Visible = False; wm.LockAspectRatio = True
wm.WrapFormat.Type = 3; wm.RelativeHorizontalPosition = 0; wm.RelativeVerticalPosition = 0
wm.Left = -999995; wm.Top = -999995
```

### 2. 段落行距、字符间距、文本边框与文字效果
```python
# 行距常量大坑: 4=固定值 (wdLineSpaceExactly), 5=多倍行距 (wdLineSpaceMultiple)!
pf = p.Format
pf.LineUnitBefore = 0.5   # 段前0.5行 (勿用 SpaceBeforeUnit)
pf.LineUnitAfter = 0.5    # 段后0.5行
pf.LineSpacingRule = 4    # 固定值=4 -> LineSpacing=18; 若是多倍行距1.25倍则 Rule=5, LineSpacing=wps.LinesToPoints(1.25)
pf.LineSpacing = 18

# 字符间距加宽 0.1cm (Font无 CharacterWidth)
p.Range.Font.Spacing = wps.CentimetersToPoints(0.1)

# 文本边框与底纹 (必须排除段落末尾 \r，且设在 Font 上而非 Paragraph 上!)
rng_t = doc.Range(p1.Range.Start, p1.Range.End - 1)
for bi in [1, 2, 3, 4]:
    b = rng_t.Font.Borders(bi)
    b.LineStyle = 7       # 7=双实线
    b.LineWidth = 12      # 12=1.5磅 (eighths of pt)
    b.Color = 10498160    # 紫色 RGB(112,48,160) -> BGR 整数
rng_t.Font.Shading.Texture = -9  # -9=浅色下斜线 (thinReverseDiagStripe)

# 文字效果 (阴影退化与发光饱和度坑!)
f = rng.Font
f.TextShadow.Type = 39          # 左上对角透视
f.TextShadow.Blur = 4.7249608   # 必须显式设 Blur 与 Transparency，否则重开退化为 -2!
f.TextShadow.Transparency = 0.68
f.Reflection.Type = 5           # 半倒影 4pt偏移量
f.Line.Visible = -1; f.Line.Style = 1; f.Line.DashStyle = 8; f.Line.ForeColor.RGB = 255; f.Line.Weight = 1.0
f.Glow.Radius = 18; f.Glow.Color.ObjectThemeColor = 7; f.Glow.Transparency = 0.6  # 发光必设 Transparency=0.6，且需补充 w14:satMod="175000" 再由 WPS 重存!
```

### 3. 查找替换、批注与图片形状裁剪
```python
# 批注锚点在目标词后
rng_c = doc.Content
if rng_c.Find.Execute("费用小结："):
    rng_c.Collapse(0)  # wdCollapseEnd=0
    doc.Comments.Add(Range=rng_c, Text="未含住宿费")

# 查找替换（Avoid Find.Execute(Replace=2) 静默失效坑：用显式 while 循环替换！）
while True:
    rf = doc.Content
    rf.Find.ClearFormatting()
    if not rf.Find.Execute("通信"):
        break
    rf.Text = "相信"
    rf.Font.EmphasisMark = 1      # "." 着重号
    rf.Font.Underline = 43        # 双波浪下划线
    rf.Font.UnderlineColor = 3942550  # RGB(150,40,60) BGR整数

# 图片调尺寸与椭圆裁剪: 先调不需要转换的图尺寸，再转Shape裁剪，防止索引偏移!
is2 = doc.InlineShapes(2)
is2.LockAspectRatio = 0  # 修改宽高前必关锁定纵横比!
is2.Width = wps.CentimetersToPoints(6.4); is2.Height = wps.CentimetersToPoints(4.2)
sh1 = doc.InlineShapes(1).ConvertToShape()
sh1.AutoShapeType = 9    # 9=椭圆 (msoShapeOval), 2=平行四边形
sh1.ConvertToInlineShape()
```

### 4. 表格转换、合并顺序、网格表1与内嵌图表
- **分隔符**：`rng.ConvertToTable(Separator=1, ...)` — **`1` = 制表符（`wdSeparateByTabs`）**，`0` = 段落。
- **合并前后硬规则**：
  1. **必须在 `Merge()` 之前设完行高与列宽**（列宽用 `tbl.AllowAutoFit = False; tbl.Columns(c).SetWidth(w, 0)`）；
  2. 一旦发生纵向合并，`tbl.Rows(r)` 会直接抛 `-2146822297`！合并后**一律改用 `tbl.Cell(r, c)`**，校验字体时读 `tbl.Cell(r, c).Range.Characters(1).Font.Name`（避免行末 `\r\x07` 导致 `NameFarEast` 返回空串）；
  3. 某行改过 3 磅粗线（`LineWidth=24`）后，全表 `tbl.Borders(-5)` 会变为混合值 `9999999`，自检须分别断言各 `Cell` 上/下边框。
- **WPS“网格表1”与主题底纹槽位**：`Styles` 中实际名为 `"网格型"`，再配合 `ApplyStyleHeadingRows = True; ApplyStyleRowBands = True` + 首行单元格 `#3F3F3F` 粗体白字 + 偶数行 `#8B8B8B` 白字 + 奇数行 `#FFFFFF` 黑字；底纹必须逐个 `tbl.Cell(r, c).Shading.BackgroundPatternColor` 设置（`Rows(1).Shading` 无效）。**若题干指定主题底纹颜色（如“橙色，着色2”），COM 无法写入 `w:themeFill`，须在 `<w:shd>` 上补 `w:themeFill="accent2"` 再用 `Kwps` 重新打开 `Save()`**。
- **Word 插入图表锚点与数据区扩容坑**：调用 `doc.InlineShapes.AddChart2(-1, 57, rng_ins)` 前**必须先 `rng_ins.Select()`**，否则常锚回第 1 页；**Word 图表默认仅 4 行数据区，`SetSourceData` 无法自动撑开，必须在激活 `ChartData` 后用 `ws.Rows(5).Insert()` 插入 `N - 4` 行再填数、删多余列与系列并 `wb.Close(True)`**！校验时必先 `doc.Repaginate()` 并读取 `word/charts/chart1.xml` 校验实际 `<c:ptCount>`。
- **Word 垂直文字方向旋转页面坑**：`PageSetup` 无 `TextOrientation` 属性，文字方向在 `doc.Sections(1).Range.Orientation`（`11` = 垂直从左往右 `tbLrV`）；设置它会**自动将页面方向、宽高和页边距旋转 90°**，故必须**先设 `Range.Orientation` 再设 `PageSetup` 横向/尺寸/边距**！

---

## 四、 WPS 表格 (Excel) 核心考点与防踩坑字典 (`Ket`)

```python
import win32com.client, os

et = win32com.client.Dispatch("Ket.Application")
et.Visible = False
et.DisplayAlerts = False
wb = None
try:
    wb = et.Workbooks.Open(os.path.abspath(xlsx_path))
    ws = wb.Sheets(1)

    # 1. 公式必须使用英文函数名 + 英文逗号，且务必先查清数据末行（防把末行“汇总”算进 MAX/MIN/COUNTIF）与真实列号
    ws.Range("G2:G100").Formula = "=VLOOKUP(A2, 码表!A:D, 3, FALSE)"
    ws.Range("H2:H100").Formula = "=SUMIFS(数据!C:C, 数据!A:A, A2, 数据!B:B, B2)"
    ws.Range("I2:I100").Formula = '=IF(OR(F2>10,H2>20000),"优","一般")'

    # 2. TextToColumns 防静默失效：在无头 Ket 中按固定位置分列经常不报错但B列仍为空！
    #    对“日期+公司名”类混合列，直接 Python 正则拆分写回 A/B 列并设 NumberFormat="yyyy/m/d"

    # 3. 透视表刷新与经典行标签布局
    #    pc = wb.PivotCaches().Create(1, "'销售数据'!$A$1:$D$122")
    #    pt = pc.CreatePivotTable(sh.Range("A3"), "数据透视表1")
    #    pt.PivotFields("客户级别").Orientation = 1
    #    pt.AddDataField(pt.PivotFields("公司名称"), "计数项:公司名称", -4112)  # -4112=xlCount
    #    pt.RowAxisLayout(1)  # 避免显示“行标签”，改为字段名
    #    pt.PivotCache().Refresh(); pt.RefreshTable()

    # 4. 会计专用格式与边框
    ws.Range("C2:C122").NumberFormat = '_ "¥"* #,##0.00_ ;_ "¥"* -#,##0.00_ ;_ "¥"* "-"??_ ;_ @_ '
    rng = ws.Range("A1:H100")
    rng.HorizontalAlignment = 3  # xlCenter
    rng.VerticalAlignment = 2
    for border_id in [7, 8, 9, 10, 11, 12]:
        rng.Borders(border_id).LineStyle = 1
        rng.Borders(border_id).Weight = 2
    rng.EntireColumn.AutoFit()

    wb.Save()
finally:
    if wb:
        try: wb.Close(SaveChanges=False)
        except: pass
    et.Quit()
```

---

## 五、 秒级作答与二次自检标准骨架 (Standard Boilerplate)

### 1. 作答脚本骨架（`solve.py`）
```python
import win32com.client, os, time, sys

t0 = time.time()
file_path = os.path.abspath(r"...")
with open(file_path, "r+b"):  # 1ms 写锁预检，防文件占用触发 COM 弹窗挂起
    pass

app = win32com.client.Dispatch("Kwps.Application")  # 按题型选 Kwps / Ket / Kwpp
app.Visible = False
app.DisplayAlerts = False
doc_obj = None

try:
    doc_obj = app.Documents.Open(file_path)  # Excel: Workbooks.Open | PPT: Presentations.Open(file_path, False, False, False)
    # --- 单段脚本执行全部小问 ---
    doc_obj.Save()
    print(f"SOLVE DONE in {time.time() - t0:.2f}s")
except Exception as e:
    print(f"ERROR: {e}", file=sys.stderr)
    raise
finally:
    if doc_obj:
        try: doc_obj.Close(SaveChanges=False)
        except: pass
    app.Quit()
```

### 2. 二次只读自检骨架（`check.py`，全 PASS 方可交付）
```python
import win32com.client, os, time

t0 = time.time()
file_path = os.path.abspath(r"...")
app = win32com.client.Dispatch("Kwps.Application")
app.Visible = False
app.DisplayAlerts = False
doc_obj = app.Documents.Open(file_path, ReadOnly=True)
res = []

def check(name, cond, detail=""):
    res.append(cond)
    print(f"[{'PASS' if cond else 'FAIL'}] {name}" + (f" ({detail})" if detail else ""))

try:
    # 若涉及 Word 页码断言，必须先调用 doc_obj.Repaginate()
    # 逐项对题干全部硬性指标编写布尔断言，严禁在只读模式调用 ChartData.Activate()
    print(f"CHECK DONE in {time.time() - t0:.2f}s | {'ALL PASS' if all(res) else 'HAS FAIL'}")
finally:
    doc_obj.Close(SaveChanges=False)
    app.Quit()
```

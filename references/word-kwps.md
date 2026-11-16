# WPS 文字 (Kwps) 考点与 COM/OOXML 防踩坑字典

> 所有写入操作必须使用 `scripts/wps_env.py` 中的 `wps_session("word", path, backup=True)`；若调用了 `ooxml_patch.py` 的 XML 补丁，**最后一步必须重新用 `wps_session("word", path)` 打开并执行 `doc.Save()`**。

---

## 1. 页面设置、纸张、文字方向与背景

```python
ps = doc.PageSetup
m = app.CentimetersToPoints(2.5)
ps.TopMargin = ps.BottomMargin = ps.LeftMargin = ps.RightMargin = m
ps.GutterPos = 0  # 0=左装订线, 1=上装订线

# “适中”页边距标准值：上下 2.54cm (1in=72pt)，左右 1.91cm (0.75in=54pt)
ps.TopMargin = ps.BottomMargin = app.InchesToPoints(1.0)
ps.LeftMargin = ps.RightMargin = app.InchesToPoints(0.75)
```

- **垂直文字方向必保顺序**：`PageSetup` 无 `TextOrientation` 属性。必须**先设 `doc.Sections(1).Range.Orientation = 11`（垂直从左往右 `tbLrV`），再设 `PageSetup.Orientation`（横向=`1`）、`PageWidth`/`PageHeight` 和页边距**——因为修改文字方向会自动把页面方向与页边距旋转 90°。
- **页面背景色写盘坑**：必须先调用 `doc.Background.Fill.Solid()`，否则 `ForeColor` 不会落盘：
  ```python
  doc.Background.Fill.Solid()
  doc.Background.Fill.Visible = -1
  doc.Background.Fill.ForeColor.ObjectThemeColor = 7  # 7=accent4
  doc.Background.Fill.ForeColor.Brightness = 0.8      # 浅色 80%
  ```

---

## 2. 分节、页眉域、横线与水印（高危区）

- **下一页分节符与首行空白清理**：
  ```python
  r = doc.Paragraphs(1).Range
  r.MoveEnd(1, -1)
  r.Collapse(0)
  r.InsertBreak(2)  # 2 = wdSectionBreakNextPage
  fp = doc.Sections(2).Range.Paragraphs(1)
  if fp.Range.Text == "\r":
      fp.Range.Delete()
  doc.Sections(2).Headers(1).LinkToPrevious = False  # 仅第1节有水印时必断开链接
  ```
- **页眉操作不可逆顺序**：
  1. **先插日期域与页眉底边框**：因为 `Fields.Add(header.Range, ...)` 会清空 `header.Range` 内已有图形！且 `Type=32`（`wdFieldTime`）会自动注入 `TIME`，`Text` 参数只需传 `r'\@ "yyyy年M月d日"'`（切勿传 `"TIME \\@ ..."` 导致 `TIME Time` 重复）。
  2. **再插水印并强传第 9 参数 `Anchor=header.Range`**：不传 `Anchor` 会导致水印锚进正文第 1 段冲掉标题边框。
  ```python
  header = doc.Sections(1).Headers(1)  # 1 = wdHeaderFooterPrimary
  header.Range.Fields.Add(header.Range, 32, r'\@ "yyyy年M月d日"', False)
  hp = header.Range.Paragraphs(1)
  hp.Borders(-3).LineStyle = 7  # -3=底边框, 7=双实线
  hp.Borders(-3).LineWidth = 8  # 8=1磅 (eighths of pt)

  wm = header.Shapes.AddTextEffect(0, "水印文字", "微软雅黑", 36, False, False, 0, 0, header.Range)
  wm.Name = "PowerPlusWaterMarkObject"
  wm.Rotation = 315
  wm.Fill.Visible = True; wm.Fill.Solid(); wm.Fill.ForeColor.RGB = 12632256; wm.Fill.Transparency = 0.5
  wm.Line.Visible = False; wm.LockAspectRatio = True
  wm.WrapFormat.Type = 3; wm.RelativeHorizontalPosition = 0; wm.RelativeVerticalPosition = 0
  wm.Left = -999995; wm.Top = -999995  # wdShapeCenter
  ```
- **图片水印（取消冲蚀 + 120% 缩放 + `behindDoc="1"`）**：
  ```python
  shp = header.Shapes.AddPicture(FileName=img, LinkToFile=False, SaveWithDocument=True, Anchor=header.Range)
  w0, h0 = shp.Width, shp.Height
  shp.LockAspectRatio = 0
  shp.Width, shp.Height = w0 * 1.2, h0 * 1.2
  shp.LockAspectRatio = -1
  shp.Name = "WordPictureWatermark1"
  shp.PictureFormat.ColorType = 1  # 自动(非冲蚀4)
  shp.PictureFormat.Brightness = 0.5
  shp.PictureFormat.Contrast = 0.5
  shp.WrapFormat.Type = 5          # wdWrapBehind -> behindDoc="1"
  shp.RelativeHorizontalPosition = 0; shp.RelativeVerticalPosition = 0
  shp.Left = -999995; shp.Top = -999995
  ```

---

## 3. 内建样式、多级列表与段落格式

- **用枚举 ID 定位内建样式（跨语言/跨版本稳定）**：
  - 正文 = `doc.Styles(-1)`，标题 1 = `doc.Styles(-2)`，标题 2 = `doc.Styles(-3)`，标题 3 = `doc.Styles(-4)`。
  - **字号磅值速查**：小初=36，一号=26，小一=24，二号=22，小二=18，三号=16，小三=15，四号=14，小四=12，五号=10.5。
- **中文多级列表（一、 / （一） / 1、）防错模板**：
  不同 WPS 版本中文计数 `NumberStyle` 枚举值在 `37~40` 间有差异，绑定后实测 `ListString` 兜底校正是唯一稳妥写法：
  ```python
  h1, h2, h3 = doc.Styles(-2), doc.Styles(-3), doc.Styles(-4)
  lt = doc.ListTemplates.Add(True)  # OutlineNumbered=True
  for lv, fmt, nst, st, rst in [(1, "%1、", 37, h1, 0), (2, "（%2）", 37, h2, 1), (3, "%3、", 0, h3, 2)]:
      L = lt.ListLevels(lv)
      L.NumberFormat = fmt
      L.NumberStyle = nst
      L.StartAt = 1
      try: L.ResetOnHigher = rst
      except Exception: pass
      L.LinkedStyle = st.NameLocal
  for lv, st in [(1, h1), (2, h2), (3, h3)]:
      try: st.LinkToListTemplate(lt, lv)
      except Exception: pass

  # 实测校正 chineseCounting 枚举值
  for alt in (37, 38, 39, 40):
      if h1_para.Range.ListFormat.ListString == "一、" and h2_para.Range.ListFormat.ListString == "（一）":
          break
      lt.ListLevels(1).NumberStyle = alt
      lt.ListLevels(2).NumberStyle = alt
  ```
- **独立字母列表（A、B、C、）**：删除原手动字母后，用 `ltl = doc.ListTemplates.Add(False)` 设 `NumberStyle = 3`、`NumberFormat = "%1、"`，再分别对每组连续段落调用 `rng.ListFormat.ApplyListTemplate(ltl, False, 0)` 确保互相独立不连续编号。
- **行距与缩进常量**：
  - `pf.CharacterUnitFirstLineIndent = 2`（首行缩进 2 字符）
  - `pf.LineUnitBefore = 0.5` / `pf.LineUnitAfter = 0.5`（段前/段后行数，**切勿用 `SpaceBeforeUnit`**；若改用磅值，必先将 `LineUnitBefore = LineUnitAfter = 0`）
  - `pf.LineSpacingRule`：`0`=单倍，`1`=1.5倍，`4`=固定值（`wdLineSpaceExactly`），`5`=多倍行距（`wdLineSpaceMultiple`，配合 `app.LinesToPoints(1.25)`）。

---

## 4. 字体、字符间距、文字效果与拼音指南

- **中西文字体必须同步设**：只设 `Font.Name` 在部分模板下会漏掉东亚或西文 run，务必同时赋值：
  ```python
  f.Name = f.NameFarEast = f.NameAscii = "宋体"
  ```
- **字符间距加宽**：`rng.Font.Spacing = app.CentimetersToPoints(0.1)`。
- **文本边框与底纹**：必须排除段落末尾 `\r`（`doc.Range(p.Range.Start, p.Range.End - 1)`），且设在 `.Font.Borders` / `.Font.Shading` 上而非段落上！
- **文字效果与 OOXML 补丁**：
  ```python
  f = rng.Font
  f.TextShadow.Type = 39
  f.TextShadow.Blur = 4.7249608    # 不显式设 Blur & Transparency，重开会退化为 -2
  f.TextShadow.Transparency = 0.68
  f.Reflection.Type = 5
  f.Line.Visible = -1; f.Line.Style = 1; f.Line.DashStyle = 8; f.Line.ForeColor.RGB = 255; f.Line.Weight = 1.0
  f.Glow.Radius = 18; f.Glow.Color.ObjectThemeColor = 7; f.Glow.Transparency = 0.6
  ```
  > 若题干要求渐变填充（`w14:textFill`）或主题发光饱和度（`w14:satMod="175000"`），先用 COM 存盘，再调用 `ooxml_patch.patch_word_grad_fill()` / `patch_word_glow()`，最后用 `Kwps` 重新打开 `Save()`。
- **查找替换防静默失效**：避开 `Find.Execute(Replace=2)` 对格式设置偶尔失效的坑，改用显式 `while` 循环逐处定位赋值 `rf.Text` 与 `rf.Font`（注意设防死循环上限或确保替换词不含原词）。
- **拼音指南（倒序防偏移）**：
  ```python
  for i in range(len(pinyins) - 1, -1, -1):
      c_rng = doc.Range(start_pos + i, start_pos + i + 1)
      c_rng.PhoneticGuide(pinyins[i], 3, 0, 14, "Arial")  # 3=左对齐, 0=偏移量, 14pt, Arial
  ```

---

## 5. 表格、合并顺序、网格表与内嵌图表

- **文本转表格**：`rng.ConvertToTable(Separator=1, NumColumns=..., NumRows=...)`（`1` = 制表符 `wdSeparateByTabs`，`0` = 段落）。
- **合并铁律**：
  1. **先设列宽与行高，再调用 `Merge()`**（列宽设 `tbl.AllowAutoFit = False; tbl.Columns(c).SetWidth(w, 0)`）。
  2. 一旦纵向合并，`tbl.Rows(r)` 会直接抛 `-2146822297`，之后一律改用 `tbl.Cell(r, c)`；读字体属性用 `tbl.Cell(r, c).Range.Characters(1).Font` 避开单元格结束符 `\r\x07`。
  3. 修改局部粗线（如 3 磅 `LineWidth=24`）后，全表 `tbl.Borders(-5)` 会返回混合值 `9999999`，自检须逐格断言 `Cell(r, c).Borders(...)`。
  4. **单双斜线表头**：双斜线需通过 `ooxml_patch.patch_word_double_diagonal_header()` 写入 `wpsCustomData:diagonals`，再 `Kwps` 重存。
- **网格表 1 与主题底纹**：`Kwps` 内置名通常为 `"网格型"`（MS Word 为 `"网格表 4"`），配合 `ApplyStyleHeadingRows = True; ApplyStyleRowBands = True` 并逐格设 `Cell(r, c).Shading.BackgroundPatternColor`。若题干指定主题色名称（如“橙色，着色 2”），需补 `patch_word_table_theme_fill(..., theme_fill="accent2")` 并 `Kwps` 重存。
- **Word 插入图表锚点与 4 行扩容**：
  1. 插入前必先 `rng_ins.Select()`，否则常锚回第 1 页；插入后如要设为浮动环绕，调用 `sh = ish.ConvertToShape()` 并设 `WrapFormat.Type`（`0`=四周型，`3`=浮于上方）。
  2. `AddChart2` 默认仅 4 行数据区，`SetSourceData` 无法自动扩行，必须在 `ChartData.Activate()` 后循环 `ws.Rows(5).Insert()` 扩行、填数、删多余列/系列、刷新并 `wb.Close(True)`。
  3. **只读检查阶段严禁调用 `ChartData.Activate()`**（无头只读下会死锁内嵌表格进程），直接用 `ooxml_patch.audit_word_xml()` 核验 `word/charts/chart*.xml`！

---

## 6. `.wps` 旧格式与文件导出

- 若题目要求保存为 `.wps` 二进制格式，必须显式传 `FileFormat=0`：`doc.SaveAs2(wps_path, 0)`（省略格式参数默认按 `.docx` 格式 `12` 写入）。
- 导出 PDF：`doc.ExportAsFixedFormat(OutputFileName=pdf_path, ExportFormat=17)`。

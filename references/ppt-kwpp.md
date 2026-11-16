# WPS 演示 (Kwpp) 考点与 COM/OOXML 防踩坑字典

> 所有写入操作使用 `wps_session("ppt", path, with_window=..., backup=True)`；涉及 `SmartArt.AllNodes.Delete()` 或视图对象时，**必须传 `with_window=True`**。

---

## 1. 幻灯片大小、宽屏 16:9 “确保适合”与母版背景

- **宽屏 16:9 尺寸坑**：在 `Kwpp` 中仅设 `SlideSize = 15` 会缩成 `720×405pt`！现代 16:9（`12192000×6858000` EMU）必须写：
  ```python
  prs.PageSetup.SlideSize = 7       # 7 = ppSlideSizeCustom / 16:9 960x540
  prs.PageSetup.SlideWidth = 960.0
  prs.PageSetup.SlideHeight = 540.0
  ```
- **“确保适合”必须联动缩放母版/版式并居中平移内容**：
  若题干是从 4:3（`720×540`）转 16:9（`960×540`）且要求“确保适合”，仅改 `PageSetup` 宽高会导致母版占位符比例和内容居中失分：
  ```python
  ratio = 960.0 / 720.0
  for sp in prs.SlideMaster.Shapes:
      sp.Left *= ratio
      sp.Width *= ratio
  for cl in prs.SlideMaster.CustomLayouts:
      for sp in cl.Shapes:
          sp.Left *= ratio
          sp.Width *= ratio
  # 非首页内容页按页面居中增量平移 +120pt
  for si in range(2, prs.Slides.Count + 1):
      for sp in prs.Slides(si).Shapes:
          sp.Left += (960.0 - 720.0) / 2.0
  ```
- **母版背景 vs 第 1 页**：首页常默认 `FollowMasterBackground = 0`，只改 `SlideMaster.Background.Fill.UserPicture(...)` 会漏掉首页！必须补循环：
  ```python
  prs.SlideMaster.Background.Fill.UserPicture(bg_img)
  for sl in prs.Slides:
      sl.FollowMasterBackground = -1
  ```

---

## 2. 贴边、几何分布与文本框锚定（`anchorCtr="0"` 大坑）

- **严禁在 `WithWindow=False` 下使用 `ShapeRange.Align` / `Distribute`**：无头模式下会静默失效或把文本框负间距叠在一起！一律用几何公式直接赋值 `Left` / `Top`：
  ```python
  W, H = prs.PageSetup.SlideWidth, prs.PageSetup.SlideHeight
  # 靠右靠顶
  pic.Left = W - pic.Width
  pic.Top = 0
  # 画布水平垂直居中
  sp.Left = (W - sp.Width) / 2.0
  sp.Top = (H - sp.Height) / 2.0
  # 顶端对齐 + 横向均匀分布
  sps = sorted([s2.Shapes(1), s2.Shapes(2), s2.Shapes(3)], key=lambda x: x.Left)
  top_min = min(x.Top for x in sps)
  x_min, x_max = sps[0].Left, sps[-1].Left
  for idx, x in enumerate(sps):
      x.Top = top_min
      x.Left = x_min + idx * (x_max - x_min) / (len(sps) - 1)
  ```
- **顶端对齐必须同时设 `TextFrame2.VerticalAnchor` 与 `HorizontalAnchor`**：
  单设 `TextFrame.VerticalAnchor = 1` 在底层 XML 只写出 `anchor="t"` 却缺失 `anchorCtr="0"`！必须写：
  ```python
  tf2 = shape.TextFrame2
  tf2.VerticalAnchor = 1    # 1=顶端, 3=居中, 4=底端
  tf2.HorizontalAnchor = 1  # 必须设 1 (msoAnchorNone) 以写出 anchorCtr="0"
  shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2  # 1=左, 2=居中, 3=右
  ```

---

## 3. 艺术字图案填充、阴影与小型大写字母

```python
f2 = shape.TextFrame2.TextRange.Font
f2.Smallcaps = -1          # 小型大写字母
f2.Fill.Patterned(16)      # 16=深色上对角线 (dkUpDiag)
f2.Fill.ForeColor.ObjectThemeColor = 14
f2.Fill.ForeColor.Brightness = -0.5
f2.Fill.BackColor.ObjectThemeColor = 13
f2.Fill.BackColor.Brightness = 0.25

sh = f2.Shadow
sh.Type = 21               # 预设外阴影
sh.Blur = 3.0
sh.OffsetX = 1.06066
sh.OffsetY = 1.06066
sh.Transparency = 0.6
sh.ForeColor.ObjectThemeColor = 1
sh.ForeColor.Brightness = -0.5

shape.TextFrame2.ThreeD.SetPresetCamera(19)
```

---

## 4. 智能图形 (SmartArt) 与表格合并对齐

- **必须 `WithWindow=True` 打开**：`WithWindow=False` 时调用 `SmartArt.AllNodes.Delete()` 会直接抛 `-2147417851 服务器出现意外情况` 崩溃。
- **勿盲目调用 `AllNodes.Add()`**：优先使用模板自带节点、`Delete()` 删多余节点、`Promote()` 升降级并回写 `TextFrame2.TextRange.Text`；先保存原形状引用并在全部读写完成后统一 `Delete()`。
- **PPT 表格合并后隐藏单元格 `hMerge="1"` 丢失对齐坑**：
  `Kwpp` 在 `tbl.Cell(r, 1).Merge(tbl.Cell(r, C))` 时会剥离隐藏 `hMerge="1"` 单元格里的 `<a:pPr algn="ctr"/>`！
  1. 合并前清空待合并冗余格文字并先全部设居中；
  2. 合并后再设一次；
  3. 存盘后调用 `ooxml_patch.patch_ppt_hmerge_align(pptx_path, slide_number, "ctr")`，再用 `Kwpp` 重新打开 `Save()`。

---

## 5. PPT 表格转图表与动作按钮

- **4 行模板扩容**：`AddChart2(-1, 51, left, top, w, h)` 默认只有 4 行数据区；超过 4 行（含表头即 `len(data) > 5`）必须先循环 `ws.Rows(5).Insert()` 扩行，填完数删多余列与系列，`ch.Refresh()` 并 `wb.Close(True)`（严禁跨进程向 `SetSourceData` 传 Excel `Range` 对象，必须传字符串区域地址如 `"='Sheet1'!$A$1:$C$5"`）。
- **动作按钮弹窗锁死规避**：严禁直接调用 `AddShape(138, ...)`（会强制弹出“操作设置”阻塞脚本），应先 `AddShape(1, ...)` 创建普通矩形，再改 `sh.AutoShapeType = 138`（或其他动作按钮枚举）并配置 `ActionSettings(1).Hyperlink` / `Action`。

---

## 6. 打印设置（`PrintOptions` 不写盘坑）、批注与二维码

- **`Kwpp` 的 `prs.PrintOptions` 在 `prs.Save()` 时不会持久化 `<p:prnPr>`！**
  必须先在 COM 中设好（或直接存盘），再调用 `ooxml_patch.patch_ppt_print_options(pptx_path, prn_what="handouts6", clr_mode="bw", hidden_slides=True)` 注入 `ppt/presProps.xml`，最后用 `Kwpp` 重新打开 `Save()`。
- **批注方法兼容**：部分版本支持 `sl.Comments.Add(...)`，部分仅支持 `Add2(...)`，务必 `try Add except Add2`。
- **隐藏幻灯片**：`prs.Slides(k).SlideShowTransition.Hidden = -1`（对应 XML `show="0"`）。
- **切换效果与排练计时**：如页面卷曲 `EntryEffect = 3948`、`AdvanceOnTime = -1`、`AdvanceTime = 5.0`。

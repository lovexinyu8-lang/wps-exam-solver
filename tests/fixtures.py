"""Programmatic OOXML & binary fixtures for offline regression tests."""

from __future__ import annotations

import io
import zipfile
from pathlib import Path
import docx
from docx.oxml import parse_xml


NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def build_word_fixture(path: Path, *, variant: str = "pass") -> Path:
    """Create a rich synthetic .docx covering styles, numbering, header, table, chart, and textbox."""
    doc = docx.Document()
    p_title = doc.add_paragraph("主标题文本")
    p_title.style = doc.styles["Heading 1"]
    p_body = doc.add_paragraph("正文内容段落")
    p_body.style = doc.styles["Normal"]

    # Add textbox paragraph (inside v:textbox -> w:txbxContent)
    txbx_p = parse_xml(
        f'<w:p xmlns:w="{NS_W}" xmlns:v="urn:schemas-microsoft-com:vml">'
        '<w:r><w:pict><v:shape><v:textbox><w:txbxContent>'
        '<w:p><w:r><w:t>文本框渐变目标</w:t></w:r></w:p>'
        '</w:txbxContent></v:textbox></v:shape></w:pict></w:r></w:p>'
    )
    doc.element.body.insert(2, txbx_p)

    # Add table
    tbl = doc.add_table(rows=2, cols=2)
    tbl.cell(0, 0).text = "表头发光目标"
    tbl.cell(0, 1).text = "数值"
    tbl.cell(1, 0).text = "项目A"
    tbl.cell(1, 1).text = "100"

    buf = io.BytesIO()
    doc.save(buf)

    # Now enrich the OOXML archive with realistic section/header/styles/numbering/chart XML
    out_buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(buf.getvalue()), "r") as zin, zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
        existing = {item.filename: zin.read(item.filename) for item in zin.infolist()}

        # Ensure styles.xml has explicit Heading1 + Normal properties
        styles_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{NS_W}">
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:pPr>
      <w:spacing w:beforeLines="50" w:afterLines="50" w:line="360" w:lineRule="auto"/>
      <w:jc w:val="left"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="宋体" w:eastAsia="宋体" w:hAnsi="宋体"/>
      <w:sz w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr>
      <w:pBdr><w:bottom w:val="single" w:sz="12" w:space="1" w:color="auto"/></w:pBdr>
      <w:spacing w:before="120" w:after="120" w:beforeLines="50" w:afterLines="50" w:line="240" w:lineRule="auto"/>
      <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="黑体" w:eastAsia="黑体" w:hAnsi="黑体"/>
      <w:b/>
      <w:sz w:val="32"/>
    </w:rPr>
  </w:style>
</w:styles>'''
        existing["word/styles.xml"] = styles_xml.encode("utf-8")

        # Numbering: ilvl=1 omits lvlRestart when variant=="fail_restart" to test lvlRestart=0 bug
        lvl1_restart = '<w:lvlRestart w:val="0"/>' if variant != "fail_restart" else ""
        numbering_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="{NS_W}">
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="chineseCounting"/>
      <w:lvlText w:val="%1、"/>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="chineseCounting"/>
      <w:lvlText w:val="（%2）"/>
      {lvl1_restart}
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>'''
        existing["word/numbering.xml"] = numbering_xml.encode("utf-8")

        # Header XML with TIME field, bottom border, and behindDoc="1" (or "0" in fail)
        behind_val = "0" if variant == "fail_header" else "1"
        washout_attr = ' gain="19661f" blacklevel="22938f"' if variant == "fail_header" else ""
        header_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="{NS_W}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:v="urn:schemas-microsoft-com:vml">
  <w:p>
    <w:pPr><w:pBdr><w:bottom w:val="double" w:sz="8" w:space="1" w:color="auto"/></w:pBdr></w:pPr>
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> TIME \\@ "yyyy年M月d日" </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>2026年11月19日</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    <w:r><w:drawing><wp:anchor behindDoc="{behind_val}"><wp:extent cx="914400" cy="914400"/></wp:anchor></w:drawing></w:r>
    <w:r><w:pict><v:shape><v:imagedata{washout_attr}/><v:textpath string="机密水印"/></v:shape></w:pict></w:r>
  </w:p>
</w:hdr>'''
        existing["word/header1.xml"] = header_xml.encode("utf-8")

        # Settings XML with gutterAtTop when variant == "gutter_top"
        existing["word/settings.xml"] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:settings xmlns:w="{NS_W}">'
            f'{"<w:gutterAtTop/>" if variant == "gutter_top" else ""}'
            f"</w:settings>"
        ).encode("utf-8")

        # Chart XML
        chart_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>销售趋势</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:tx><c:v>销售额</c:v></c:tx>
          <c:cat><c:strRef><c:strCache><c:ptCount val="5"/><c:pt idx="0"><c:v>1月</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:ptCount val="5"/><c:pt idx="0"><c:v>120</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
    <c:legend/>
  </c:chart>
</c:chartSpace>'''
        existing["word/charts/chart1.xml"] = chart_xml.encode("utf-8")
        if variant == "extra_chart":
            existing["word/charts/chart2.xml"] = chart_xml.encode("utf-8")

        # Inject document.xml relationships for header1 & charts + content types so python-docx retains them
        doc_rels = existing.get("word/_rels/document.xml.rels", b"").decode("utf-8")
        add_rels = ""
        if "rIdHdr1" not in doc_rels:
            add_rels += '<Relationship Id="rIdHdr1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>'
        if "rIdChart1" not in doc_rels:
            add_rels += '<Relationship Id="rIdChart1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/>'
        if variant == "extra_chart" and "rIdChart2" not in doc_rels:
            add_rels += '<Relationship Id="rIdChart2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart2.xml"/>'
        if add_rels:
            existing["word/_rels/document.xml.rels"] = doc_rels.replace("</Relationships>", add_rels + "</Relationships>").encode("utf-8")

        ctypes = existing.get("[Content_Types].xml", b"").decode("utf-8")
        add_ct = ""
        if "/word/charts/chart1.xml" not in ctypes:
            add_ct += '<Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
        if variant == "extra_chart" and "/word/charts/chart2.xml" not in ctypes:
            add_ct += '<Override PartName="/word/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>'
        if add_ct:
            existing["[Content_Types].xml"] = ctypes.replace("</Types>", add_ct + "</Types>").encode("utf-8")

        # Patch document.xml: margins 2.5cm (~1417 twips), headerReference, table shading/tblHeader/align, paragraph formatting, w14 effects
        doc_xml = existing["word/document.xml"].decode("utf-8")
        doc_xml = doc_xml.replace(
            '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"',
            '<w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417"',
        )
        # Ensure headerReference sits in sectPr
        doc_xml = doc_xml.replace(
            "<w:pgSz ",
            f'<w:headerReference w:type="default" r:id="rIdHdr1" xmlns:r="{NS_R}"/><w:pgSz ',
        )
        # Add tblHeader, table jc center, and cell shading on first table row
        doc_xml = doc_xml.replace(
            "<w:tblPr>",
            '<w:tblPr><w:jc w:val="center"/>',
            1,
        )
        doc_xml = doc_xml.replace(
            "<w:tr>",
            '<w:tr><w:trPr><w:tblHeader/></w:trPr>',
            1,
        )
        # Attach single clean pPr with firstLineChars="200" and beforeLines="50" on second paragraph ("正文内容段落")
        target_p_empty = "<w:p><w:pPr/><w:r><w:t>正文内容段落</w:t></w:r></w:p>"
        target_p_plain = "<w:p><w:r><w:t>正文内容段落</w:t></w:r></w:p>"
        if variant != "fail_para":
            new_p = (
                '<w:p><w:pPr><w:pStyle w:val="Normal"/>'
                '<w:spacing w:beforeLines="50" w:afterLines="50" w:line="360" w:lineRule="auto"/>'
                '<w:ind w:firstLineChars="200"/><w:jc w:val="both"/></w:pPr>'
                '<w:r><w:rPr><w:sz w:val="24"/></w:rPr><w:t>正文内容段落</w:t></w:r></w:p>'
            )
            if target_p_empty in doc_xml:
                doc_xml = doc_xml.replace(target_p_empty, new_p, 1)
            elif target_p_plain in doc_xml:
                doc_xml = doc_xml.replace(target_p_plain, new_p, 1)
        if variant == "forbidden_text":
            doc_xml = doc_xml.replace("正文内容段落", "正文内容段落_请删除占位符")
        existing["word/document.xml"] = doc_xml.encode("utf-8")

        for name, raw in existing.items():
            zout.writestr(name, raw)

    path.write_bytes(out_buf.getvalue())
    return path


def build_excel_reordered_fixture(path: Path, *, variant: str = "pass") -> Path:
    """Create an .xlsx where workbook sheet order ('Alpha', 'Beta') points to ('sheet2.xml', 'sheet1.xml') via rId!"""
    wb_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Alpha" sheetId="1" r:id="rId2"/>
    <sheet name="Beta" sheetId="2" r:id="rId1"/>
  </sheets>
  <pivotCaches>
    <pivotCache cacheId="10" r:id="rIdCache1"/>
  </pivotCaches>
</workbook>'''

    wb_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet2.xml"/>
  <Relationship Id="rIdCache1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>'''

    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="_ &quot;¥&quot;* #,##0.00_ ;_ &quot;¥&quot;* -#,##0.00_ ;_ &quot;-&quot;??_ ;_ @_ "/>
  </numFmts>
  <fonts count="1"><font><sz val="11"/><name val="宋体"/></font></fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="2">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
  <dxfs count="1">
    <dxf><fill><patternFill><bgColor rgb="FFFFC000"/></patternFill></fill></dxf>
  </dxfs>
</styleSheet>'''

    # sheet1.xml belongs to Beta (rId1) -> has sortState
    sheet1_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>金额</t></is></c></row>
    <row r="2"><c r="A2" s="1"><v>100</v></c></row>
    <row r="3"><c r="A3" s="1"><v>250</v></c></row>
  </sheetData>
  <sortState ref="A1:A3">
    <sortCondition ref="A2:A3"/>
  </sortState>
</worksheet>'''

    # sheet2.xml belongs to Alpha (rId2) -> has conditionalFormatting & formula & pivotTable link
    cf_formula = "WEEKDAY($B3,2)&gt;5" if variant == "fail_cf" else "WEEKDAY($B2,2)&gt;5"
    fml_str = "SUM(A2:A3)" if variant == "fail_formula" else "VLOOKUP(A2,Beta!$A$2:$A$3,1,FALSE)"
    val_xml = "" if variant == "no_cached_val" else f'<v>{"999" if variant == "fail_val" else "100"}</v>'
    sheet2_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1"><c r="A1" t="inlineStr"><is><t>序号</t></is></c><c r="B1" t="inlineStr"><is><t>结果</t></is></c></row>
    <row r="2"><c r="A2"><v>100</v></c><c r="B2"><f>{fml_str}</f>{val_xml}</c></row>
  </sheetData>
  <conditionalFormatting sqref="A2:H10">
    <cfRule type="expression" dxfId="0" priority="1"><formula>{cf_formula}</formula></cfRule>
  </conditionalFormatting>
</worksheet>'''

    sheet2_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdPt1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>'''

    pivot_cache_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" recordCount="1">
  <cacheFields count="2">
    <cacheField name="客户级别" numFmtId="0"/>
    <cacheField name="公司名称" numFmtId="0"/>
  </cacheFields>
</pivotCacheDefinition>'''

    pivot_table_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="数据透视表1" cacheId="10">
  <location ref="D3:E8"/>
  <rowFields count="1"><field x="0"/></rowFields>
  <dataFields count="1"><dataField name="计数项:公司名称" fld="1" subtotal="count"/></dataFields>
</pivotTableDefinition>'''

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>'''

    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>'''

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("xl/workbook.xml", wb_xml)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", styles_xml)
        z.writestr("xl/worksheets/sheet1.xml", sheet1_xml)
        z.writestr("xl/worksheets/sheet2.xml", sheet2_xml)
        z.writestr("xl/worksheets/_rels/sheet2.xml.rels", sheet2_rels)
        z.writestr("xl/pivotCache/pivotCacheDefinition1.xml", pivot_cache_xml)
        z.writestr("xl/pivotTables/pivotTable1.xml", pivot_table_xml)

    path.write_bytes(buf.getvalue())
    return path


def build_ppt_fixture(path: Path, *, variant: str = "pass") -> Path:
    """Create a rich synthetic .pptx covering 16:9, presProps, master bg, align/distribute, hMerge, SmartArt, and chart."""
    cx = 9144000 if variant == "fail_size" else 12192000  # 720pt vs 960pt
    cy = 6858000
    sz_type = "screen4x3" if variant == "fail_size" else "custom"

    pres_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldSz cx="{cx}" cy="{cy}" type="{sz_type}"/>
</p:presentation>'''

    prn_tag = '<p:prnPr prnWhat="handouts6" clrMode="bw" hiddenSlides="1"/>' if variant != "fail_print" else ""
    pres_props = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentationPr xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">{prn_tag}</p:presentationPr>'''

    master_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:slideMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdBg"/></a:blipFill></p:bgPr></p:bg>
    <p:spTree/>
  </p:cSld>
</p:slideMaster>'''

    # 1pt = 12700 EMU
    def _sp(name: str, text: str, left_pt: float, top_pt: float, w_pt: float, h_pt: float, *, anchor: str = "t", anchor_ctr: str = "0", algn: str = "ctr", smallcaps: bool = False, patt: str | None = None, shadow: bool = False, tx_box: bool = False) -> str:
        cap_attr = ' cap="small"' if smallcaps else ""
        patt_xml = f'<a:pattFill prst="{patt}"/>' if patt else ""
        eff_xml = "<a:effectLst><a:outerShdw/></a:effectLst>" if shadow else ""
        cnv_sp = '<p:cNvSpPr txBox="1"/>' if tx_box else "<p:cNvSpPr/>"
        return f'''<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="{name}"/>{cnv_sp}<p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="{int(left_pt * 12700)}" y="{int(top_pt * 12700)}"/><a:ext cx="{int(w_pt * 12700)}" cy="{int(h_pt * 12700)}"/></a:xfrm></p:spPr>
  <p:txBody>
    <a:bodyPr anchor="{anchor}" anchorCtr="{anchor_ctr}"/>
    <a:p><a:pPr algn="{algn}"/><a:r><a:rPr{cap_attr}>{patt_xml}{eff_xml}</a:rPr><a:t>{text}</a:t></a:r></a:p>
  </p:txBody>
</p:sp>'''

    top2 = 160.0 if variant == "fail_align" else 120.0
    left3 = 680.0 if variant == "fail_align" else 620.0  # evenly spaced: 100 (w=120, gap=140), 360 (w=120, gap=140), 620
    anchor_ctr_val = "1" if variant == "fail_anchor" else "0"
    slide1_bg = '<p:bg><p:bgPr><a:solidFill/></p:bgPr></p:bg>' if variant == "fail_master_bg" else ""
    hmerge_p = "<a:p><a:pPr/></a:p>" if variant == "fail_hmerge" else '<a:p><a:pPr algn="ctr"/></a:p>'

    slide1_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:slide xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">
  <p:cSld>
    {slide1_bg}
    <p:spTree>
      {_sp("TitleBox", "演示标题", 860.0, 0.0, 100.0, 40.0, anchor="t", anchor_ctr=anchor_ctr_val, algn="ctr", smallcaps=True, patt="dkUpDiag", shadow=True, tx_box=True)}
      {_sp("Card1", "卡片1", 100.0, 120.0, 120.0, 60.0)}
      {_sp("Card2", "卡片2", 360.0, top2, 120.0, 60.0)}
      {_sp("Card3", "卡片3", left3, 120.0, 120.0, 60.0)}
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="10" name="Table1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="635000" y="2540000"/><a:ext cx="3810000" cy="1270000"/></p:xfrm>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
          <a:tbl><a:tr h="370840"><a:tc><a:txBody><a:p><a:pPr algn="ctr"/><a:r><a:t>合并格</a:t></a:r></a:p></a:txBody></a:tc><a:tc hMerge="1"><a:txBody>{hmerge_p}</a:txBody></a:tc></a:tr></a:tbl>
        </a:graphicData></a:graphic>
      </p:graphicFrame>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="11" name="SmartArt1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="635000" y="3810000"/><a:ext cx="2540000" cy="1270000"/></p:xfrm>
        <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">
          <dgm:relIds r:dm="rIdDgm1"/>
        </a:graphicData></a:graphic>
      </p:graphicFrame>
    </p:spTree>
  </p:cSld>
  <p:transition advTm="5000"><p:prstTrans prst="pageCurlDouble"/></p:transition>
</p:slide>'''

    slide1_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdDgm1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>
</Relationships>'''

    diagram_data_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <dgm:ptLst>
    <dgm:pt modelId="0" type="doc"/>
    <dgm:pt modelId="1" type="node"><dgm:t><a:p><a:r><a:t>步骤一</a:t></a:r></a:p></dgm:t></dgm:pt>
    <dgm:pt modelId="2" type="node"><dgm:t><a:p><a:r><a:t>步骤二</a:t></a:r></a:p></dgm:t></dgm:pt>
  </dgm:ptLst>
</dgm:dataModel>'''

    chart_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>年度统计</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/><c:grouping val="clustered"/>
        <c:ser>
          <c:tx><c:v>人数</c:v></c:tx>
          <c:cat><c:strRef><c:strCache><c:ptCount val="5"/><c:pt idx="0"><c:v>A</c:v></c:pt></c:strCache></c:strRef></c:cat>
          <c:val><c:numRef><c:numCache><c:ptCount val="5"/><c:pt idx="0"><c:v>10</c:v></c:pt></c:numCache></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>'''

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("ppt/presentation.xml", pres_xml)
        z.writestr("ppt/presProps.xml", pres_props)
        z.writestr("ppt/slideMasters/slideMaster1.xml", master_xml)
        z.writestr("ppt/slides/slide1.xml", slide1_xml)
        z.writestr("ppt/slides/_rels/slide1.xml.rels", slide1_rels)
        z.writestr("ppt/diagrams/data1.xml", diagram_data_xml)
        z.writestr("ppt/charts/chart1.xml", chart_xml)
        if variant == "extra_chart":
            z.writestr("ppt/charts/chart2.xml", chart_xml)
    path.write_bytes(buf.getvalue())
    return path


def build_ole2_dummy(path: Path) -> Path:
    """Create a dummy OLE2 binary signature file for container detection tests."""
    path.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 504)
    return path

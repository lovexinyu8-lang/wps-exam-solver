"""OOXML surgical patchers (for WPS COM blind spots) and read-only XML auditors."""

from __future__ import annotations

import json
import os
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Callable

PT_PER_CM = 28.346456692913385

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "wpsCustomData": "http://www.wps.cn/officeDocument/2013/wpsCustomData",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
}


def cm_to_pt(cm: float) -> float:
    return cm * PT_PER_CM


def pt_to_cm(pt: float) -> float:
    return pt / PT_PER_CM


def rgb_to_bgr_int(r: int, g: int, b: int) -> int:
    """Convert 0-255 RGB channels to WPS/Office COM BGR integer."""
    return (int(b) << 16) | (int(g) << 8) | int(r)


def bgr_int_to_hex(val: int) -> str:
    """Convert WPS/Office COM BGR integer to uppercase RRGGBB hex string."""
    r = val & 0xFF
    g = (val >> 8) & 0xFF
    b = (val >> 16) & 0xFF
    return f"{r:02X}{g:02X}{b:02X}"


def _q(prefix: str, tag: str) -> str:
    return f"{{{NS[prefix]}}}{tag}"


def rewrite_zip(file_path: str, handlers: dict[str, Callable[[bytes], bytes]]) -> list[str]:
    """Atomically rewrite selected entries inside an OOXML ZIP archive."""
    abs_path = os.path.abspath(file_path)
    if not zipfile.is_zipfile(abs_path):
        raise ValueError(
            f"Not an OOXML ZIP container (possibly legacy binary .wps/.et/.doc): {abs_path}"
        )
    tmp_path = abs_path + ".ooxml.tmp"
    modified: list[str] = []
    with zipfile.ZipFile(abs_path, "r") as zin, zipfile.ZipFile(
        tmp_path, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            raw = zin.read(item.filename)
            if item.filename in handlers:
                new_raw = handlers[item.filename](raw)
                if new_raw != raw:
                    modified.append(item.filename)
                raw = new_raw
            zout.writestr(item, raw)
    shutil.move(tmp_path, abs_path)
    return modified


# ---------------------------------------------------------------------------
# Word OOXML Patchers (reopen & Save() with Kwps.Application afterward)
# ---------------------------------------------------------------------------

def _require_docx():
    try:
        import docx
        from docx.oxml import parse_xml
        return docx, parse_xml
    except ImportError as exc:
        raise RuntimeError("python-docx is required for Word OOXML surgical patches.") from exc


def patch_word_grad_fill(
    docx_path: str,
    target_text: str,
    stops: list[tuple[int, str, int | None, int | None]],
    angle_deg: float = 90.0,
) -> int:
    """Inject <w14:textFill><w14:gradFill> onto runs covering target_text.

    Args:
        stops: List of (pos_0_to_100000, scheme_clr, lum_mod_or_none, lum_off_or_none).
               e.g. [(0, 'accent2', 25000, 75000), (50000, 'accent2', None, None), (100000, 'accent2', 85000, None)]
    """
    docx, parse_xml = _require_docx()
    gs_parts = []
    for pos, clr, lum_mod, lum_off in stops:
        mods = ""
        if lum_mod is not None:
            mods += f'<w14:lumMod w14:val="{int(lum_mod)}"/>'
        if lum_off is not None:
            mods += f'<w14:lumOff w14:val="{int(lum_off)}"/>'
        inner = f'<w14:schemeClr w14:val="{clr}">{mods}</w14:schemeClr>' if mods else f'<w14:schemeClr w14:val="{clr}"/>'
        gs_parts.append(f'<w14:gs w14:pos="{int(pos)}">{inner}</w14:gs>')
    ang = int(round(angle_deg * 60000))
    grad_xml = (
        f'<w14:textFill xmlns:w14="{NS["w14"]}">'
        f'<w14:gradFill><w14:gsLst>{"".join(gs_parts)}</w14:gsLst>'
        f'<w14:lin w14:ang="{ang}" w14:scaled="1"/></w14:gradFill></w14:textFill>'
    )

    doc = docx.Document(docx_path)
    patched = 0
    for p in doc.paragraphs:
        full_text = "".join(r.text for r in p.runs)
        idx = full_text.find(target_text)
        if idx == -1:
            continue
        t_start, t_end = idx, idx + len(target_text)
        cur = 0
        for r in p.runs:
            s, e = cur, cur + len(r.text)
            cur = e
            if r.text and not (e <= t_start or s >= t_end):
                rPr = r._r.get_or_add_rPr()
                for old in rPr.findall(_q("w14", "textFill")):
                    rPr.remove(old)
                rPr.append(parse_xml(grad_xml))
                patched += 1
    doc.save(docx_path)
    return patched


def patch_word_glow(
    docx_path: str,
    target_text: str,
    *,
    radius_pt: float = 18.0,
    scheme_clr: str = "accent4",
    sat_mod: int = 175000,
    alpha: int = 60000,
) -> int:
    """Inject or normalize <w14:glow> (with w14:satMod and w14:alpha) on runs containing target_text."""
    docx, parse_xml = _require_docx()
    rad = int(round(radius_pt * 12700))
    glow_xml = (
        f'<w14:glow xmlns:w14="{NS["w14"]}" w14:rad="{rad}">'
        f'<w14:schemeClr w14:val="{scheme_clr}">'
        f'<w14:satMod w14:val="{int(sat_mod)}"/>'
        f'<w14:alpha w14:val="{int(alpha)}"/>'
        f"</w14:schemeClr></w14:glow>"
    )

    doc = docx.Document(docx_path)
    patched = 0
    for p in doc.paragraphs:
        full_text = "".join(r.text for r in p.runs)
        idx = full_text.find(target_text)
        if idx == -1:
            continue
        t_start, t_end = idx, idx + len(target_text)
        cur = 0
        for r in p.runs:
            s, e = cur, cur + len(r.text)
            cur = e
            if r.text and not (e <= t_start or s >= t_end):
                rPr = r._r.get_or_add_rPr()
                for old in rPr.findall(_q("w14", "glow")):
                    rPr.remove(old)
                rPr.append(parse_xml(glow_xml))
                patched += 1
    doc.save(docx_path)
    return patched


def patch_word_table_theme_fill(
    docx_path: str,
    *,
    table_index: int = 0,
    row_index: int = 0,
    theme_fill: str = "accent2",
    fill_hex: str | None = None,
) -> int:
    """Attach w:themeFill (which WPS COM Shading omits) onto <w:shd> of a table row."""
    docx, _ = _require_docx()
    doc = docx.Document(docx_path)
    tbl = doc.tables[table_index]
    row = tbl.rows[row_index]
    patched = 0
    seen_tcs = set()
    for cell in row.cells:
        tc = cell._tc
        if id(tc) in seen_tcs:
            continue
        seen_tcs.add(id(tc))
        tcPr = tc.get_or_add_tcPr()
        shd = tcPr.find(_q("w", "shd"))
        if shd is None:
            from docx.oxml import OxmlElement
            shd = OxmlElement("w:shd")
            shd.set(_q("w", "val"), "clear")
            shd.set(_q("w", "color"), "auto")
            tcPr.append(shd)
        shd.set(_q("w", "themeFill"), theme_fill)
        if fill_hex:
            shd.set(_q("w", "fill"), fill_hex.lstrip("#").upper())
        patched += 1
    doc.save(docx_path)
    return patched


def patch_word_double_diagonal_header(
    docx_path: str,
    *,
    table_index: int = 0,
    row_index: int = 0,
    col_index: int = 0,
    sz_eighths: int = 24,
    diagonals: tuple[tuple[int, int], ...] = ((10000, 35000), (10000, 25000)),
) -> None:
    """Inject WPS native double-diagonal header (wpsCustomData:diagonals) into a table cell."""
    docx, parse_xml = _require_docx()
    doc = docx.Document(docx_path)
    tc = doc.tables[table_index].rows[row_index].cells[col_index]._tc
    tcPr = tc.get_or_add_tcPr()
    diag_items = "".join(
        f'<wpsCustomData:diagonal from="{f}" to="{t}">'
        f'<wpsCustomData:border w:val="single" w:color="auto" w:sz="{int(sz_eighths)}" w:space="0"/>'
        f"</wpsCustomData:diagonal>"
        for f, t in diagonals
    )
    diag_xml = (
        f'<mc:AlternateContent xmlns:mc="{NS["mc"]}" '
        f'xmlns:wpsCustomData="{NS["wpsCustomData"]}" xmlns:w="{NS["w"]}">'
        f'<mc:Choice Requires="wpsCustomData"><wpsCustomData:diagonals>{diag_items}'
        f"</wpsCustomData:diagonals></mc:Choice></mc:AlternateContent>"
    )
    for old in list(tcPr.findall(_q("mc", "AlternateContent"))):
        if old.find(f".//{_q('wpsCustomData', 'diagonals')}") is not None:
            tcPr.remove(old)
    tcPr.append(parse_xml(diag_xml))

    para_diag = (
        f'<mc:AlternateContent xmlns:mc="{NS["mc"]}" '
        f'xmlns:wpsCustomData="{NS["wpsCustomData"]}">'
        f'<mc:Choice Requires="wpsCustomData"><wpsCustomData:diagonalParaType/>'
        f"</mc:Choice></mc:AlternateContent>"
    )
    for p_el in tc.findall(_q("w", "p"))[:2]:
        pPr = p_el.find(_q("w", "pPr"))
        if pPr is None:
            from docx.oxml import OxmlElement
            pPr = OxmlElement("w:pPr")
            p_el.insert(0, pPr)
        pPr.append(parse_xml(para_diag))
    doc.save(docx_path)


# ---------------------------------------------------------------------------
# PPT OOXML Patchers (reopen & Save() with Kwpp.Application afterward)
# ---------------------------------------------------------------------------

def patch_ppt_print_options(
    pptx_path: str,
    *,
    prn_what: str = "handouts6",
    clr_mode: str = "bw",
    hidden_slides: bool = True,
) -> list[str]:
    """Inject <p:prnPr> into ppt/presProps.xml because Kwpp drops PrintOptions on Save()."""
    prn_tag = (
        f'<p:prnPr prnWhat="{prn_what}" clrMode="{clr_mode}" '
        f'hiddenSlides="{"1" if hidden_slides else "0"}"/>'
    )

    def _patch(raw: bytes) -> bytes:
        s = raw.decode("utf-8")
        s = re.sub(r"<p:prnPr[^>]*/>", "", s)
        for anchor in ("<p:showPr", "<p:clrMru", "<p:extLst", "</p:presentationPr>"):
            if anchor in s:
                return s.replace(anchor, prn_tag + anchor, 1).encode("utf-8")
        return raw

    return rewrite_zip(pptx_path, {"ppt/presProps.xml": _patch})


def patch_ppt_hmerge_align(pptx_path: str, slide_number: int, algn: str = "ctr") -> list[str]:
    """Restore algn on hidden hMerge='1' cells whose <a:pPr> Kwpp strips on Merge()."""
    entry = f"ppt/slides/slide{int(slide_number)}.xml"

    def _fix_p(m: re.Match[str]) -> str:
        p = m.group(0)
        if p == "<a:p/>":
            return f'<a:p><a:pPr algn="{algn}"/></a:p>'
        if "<a:pPr" in p:
            pp = re.search(r"<a:pPr[^>]*?/?>", p).group(0)
            if "algn=" in pp:
                newpp = re.sub(r'algn="[^"]*"', f'algn="{algn}"', pp)
            else:
                newpp = pp.replace("<a:pPr", f'<a:pPr algn="{algn}"', 1)
            return p.replace(pp, newpp, 1)
        return p.replace("<a:p>", f'<a:p><a:pPr algn="{algn}"/>', 1)

    def _fix_tc(m: re.Match[str]) -> str:
        tc = m.group(0)
        head = re.match(r"<a:tc[^>]*>", tc)
        if not head or 'hMerge="1"' not in head.group(0):
            return tc
        return re.sub(r"<a:p>.*?</a:p>|<a:p/>", _fix_p, tc, flags=re.S)

    def _patch(raw: bytes) -> bytes:
        s = raw.decode("utf-8")
        return re.sub(r"<a:tc[ >].*?</a:tc>", _fix_tc, s, flags=re.S).encode("utf-8")

    return rewrite_zip(pptx_path, {entry: _patch})


# ---------------------------------------------------------------------------
# Read-only OOXML Auditors (zero COM lock / ChartData.Activate() risk)
# ---------------------------------------------------------------------------

def _parse_rels(z: zipfile.ZipFile, rels_path: str) -> dict[str, str]:
    if rels_path not in z.namelist():
        return {}
    root = ET.fromstring(z.read(rels_path))
    return {
        el.attrib.get("Id", ""): el.attrib.get("Target", "")
        for el in root
        if "Id" in el.attrib and "Target" in el.attrib
    }


def _audit_chart_xml(raw: bytes, name: str) -> dict[str, Any]:
    root = ET.fromstring(raw)
    plot_area = root.find(f".//{_q('c', 'plotArea')}")
    chart_types: list[str] = []
    if plot_area is not None:
        for child in plot_area:
            local = child.tag.split("}")[-1]
            if local.endswith("Chart"):
                chart_types.append(local)

    series_info: list[dict[str, Any]] = []
    for ser in root.findall(f".//{_q('c', 'ser')}"):
        title_el = ser.find(f"{_q('c', 'tx')}//{_q('c', 'v')}")
        cat_pts = [el.attrib.get("val") for el in ser.findall(f"{_q('c', 'cat')}//{_q('c', 'ptCount')}")]
        val_pts = [el.attrib.get("val") for el in ser.findall(f"{_q('c', 'val')}//{_q('c', 'ptCount')}")]
        cats = [el.text for el in ser.findall(f"{_q('c', 'cat')}//{_q('c', 'v')}")][:8]
        vals = [el.text for el in ser.findall(f"{_q('c', 'val')}//{_q('c', 'v')}")][:8]
        series_info.append({
            "name": title_el.text if title_el is not None else None,
            "cat_pt_count": int(cat_pts[0]) if cat_pts and cat_pts[0] is not None else None,
            "val_pt_count": int(val_pts[0]) if val_pts and val_pts[0] is not None else None,
            "sample_cats": cats,
            "sample_vals": vals,
        })

    title_texts = [
        el.text for el in root.findall(f".//{_q('c', 'chart')}/{_q('c', 'title')}//{_q('a', 't')}") if el.text
    ]
    return {
        "file": name,
        "chart_types": chart_types,
        "series_count": len(series_info),
        "series": series_info,
        "title": "".join(title_texts) if title_texts else None,
        "has_legend": root.find(f".//{_q('c', 'legend')}") is not None,
    }


def audit_word_xml(docx_path: str) -> dict[str, Any]:
    """Read-only OOXML audit for Word (.docx / zip-based .wps)."""
    with zipfile.ZipFile(docx_path, "r") as z:
        names = set(z.namelist())
        doc_root = ET.fromstring(z.read("word/document.xml"))
        rels = _parse_rels(z, "word/_rels/document.xml.rels")

        sections: list[dict[str, Any]] = []
        for idx, sect in enumerate(doc_root.findall(f".//{_q('w', 'sectPr')}"), 1):
            pgSz = sect.find(_q("w", "pgSz"))
            pgMar = sect.find(_q("w", "pgMar"))
            typ = sect.find(_q("w", "type"))
            hdr_refs = []
            for hr in sect.findall(_q("w", "headerReference")):
                rid = hr.attrib.get(_q("r", "id"), "")
                htype = hr.attrib.get(_q("w", "type"), "default")
                target = rels.get(rid, "")
                hdr_entry = f"word/{target}" if target and not target.startswith("word/") else target
                hdr_info: dict[str, Any] = {"type": htype, "target": hdr_entry}
                if hdr_entry and hdr_entry in names:
                    hraw = z.read(hdr_entry).decode("utf-8", errors="replace")
                    hdr_info["drawings"] = hraw.count("<w:drawing")
                    hdr_info["picts"] = hraw.count("<w:pict")
                    hdr_info["behind_doc"] = re.findall(r'behindDoc="(\d)"', hraw)
                    hdr_info["instr_text"] = re.findall(r"<w:instrText[^>]*>(.*?)</w:instrText>", hraw)
                hdr_refs.append(hdr_info)
            sections.append({
                "section": idx,
                "start_type": typ.attrib.get(_q("w", "val")) if typ is not None else "nextPage",
                "pgSz": {k.split("}")[-1]: v for k, v in (pgSz.attrib.items() if pgSz is not None else [])},
                "pgMar": {k.split("}")[-1]: v for k, v in (pgMar.attrib.items() if pgMar is not None else [])},
                "headers": hdr_refs,
            })

        numbering_levels: list[dict[str, Any]] = []
        if "word/numbering.xml" in names:
            num_root = ET.fromstring(z.read("word/numbering.xml"))
            for abs_el in num_root.findall(_q("w", "abstractNum")):
                aid = abs_el.attrib.get(_q("w", "abstractNumId"))
                lvls = []
                for lv in abs_el.findall(_q("w", "lvl")):
                    nf = lv.find(_q("w", "numFmt"))
                    lt = lv.find(_q("w", "lvlText"))
                    lr = lv.find(_q("w", "lvlRestart"))
                    lvls.append({
                        "ilvl": int(lv.attrib.get(_q("w", "ilvl"), "0")),
                        "numFmt": nf.attrib.get(_q("w", "val")) if nf is not None else None,
                        "lvlText": lt.attrib.get(_q("w", "val")) if lt is not None else None,
                        "lvlRestart": lr.attrib.get(_q("w", "val")) if lr is not None else None,
                    })
                numbering_levels.append({"abstractNumId": aid, "levels": lvls})

        doc_raw = z.read("word/document.xml").decode("utf-8", errors="replace")
        theme_fills = re.findall(r'w:themeFill="([^"]+)"', doc_raw)
        glow_sat_mods = re.findall(r'w14:satMod\s+w14:val="(\d+)"', doc_raw)
        charts = [
            _audit_chart_xml(z.read(n), n)
            for n in sorted(names)
            if n.startswith("word/charts/chart") and n.endswith(".xml")
        ]

        return {
            "kind": "word",
            "sections_count": len(sections),
            "sections": sections,
            "body_drawings": doc_raw.count("<w:drawing"),
            "has_w14_textFill": "<w14:textFill" in doc_raw,
            "has_w14_glow": "<w14:glow" in doc_raw,
            "glow_sat_mods": glow_sat_mods,
            "has_w14_shadow": "<w14:shadow" in doc_raw,
            "has_w14_reflection": "<w14:reflection" in doc_raw,
            "theme_fills": theme_fills,
            "has_wps_diagonals": "wpsCustomData:diagonals" in doc_raw,
            "numbering_abstracts": numbering_levels,
            "charts": charts,
        }


def audit_ppt_xml(pptx_path: str) -> dict[str, Any]:
    """Read-only OOXML audit for Presentation (.pptx)."""
    with zipfile.ZipFile(pptx_path, "r") as z:
        names = set(z.namelist())
        pres_root = ET.fromstring(z.read("ppt/presentation.xml"))
        sld_sz = pres_root.find(_q("p", "sldSz"))
        sz_info = {k: v for k, v in (sld_sz.attrib.items() if sld_sz is not None else [])}

        prn_info: dict[str, str] = {}
        if "ppt/presProps.xml" in names:
            pp_root = ET.fromstring(z.read("ppt/presProps.xml"))
            prn = pp_root.find(_q("p", "prnPr"))
            if prn is not None:
                prn_info = dict(prn.attrib)

        slide_files = sorted(
            (n for n in names if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
            key=lambda s: int(re.search(r"(\d+)", s).group(1)),
        )
        slides_info: list[dict[str, Any]] = []
        for sf in slide_files:
            s_root = ET.fromstring(z.read(sf))
            body_prs = [
                {
                    "anchor": bp.attrib.get("anchor"),
                    "anchorCtr": bp.attrib.get("anchorCtr"),
                }
                for bp in s_root.findall(f".//{_q('a', 'bodyPr')}")
            ]
            shapes_body_pr = []
            for sp_el in s_root.findall(f".//{_q('p', 'spTree')}/{_q('p', 'sp')}"):
                cnv = sp_el.find(f"{_q('p', 'nvSpPr')}/{_q('p', 'cNvPr')}")
                bp = sp_el.find(f"{_q('p', 'txBody')}/{_q('a', 'bodyPr')}")
                texts = [t.text for t in sp_el.findall(f"{_q('p', 'txBody')}//{_q('a', 't')}") if t.text]
                shapes_body_pr.append({
                    "name": cnv.attrib.get("name") if cnv is not None else None,
                    "text": "".join(texts),
                    "anchor": bp.attrib.get("anchor") if bp is not None else None,
                    "anchorCtr": bp.attrib.get("anchorCtr") if bp is not None else None,
                })
            hmerge_missing_ctr = 0
            for tc in s_root.findall(f".//{_q('a', 'tc')}"):
                if tc.attrib.get("hMerge") == "1":
                    for p_el in tc.findall(_q("a", "p")):
                        pPr = p_el.find(_q("a", "pPr"))
                        if pPr is None or "algn" not in pPr.attrib:
                            hmerge_missing_ctr += 1
            slides_info.append({
                "file": sf,
                "hidden": s_root.attrib.get("show") == "0",
                "has_transition": s_root.find(f".//{_q('p', 'transition')}") is not None,
                "has_timing": s_root.find(f".//{_q('p', 'timing')}") is not None,
                "body_prs": body_prs,
                "shapes_body_pr": shapes_body_pr,
                "hmerge_missing_algn": hmerge_missing_ctr,
            })

        charts = [
            _audit_chart_xml(z.read(n), n)
            for n in sorted(names)
            if n.startswith("ppt/charts/chart") and n.endswith(".xml")
        ]
        return {
            "kind": "ppt",
            "sldSz": sz_info,
            "prnPr": prn_info,
            "slides_count": len(slides_info),
            "slides": slides_info,
            "charts": charts,
        }


def audit_excel_xml(xlsx_path: str) -> dict[str, Any]:
    """Read-only OOXML audit for Spreadsheet (.xlsx)."""
    with zipfile.ZipFile(xlsx_path, "r") as z:
        names = set(z.namelist())
        wb_root = ET.fromstring(z.read("xl/workbook.xml"))
        sheets_meta = [
            {
                "name": s.attrib.get("name"),
                "sheetId": s.attrib.get("sheetId"),
                "state": s.attrib.get("state", "visible"),
            }
            for s in wb_root.findall(f".//{_q('x', 'sheet')}")
        ]

        sheet_files = sorted(
            (n for n in names if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)),
            key=lambda s: int(re.search(r"(\d+)", s).group(1)),
        )
        sheets_xml: list[dict[str, Any]] = []
        for idx, sf in enumerate(sheet_files):
            s_root = ET.fromstring(z.read(sf))
            cfs = []
            for cf in s_root.findall(_q("x", "conditionalFormatting")):
                rules = []
                for r_el in cf.findall(_q("x", "cfRule")):
                    fmls = [f.text for f in r_el.findall(_q("x", "formula"))]
                    rules.append({
                        "type": r_el.attrib.get("type"),
                        "operator": r_el.attrib.get("operator"),
                        "dxfId": r_el.attrib.get("dxfId"),
                        "formulas": fmls,
                    })
                cfs.append({"sqref": cf.attrib.get("sqref"), "rules": rules})

            ss_el = s_root.find(_q("x", "sortState"))
            sort_state = None
            if ss_el is not None:
                conds = [dict(c.attrib) for c in ss_el.findall(_q("x", "sortCondition"))]
                sort_state = {"ref": ss_el.attrib.get("ref"), "conditions": conds}

            af_el = s_root.find(_q("x", "autoFilter"))
            sheets_xml.append({
                "file": sf,
                "name": sheets_meta[idx]["name"] if idx < len(sheets_meta) else None,
                "conditional_formatting": cfs,
                "sort_state": sort_state,
                "auto_filter": af_el.attrib.get("ref") if af_el is not None else None,
            })

        pivots = []
        for pn in sorted(n for n in names if n.startswith("xl/pivotTables/pivotTable") and n.endswith(".xml")):
            pr = ET.fromstring(z.read(pn))
            loc = pr.find(_q("x", "location"))
            pivots.append({
                "file": pn,
                "name": pr.attrib.get("name"),
                "cacheId": pr.attrib.get("cacheId"),
                "location": loc.attrib.get("ref") if loc is not None else None,
            })

        charts = [
            _audit_chart_xml(z.read(n), n)
            for n in sorted(names)
            if n.startswith("xl/charts/chart") and n.endswith(".xml")
        ]
        return {
            "kind": "excel",
            "sheets": sheets_meta,
            "worksheets": sheets_xml,
            "pivot_tables": pivots,
            "charts": charts,
        }


def audit_ooxml(file_path: str) -> dict[str, Any]:
    """Dispatch read-only OOXML audit by container & internal root files."""
    abs_path = os.path.abspath(file_path)
    if not zipfile.is_zipfile(abs_path):
        return {"container": "ole2_or_binary", "path": abs_path, "ooxml_supported": False}
    with zipfile.ZipFile(abs_path, "r") as z:
        names = set(z.namelist())
    if "word/document.xml" in names:
        return audit_word_xml(abs_path)
    if "ppt/presentation.xml" in names:
        return audit_ppt_xml(abs_path)
    if "xl/workbook.xml" in names:
        return audit_excel_xml(abs_path)
    return {"container": "zip_unknown", "path": abs_path, "ooxml_supported": False}


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser(description="Read-only OOXML audit for WPS exam files.")
    parser.add_argument("file", help="Path to .docx / .pptx / .xlsx file.")
    args = parser.parse_args()
    print(json.dumps(audit_ooxml(args.file), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

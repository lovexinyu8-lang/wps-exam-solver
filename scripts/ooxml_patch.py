"""OOXML surgical patchers (for WPS COM blind spots) and read-only XML auditors."""

from __future__ import annotations

import json
import os
import posixpath
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any, Callable

try:
    import openpyxl
except ImportError:  # pragma: no cover
    openpyxl = None  # type: ignore

PT_PER_CM = 28.346456692913385
EMU_PER_PT = 12700.0

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "wpsCustomData": "http://www.wps.cn/officeDocument/2013/wpsCustomData",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
}

_PPT_PATT_PRST_MAP: dict[str, int] = {
    "pct5": 1,
    "pct10": 2,
    "pct20": 3,
    "pct25": 4,
    "pct30": 5,
    "pct40": 6,
    "pct50": 7,
    "pct60": 8,
    "pct70": 9,
    "pct75": 10,
    "pct80": 11,
    "pct90": 12,
    "dnDiag": 13,
    "upDiag": 14,
    "dkDnDiag": 15,
    "dkUpDiag": 16,
    "wdDnDiag": 17,
    "wdUpDiag": 18,
    "ltDnDiag": 19,
    "ltUpDiag": 20,
    "dashDnDiag": 21,
    "dashUpDiag": 22,
}

_PPT_TRANSITION_EFFECT_MAP: dict[str, int] = {
    "cut": 257,
    "cover": 1284,
    "dissolve": 1537,
    "fade": 1793,
    "wipe": 3846,
    "split": 3847,
    "push": 3852,
    "zoom": 3856,
    "pageCurl": 3948,
    "pageCurlDouble": 3948,
    "pageCurlSingle": 3947,
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


def hex_to_bgr_int(hex_rgb: str | None) -> int | None:
    if not hex_rgb:
        return None
    clean = hex_rgb.strip().lstrip("#")
    if len(clean) == 8:  # ARGB -> RRGGBB
        clean = clean[2:]
    if len(clean) != 6 or not re.fullmatch(r"[0-9A-Fa-f]{6}", clean):
        return None
    r = int(clean[0:2], 16)
    g = int(clean[2:4], 16)
    b = int(clean[4:6], 16)
    return rgb_to_bgr_int(r, g, b)


def _q(prefix: str, tag: str) -> str:
    return f"{{{NS[prefix]}}}{tag}"


def _resolve_package_target(base_dir: str, target: str) -> str:
    t = target.replace("\\", "/")
    if t.startswith("/"):
        return posixpath.normpath(t.lstrip("/"))
    base = base_dir.strip("/")
    joined = f"{base}/{t}" if base else t
    return posixpath.normpath(joined)


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
        from docx.text.paragraph import Paragraph
        return docx, parse_xml, Paragraph
    except ImportError as exc:
        raise RuntimeError("python-docx is required for Word OOXML surgical patches.") from exc


def _iter_word_body_paragraphs(doc: Any, ParagraphCls: Any) -> list[Any]:
    """Collect unique Paragraph wrappers across body, tables (incl. nested), and textboxes."""
    w_p_tag = _q("w", "p")
    # findall returns XML elements in document order without duplicate proxy id collisions.
    p_elements = doc.element.body.findall(f".//{w_p_tag}")
    return [ParagraphCls(p_el, doc) for p_el in p_elements]


def patch_word_grad_fill(
    docx_path: str,
    target_text: str,
    stops: list[tuple[int, str, int | None, int | None]],
    angle_deg: float = 90.0,
) -> int:
    """Inject <w14:textFill><w14:gradFill> onto runs covering target_text across body/tables/textboxes.

    Args:
        stops: List of (pos_0_to_100000, scheme_clr, lum_mod_or_none, lum_off_or_none).
               e.g. [(0, 'accent2', 25000, 75000), (50000, 'accent2', None, None), (100000, 'accent2', 85000, None)]
    """
    docx, parse_xml, ParagraphCls = _require_docx()
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
    for p in _iter_word_body_paragraphs(doc, ParagraphCls):
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
    """Inject or normalize <w14:glow> (with w14:satMod and w14:alpha) across body/tables/textboxes."""
    docx, parse_xml, ParagraphCls = _require_docx()
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
    for p in _iter_word_body_paragraphs(doc, ParagraphCls):
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
    docx, _, _ = _require_docx()
    doc = docx.Document(docx_path)
    tbl = doc.tables[table_index]
    row = tbl.rows[row_index]
    patched = 0
    seen_tcs: list[Any] = []
    for cell in row.cells:
        tc = cell._tc
        if tc in seen_tcs:
            continue
        seen_tcs.append(tc)
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
    docx, parse_xml, _ = _require_docx()
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
        for old_p in list(pPr.findall(_q("mc", "AlternateContent"))):
            if old_p.find(f".//{_q('wpsCustomData', 'diagonalParaType')}") is not None:
                pPr.remove(old_p)
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


def _infer_xl_chart_type(plot_area: ET.Element | None) -> int | None:
    if plot_area is None:
        return None
    for child in plot_area:
        local = child.tag.split("}")[-1]
        if local == "barChart":
            bd = child.find(_q("c", "barDir"))
            gp = child.find(_q("c", "grouping"))
            bdir = bd.attrib.get("val", "col") if bd is not None else "col"
            grp = gp.attrib.get("val", "clustered") if gp is not None else "clustered"
            if bdir == "bar":
                return 58 if grp == "stacked" else 57
            return 52 if grp == "stacked" else 51
        if local == "lineChart":
            return 4
        if local == "pieChart":
            return 5
        if local == "areaChart":
            return 1
        if local == "scatterChart":
            return -4169
        if local == "doughnutChart":
            return -4120
    return None


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
    has_gridlines = root.find(f".//{_q('c', 'majorGridlines')}") is not None
    return {
        "file": name,
        "chart_types": chart_types,
        "chart_type_code": _infer_xl_chart_type(plot_area),
        "series_count": len(series_info),
        "series": series_info,
        "title": "".join(title_texts) if title_texts else None,
        "has_legend": root.find(f".//{_q('c', 'legend')}") is not None,
        "has_gridlines": has_gridlines,
    }


def _word_align_to_int(val: str | None) -> int | None:
    if val is None:
        return None
    return {"left": 0, "start": 0, "center": 1, "right": 2, "end": 2, "both": 3, "distribute": 4}.get(val)


def _word_line_rule_and_val(spacing_el: ET.Element | None) -> tuple[int | None, float | None]:
    if spacing_el is None:
        return None, None
    rule = spacing_el.attrib.get(_q("w", "lineRule"), "auto")
    line_raw = spacing_el.attrib.get(_q("w", "line"))
    if line_raw is None:
        return None, None
    val = int(line_raw)
    if rule == "exact":
        return 4, round(val / 20.0, 2)
    if rule == "atLeast":
        return 3, round(val / 20.0, 2)
    # auto (240 = single, 360 = 1.5, 480 = double, else multiple -> LinesToPoints)
    if val == 240:
        return 0, 12.0
    if val == 360:
        return 1, 18.0
    if val == 480:
        return 2, 24.0
    return 5, round((val / 240.0) * 12.0, 2)


def _parse_word_styles(z: zipfile.ZipFile, names: set[str]) -> dict[str, Any]:
    if "word/styles.xml" not in names:
        return {}
    s_root = ET.fromstring(z.read("word/styles.xml"))
    by_id: dict[str, dict[str, Any]] = {}
    aliases: dict[str, str] = {}

    for st in s_root.findall(_q("w", "style")):
        if st.attrib.get(_q("w", "type")) != "paragraph":
            continue
        sid = st.attrib.get(_q("w", "styleId"), "")
        name_el = st.find(_q("w", "name"))
        sname = name_el.attrib.get(_q("w", "val"), "") if name_el is not None else sid
        based_el = st.find(_q("w", "basedOn"))
        based_on = based_el.attrib.get(_q("w", "val")) if based_el is not None else None

        rPr = st.find(_q("w", "rPr"))
        pPr = st.find(_q("w", "pPr"))
        sz_el = rPr.find(_q("w", "sz")) if rPr is not None else None
        b_el = rPr.find(_q("w", "b")) if rPr is not None else None
        rf_el = rPr.find(_q("w", "rFonts")) if rPr is not None else None
        fonts = []
        if rf_el is not None:
            for attr in ("ascii", "eastAsia", "hAnsi", "cs"):
                fv = rf_el.attrib.get(_q("w", attr))
                if fv and fv not in fonts:
                    fonts.append(fv)

        jc_el = pPr.find(_q("w", "jc")) if pPr is not None else None
        sp_el = pPr.find(_q("w", "spacing")) if pPr is not None else None
        ind_el = pPr.find(_q("w", "ind")) if pPr is not None else None
        bdr_el = pPr.find(f"{_q('w', 'pBdr')}/{_q('w', 'bottom')}") if pPr is not None else None

        rule, lval = _word_line_rule_and_val(sp_el)
        info = {
            "style_id": sid,
            "name": sname,
            "based_on": based_on,
            "font_size": int(sz_el.attrib[_q("w", "val")]) / 2.0 if sz_el is not None and _q("w", "val") in sz_el.attrib else None,
            "font_names": fonts,
            "bold": (b_el.attrib.get(_q("w", "val"), "1") not in ("0", "false")) if b_el is not None else None,
            "align": _word_align_to_int(jc_el.attrib.get(_q("w", "val"))) if jc_el is not None else None,
            "space_before": round(int(sp_el.attrib[_q("w", "before")]) / 20.0, 2) if sp_el is not None and _q("w", "before") in sp_el.attrib else None,
            "space_after": round(int(sp_el.attrib[_q("w", "after")]) / 20.0, 2) if sp_el is not None and _q("w", "after") in sp_el.attrib else None,
            "line_unit_before": round(int(sp_el.attrib[_q("w", "beforeLines")]) / 100.0, 2) if sp_el is not None and _q("w", "beforeLines") in sp_el.attrib else None,
            "line_unit_after": round(int(sp_el.attrib[_q("w", "afterLines")]) / 100.0, 2) if sp_el is not None and _q("w", "afterLines") in sp_el.attrib else None,
            "line_spacing_rule": rule,
            "line_spacing": lval,
            "first_line_chars": round(int(ind_el.attrib[_q("w", "firstLineChars")]) / 100.0, 2) if ind_el is not None and _q("w", "firstLineChars") in ind_el.attrib else None,
            "bottom_border": {
                "has_border": bdr_el is not None and bdr_el.attrib.get(_q("w", "val"), "single") not in ("none", "nil"),
                "line_width": int(bdr_el.attrib.get(_q("w", "sz"), "0")) if bdr_el is not None and _q("w", "sz") in bdr_el.attrib else 0,
            } if bdr_el is not None else None,
        }
        by_id[sid] = info

        low_name = sname.lower()
        low_id = sid.lower()
        aliases[low_id] = sid
        aliases[low_name] = sid
        if st.attrib.get(_q("w", "default")) == "1" or low_name == "normal":
            aliases.setdefault("-1", sid)
            aliases.setdefault("正文", sid)
        if low_name == "heading 1" or low_id in ("heading1", "1"):
            aliases.setdefault("-2", sid)
            aliases.setdefault("标题 1", sid)
            aliases.setdefault("标题1", sid)
        if low_name == "heading 2" or low_id in ("heading2", "2"):
            aliases.setdefault("-3", sid)
            aliases.setdefault("标题 2", sid)
            aliases.setdefault("标题2", sid)
        if low_name == "heading 3" or low_id in ("heading3", "3"):
            aliases.setdefault("-4", sid)
            aliases.setdefault("标题 3", sid)
            aliases.setdefault("标题3", sid)

    # Lightly resolve basedOn inheritance
    resolved: dict[str, dict[str, Any]] = {}
    for sid, raw_info in by_id.items():
        merged = dict(raw_info)
        cur = raw_info.get("based_on")
        seen = {sid}
        while cur and cur in by_id and cur not in seen:
            seen.add(cur)
            parent = by_id[cur]
            for k in (
                "font_size", "bold", "align", "space_before", "space_after",
                "line_unit_before", "line_unit_after", "line_spacing_rule",
                "line_spacing", "first_line_chars", "bottom_border",
            ):
                if merged.get(k) is None and parent.get(k) is not None:
                    merged[k] = parent[k]
            if not merged.get("font_names") and parent.get("font_names"):
                merged["font_names"] = list(parent["font_names"])
            cur = parent.get("based_on")
        resolved[sid] = merged

    return {"by_id": resolved, "aliases": aliases}


def _count_header_shapes(h_root: ET.Element) -> int:
    """Count visual shapes in header XML without double-counting mc:Choice + mc:Fallback."""
    count = 0

    def _walk(node: ET.Element) -> None:
        nonlocal count
        tag = node.tag
        if tag == _q("mc", "AlternateContent"):
            choice = node.find(_q("mc", "Choice"))
            if choice is not None:
                _walk(choice)
            else:
                fb = node.find(_q("mc", "Fallback"))
                if fb is not None:
                    _walk(fb)
            return
        if tag in (_q("w", "drawing"), _q("w", "pict")):
            count += 1
            return
        for ch in node:
            _walk(ch)

    _walk(h_root)
    return count


def _extract_word_tables(doc_root: ET.Element) -> list[dict[str, Any]]:
    tables: list[dict[str, Any]] = []
    for tbl in doc_root.findall(f".//{_q('w', 'tbl')}"):
        tblPr = tbl.find(_q("w", "tblPr"))
        jc = tblPr.find(_q("w", "jc")) if tblPr is not None else None
        align = _word_align_to_int(jc.attrib.get(_q("w", "val"))) if jc is not None else 0
        grid_cols = tbl.findall(f"{_q('w', 'tblGrid')}/{_q('w', 'gridCol')}")
        trs = tbl.findall(_q("w", "tr"))
        rows_cnt = len(trs)
        cols_cnt = len(grid_cols)
        heading_repeat = False
        if trs:
            trPr = trs[0].find(_q("w", "trPr"))
            th = trPr.find(_q("w", "tblHeader")) if trPr is not None else None
            if th is not None and th.attrib.get(_q("w", "val"), "1") not in ("0", "false"):
                heading_repeat = True

        cells_map: dict[str, dict[str, Any]] = {}
        max_cols = cols_cnt
        for r_idx, tr in enumerate(trs, 1):
            c_idx = 1
            for tc in tr.findall(_q("w", "tc")):
                tcPr = tc.find(_q("w", "tcPr"))
                gs_el = tcPr.find(_q("w", "gridSpan")) if tcPr is not None else None
                span = int(gs_el.attrib.get(_q("w", "val"), "1")) if gs_el is not None else 1
                shd_el = tcPr.find(_q("w", "shd")) if tcPr is not None else None
                fill_hex = shd_el.attrib.get(_q("w", "fill")) if shd_el is not None else None
                bg_bgr = hex_to_bgr_int(fill_hex)
                txt = "".join(t.text or "" for t in tc.findall(f".//{_q('w', 't')}"))
                rf_el = tc.find(f".//{_q('w', 'rFonts')}")
                font_name = None
                if rf_el is not None:
                    font_name = (
                        rf_el.attrib.get(_q("w", "eastAsia"))
                        or rf_el.attrib.get(_q("w", "ascii"))
                        or rf_el.attrib.get(_q("w", "hAnsi"))
                    )
                cell_data = {"text": txt, "bg_color_bgr": bg_bgr, "font_name": font_name}
                for offset in range(span):
                    cells_map[f"{r_idx},{c_idx + offset}"] = cell_data
                c_idx += span
            max_cols = max(max_cols, c_idx - 1)

        tables.append({
            "rows": rows_cnt,
            "cols": cols_cnt or max_cols,
            "align": align if align is not None else 0,
            "heading_repeat": heading_repeat,
            "cells": cells_map,
        })
    return tables


def audit_word_xml(docx_path: str) -> dict[str, Any]:
    """Read-only OOXML audit for Word (.docx / zip-based .wps)."""
    with zipfile.ZipFile(docx_path, "r") as z:
        names = set(z.namelist())
        doc_root = ET.fromstring(z.read("word/document.xml"))
        rels = _parse_rels(z, "word/_rels/document.xml.rels")
        styles_bundle = _parse_word_styles(z, names)

        gutter_pos = 0
        if "word/settings.xml" in names:
            set_root = ET.fromstring(z.read("word/settings.xml"))
            gat = set_root.find(_q("w", "gutterAtTop"))
            if gat is not None and gat.attrib.get(_q("w", "val"), "1") not in ("0", "false"):
                gutter_pos = 1

        sections: list[dict[str, Any]] = []
        start_type_map = {"continuous": 0, "newColumn": 1, "nextPage": 2, "evenPage": 3, "oddPage": 4}
        for idx, sect in enumerate(doc_root.findall(f".//{_q('w', 'sectPr')}"), 1):
            pgSz = sect.find(_q("w", "pgSz"))
            pgMar = sect.find(_q("w", "pgMar"))
            typ = sect.find(_q("w", "type"))
            start_str = typ.attrib.get(_q("w", "val")) if typ is not None else "nextPage"
            sz_dict = {k.split("}")[-1]: v for k, v in (pgSz.attrib.items() if pgSz is not None else [])}
            mar_dict = {k.split("}")[-1]: v for k, v in (pgMar.attrib.items() if pgMar is not None else [])}
            hdr_refs = []
            for hr in sect.findall(_q("w", "headerReference")):
                rid = hr.attrib.get(_q("r", "id"), "")
                htype = hr.attrib.get(_q("w", "type"), "default")
                target = rels.get(rid, "")
                hdr_entry = f"word/{target}" if target and not target.startswith("word/") else target
                hdr_info: dict[str, Any] = {"type": htype, "target": hdr_entry}
                if hdr_entry and hdr_entry in names:
                    hbytes = z.read(hdr_entry)
                    hraw = hbytes.decode("utf-8", errors="replace")
                    h_root = ET.fromstring(hbytes)
                    instrs = [el.text or "" for el in h_root.findall(f".//{_q('w', 'instrText')}")]
                    instrs.extend(
                        el.attrib.get(_q("w", "instr"), "")
                        for el in h_root.findall(f".//{_q('w', 'fldSimple')}")
                        if el.attrib.get(_q("w", "instr"))
                    )
                    wm_texts = [
                        el.attrib.get("string", "")
                        for el in h_root.findall(f".//{_q('v', 'textpath')}")
                        if el.attrib.get("string")
                    ]
                    wm_texts.extend(el.text for el in h_root.findall(f".//{_q('a', 't')}") if el.text)
                    p_texts = [el.text for el in h_root.findall(f".//{_q('w', 't')}") if el.text]
                    has_bdr = any(
                        b.attrib.get(_q("w", "val"), "single") not in ("none", "nil")
                        for b in h_root.findall(f".//{_q('w', 'pBdr')}/{_q('w', 'bottom')}")
                    )
                    has_washout = False
                    for imgd in h_root.findall(f".//{_q('v', 'imagedata')}"):
                        if imgd.attrib.get("gain") == "19661f" or imgd.attrib.get("blacklevel") == "22938f":
                            has_washout = True
                    for lum in h_root.findall(f".//{_q('a', 'lum')}"):
                        if lum.attrib.get("bright") == "70000" or lum.attrib.get("contrast") == "-70000":
                            has_washout = True

                    hdr_info["drawings"] = hraw.count("<w:drawing")
                    hdr_info["picts"] = hraw.count("<w:pict")
                    hdr_info["shape_count"] = _count_header_shapes(h_root)
                    hdr_info["behind_doc"] = re.findall(r'behindDoc="(\d)"', hraw)
                    hdr_info["instr_text"] = instrs
                    hdr_info["wm_texts"] = wm_texts
                    hdr_info["texts"] = p_texts
                    hdr_info["has_bottom_border"] = has_bdr
                    hdr_info["picture_washout"] = has_washout
                hdr_refs.append(hdr_info)

            has_default_hdr = any(h.get("type") == "default" for h in hdr_refs)
            link_to_prev = False if idx == 1 else (not has_default_hdr)
            sections.append({
                "section": idx,
                "start_type": start_str,
                "section_start": start_type_map.get(start_str or "nextPage", 2),
                "orientation": 1 if sz_dict.get("orient") == "landscape" else 0,
                "gutter_pos": gutter_pos,
                "page_width_cm": round(int(sz_dict["w"]) / 20.0 / PT_PER_CM, 4) if "w" in sz_dict else None,
                "page_height_cm": round(int(sz_dict["h"]) / 20.0 / PT_PER_CM, 4) if "h" in sz_dict else None,
                "margins_cm": [
                    round(int(mar_dict[k]) / 20.0 / PT_PER_CM, 4)
                    for k in ("top", "bottom", "left", "right")
                    if k in mar_dict
                ] if all(k in mar_dict for k in ("top", "bottom", "left", "right")) else None,
                "header_link_to_previous": link_to_prev,
                "pgSz": sz_dict,
                "pgMar": mar_dict,
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

        # Extract static paragraphs & tables for --no-com grading and diffing
        styles_by_id = styles_bundle.get("by_id", {})
        default_sid = styles_bundle.get("aliases", {}).get("-1")
        paragraphs: list[dict[str, Any]] = []
        body_el = doc_root.find(_q("w", "body"))
        all_texts: list[str] = []
        if body_el is not None:
            for p_el in body_el.findall(f".//{_q('w', 'p')}"):
                txt = "".join(t.text or "" for t in p_el.findall(f".//{_q('w', 't')}")).rstrip("\r\x07\x0c")
                all_texts.append(txt)
                if not txt.strip():
                    continue
                pPr = p_el.find(_q("w", "pPr"))
                ps_el = pPr.find(_q("w", "pStyle")) if pPr is not None else None
                sid = ps_el.attrib.get(_q("w", "val")) if ps_el is not None else default_sid
                st_info = styles_by_id.get(sid or "", {})
                jc_el = pPr.find(_q("w", "jc")) if pPr is not None else None
                ind_el = pPr.find(_q("w", "ind")) if pPr is not None else None
                sp_el = pPr.find(_q("w", "spacing")) if pPr is not None else None
                num_pr = pPr.find(_q("w", "numPr")) if pPr is not None else None
                num_id = num_pr.find(_q("w", "numId")).attrib.get(_q("w", "val")) if (num_pr is not None and num_pr.find(_q("w", "numId")) is not None) else None
                ilvl = num_pr.find(_q("w", "ilvl")).attrib.get(_q("w", "val")) if (num_pr is not None and num_pr.find(_q("w", "ilvl")) is not None) else None

                l_rule, l_val = _word_line_rule_and_val(sp_el)
                r_sz = p_el.find(f"{_q('w', 'r')}/{_q('w', 'rPr')}/{_q('w', 'sz')}")
                font_sz = int(r_sz.attrib[_q("w", "val")]) / 2.0 if r_sz is not None and _q("w", "val") in r_sz.attrib else st_info.get("font_size")

                paragraphs.append({
                    "text": txt[:120],
                    "full_text": txt,
                    "style_id": sid,
                    "style": st_info.get("name", sid or ""),
                    "align": _word_align_to_int(jc_el.attrib.get(_q("w", "val"))) if jc_el is not None else st_info.get("align"),
                    "c_indent": round(int(ind_el.attrib[_q("w", "firstLineChars")]) / 100.0, 2) if ind_el is not None and _q("w", "firstLineChars") in ind_el.attrib else st_info.get("first_line_chars"),
                    "line_rule": l_rule if l_rule is not None else st_info.get("line_spacing_rule"),
                    "line_spacing": l_val if l_val is not None else st_info.get("line_spacing"),
                    "unit_before": round(int(sp_el.attrib[_q("w", "beforeLines")]) / 100.0, 2) if sp_el is not None and _q("w", "beforeLines") in sp_el.attrib else st_info.get("line_unit_before"),
                    "unit_after": round(int(sp_el.attrib[_q("w", "afterLines")]) / 100.0, 2) if sp_el is not None and _q("w", "afterLines") in sp_el.attrib else st_info.get("line_unit_after"),
                    "font_size": font_sz,
                    "num_id": num_id,
                    "ilvl": int(ilvl) if ilvl is not None else None,
                })

        tables = _extract_word_tables(doc_root)

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
            "styles": styles_bundle,
            "paragraphs": paragraphs,
            "full_text": "\n".join(all_texts),
            "tables": tables,
            "charts": charts,
        }


def _ppt_bg_fill_type(bg_el: ET.Element | None) -> int | None:
    if bg_el is None:
        return None
    bg_pr = bg_el.find(_q("p", "bgPr"))
    if bg_pr is not None:
        if bg_pr.find(_q("a", "blipFill")) is not None:
            return 6
        if bg_pr.find(_q("a", "gradFill")) is not None:
            return 3
        if bg_pr.find(_q("a", "pattFill")) is not None:
            return 2
        if bg_pr.find(_q("a", "solidFill")) is not None:
            return 1
    if bg_el.find(_q("p", "bgRef")) is not None:
        return 1
    return None


def _parse_ppt_transition(s_root: ET.Element) -> dict[str, Any]:
    trans = s_root.find(_q("p", "transition"))
    if trans is None:
        for ac in s_root.findall(f".//{_q('mc', 'AlternateContent')}"):
            ch = ac.find(_q("mc", "Choice"))
            if ch is not None and ch.find(_q("p", "transition")) is not None:
                trans = ch.find(_q("p", "transition"))
                break
    if trans is None:
        return {"present": False, "effect_tag": None, "entry_effect": 0, "advance_on_time": False, "advance_time": 0.0}

    eff_tag = None
    eff_code = 0
    for child in trans:
        local = child.tag.split("}")[-1]
        if local == "sndAc":
            continue
        if local == "prstTrans":
            eff_tag = child.attrib.get("prst", local)
        else:
            eff_tag = local
        eff_code = _PPT_TRANSITION_EFFECT_MAP.get(eff_tag, 0)
        break

    adv_tm = trans.attrib.get("advTm")
    adv_on_time = adv_tm is not None
    adv_sec = round(int(adv_tm) / 1000.0, 3) if adv_tm is not None else 0.0
    return {
        "present": True,
        "effect_tag": eff_tag,
        "entry_effect": eff_code,
        "advance_on_time": adv_on_time,
        "advance_time": adv_sec,
    }


def _first_el(*elements: ET.Element | None) -> ET.Element | None:
    for el in elements:
        if el is not None:
            return el
    return None


def _parse_ppt_shapes(
    s_root: ET.Element,
    z: zipfile.ZipFile,
    slide_rels: dict[str, str],
    names: set[str],
) -> list[dict[str, Any]]:
    sp_tree = s_root.find(f"{_q('p', 'cSld')}/{_q('p', 'spTree')}")
    if sp_tree is None:
        return []
    algn_map = {"l": 1, "ctr": 2, "r": 3, "just": 4}
    out: list[dict[str, Any]] = []

    for child in sp_tree:
        local = child.tag.split("}")[-1]
        if local not in ("sp", "pic", "graphicFrame", "grpSp"):
            continue
        cnv = _first_el(
            child.find(f"{_q('p', 'nvSpPr')}/{_q('p', 'cNvPr')}"),
            child.find(f"{_q('p', 'nvPicPr')}/{_q('p', 'cNvPr')}"),
            child.find(f"{_q('p', 'nvGraphicFramePr')}/{_q('p', 'cNvPr')}"),
            child.find(f"{_q('p', 'nvGrpSpPr')}/{_q('p', 'cNvPr')}"),
        )
        name = cnv.attrib.get("name") if cnv is not None else None
        xfrm = _first_el(
            child.find(f"{_q('p', 'spPr')}/{_q('a', 'xfrm')}"),
            child.find(_q("p", "xfrm")),
            child.find(f"{_q('p', 'grpSpPr')}/{_q('a', 'xfrm')}"),
        )
        off = xfrm.find(_q("a", "off")) if xfrm is not None else None
        ext = xfrm.find(_q("a", "ext")) if xfrm is not None else None
        left_pt = round(int(off.attrib.get("x", "0")) / EMU_PER_PT, 2) if off is not None else 0.0
        top_pt = round(int(off.attrib.get("y", "0")) / EMU_PER_PT, 2) if off is not None else 0.0
        w_pt = round(int(ext.attrib.get("cx", "0")) / EMU_PER_PT, 2) if ext is not None else 0.0
        h_pt = round(int(ext.attrib.get("cy", "0")) / EMU_PER_PT, 2) if ext is not None else 0.0

        has_table = child.find(f".//{_q('a', 'tbl')}") is not None
        has_chart = child.find(f".//{_q('c', 'chart')}") is not None
        rel_ids_el = child.find(f".//{_q('dgm', 'relIds')}")
        has_smartart = rel_ids_el is not None
        smartart_nodes: list[str] = []
        if rel_ids_el is not None:
            dm_rid = rel_ids_el.attrib.get(_q("r", "dm"), "")
            dm_target = slide_rels.get(dm_rid, "")
            dm_entry = _resolve_package_target("ppt/slides", dm_target) if dm_target else ""
            if dm_entry and dm_entry in names:
                d_root = ET.fromstring(z.read(dm_entry))
                for pt_el in d_root.findall(f".//{_q('dgm', 'pt')}"):
                    if pt_el.attrib.get("type") in ("doc", "parTrans", "sibTrans", "pres"):
                        continue
                    tx = "".join(t.text or "" for t in pt_el.findall(f".//{_q('a', 't')}")).strip()
                    if tx:
                        smartart_nodes.append(tx)

        if local == "pic":
            stype = 13
        elif local == "graphicFrame":
            stype = 24 if has_smartart else (3 if has_chart else (19 if has_table else 1))
        elif local == "grpSp":
            stype = 6
        else:
            nv_pr = child.find(f"{_q('p', 'nvSpPr')}/{_q('p', 'nvPr')}")
            c_nv_sp = child.find(f"{_q('p', 'nvSpPr')}/{_q('p', 'cNvSpPr')}")
            if nv_pr is not None and nv_pr.find(_q("p", "ph")) is not None:
                stype = 14
            elif c_nv_sp is not None and c_nv_sp.attrib.get("txBox") == "1":
                stype = 17
            else:
                stype = 1

        texts = [t.text for t in child.findall(f"{_q('p', 'txBody')}//{_q('a', 't')}") if t.text]
        p_pr = child.find(f"{_q('p', 'txBody')}/{_q('a', 'p')}/{_q('a', 'pPr')}")
        align = algn_map.get(p_pr.attrib.get("algn")) if p_pr is not None and "algn" in p_pr.attrib else None
        r_pr = child.find(f"{_q('p', 'txBody')}//{_q('a', 'rPr')}")
        smallcaps = (r_pr.attrib.get("cap") == "small") if r_pr is not None else False
        patt_el = r_pr.find(_q("a", "pattFill")) if r_pr is not None else None
        patt_code = _PPT_PATT_PRST_MAP.get(patt_el.attrib.get("prst", "")) if patt_el is not None else None
        shd_el = r_pr.find(f"{_q('a', 'effectLst')}/{_q('a', 'outerShdw')}") if r_pr is not None else None

        out.append({
            "name": name,
            "tag": local,
            "type": stype,
            "left": left_pt,
            "top": top_pt,
            "w": w_pt,
            "h": h_pt,
            "text": "".join(texts).strip(),
            "align": align,
            "smallcaps": smallcaps,
            "font_fill_pattern": patt_code,
            "shadow_visible": shd_el is not None,
            "has_table": has_table,
            "has_chart": has_chart,
            "has_smartart": has_smartart,
            "smartart_nodes": smartart_nodes,
        })
    return out


def audit_ppt_xml(pptx_path: str) -> dict[str, Any]:
    """Read-only OOXML audit for Presentation (.pptx)."""
    with zipfile.ZipFile(pptx_path, "r") as z:
        names = set(z.namelist())
        pres_root = ET.fromstring(z.read("ppt/presentation.xml"))
        sld_sz = pres_root.find(_q("p", "sldSz"))
        sz_info = {k: v for k, v in (sld_sz.attrib.items() if sld_sz is not None else [])}
        width_pt = round(int(sz_info["cx"]) / EMU_PER_PT, 1) if "cx" in sz_info else None
        height_pt = round(int(sz_info["cy"]) / EMU_PER_PT, 1) if "cy" in sz_info else None
        sz_type = sz_info.get("type", "custom")
        slide_size_code = {"screen16x9": 15, "screen4x3": 1}.get(sz_type, 7)

        master_bg_type = None
        if "ppt/slideMasters/slideMaster1.xml" in names:
            m_root = ET.fromstring(z.read("ppt/slideMasters/slideMaster1.xml"))
            master_bg_type = _ppt_bg_fill_type(m_root.find(f"{_q('p', 'cSld')}/{_q('p', 'bg')}"))

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
            s_rels = _parse_rels(z, f"ppt/slides/_rels/{Path(sf).name}.rels")
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

            trans_info = _parse_ppt_transition(s_root)
            s_bg = s_root.find(f"{_q('p', 'cSld')}/{_q('p', 'bg')}")
            slides_info.append({
                "file": sf,
                "hidden": s_root.attrib.get("show") == "0",
                "has_transition": trans_info["present"] or (s_root.find(f".//{_q('p', 'transition')}") is not None),
                "transition": trans_info,
                "has_timing": s_root.find(f".//{_q('p', 'timing')}") is not None,
                "follow_master_bg": s_bg is None,
                "bg_fill_type": _ppt_bg_fill_type(s_bg),
                "body_prs": body_prs,
                "shapes_body_pr": shapes_body_pr,
                "shapes": _parse_ppt_shapes(s_root, z, s_rels, names),
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
            "slide_size": slide_size_code,
            "width_pt": width_pt,
            "height_pt": height_pt,
            "master_bg_fill_type": master_bg_type,
            "prnPr": prn_info,
            "slides_count": len(slides_info),
            "slides": slides_info,
            "charts": charts,
        }


def _parse_xl_dxf_colors(z: zipfile.ZipFile, names: set[str]) -> list[int | None]:
    if "xl/styles.xml" not in names:
        return []
    st_root = ET.fromstring(z.read("xl/styles.xml"))
    colors: list[int | None] = []
    dxfs = st_root.find(_q("x", "dxfs"))
    if dxfs is None:
        return colors
    for dxf in dxfs.findall(_q("x", "dxf")):
        pf = dxf.find(f"{_q('x', 'fill')}/{_q('x', 'patternFill')}")
        bgr = None
        if pf is not None:
            for ctag in ("bgColor", "fgColor"):
                cel = pf.find(_q("x", ctag))
                if cel is not None and "rgb" in cel.attrib:
                    bgr = hex_to_bgr_int(cel.attrib["rgb"])
                    if bgr is not None:
                        break
        colors.append(bgr)
    return colors


def audit_excel_xml(xlsx_path: str) -> dict[str, Any]:
    """Read-only OOXML audit for Spreadsheet (.xlsx / zip-based .et)."""
    with zipfile.ZipFile(xlsx_path, "r") as z:
        names = set(z.namelist())
        wb_root = ET.fromstring(z.read("xl/workbook.xml"))
        wb_rels = _parse_rels(z, "xl/_rels/workbook.xml.rels")
        dxf_colors = _parse_xl_dxf_colors(z, names)

        sheets_meta: list[dict[str, Any]] = []
        target_to_name: dict[str, str] = {}
        ordered_ws_targets: list[tuple[str, str]] = []

        for s in wb_root.findall(f".//{_q('x', 'sheet')}"):
            sname = s.attrib.get("name") or ""
            rid = s.attrib.get(_q("r", "id"), "")
            raw_target = wb_rels.get(rid, "")
            target = _resolve_package_target("xl", raw_target) if raw_target else None
            sheets_meta.append({
                "name": sname,
                "sheetId": s.attrib.get("sheetId"),
                "state": s.attrib.get("state", "visible"),
                "rId": rid,
                "file": target,
            })
            if target and target in names and target.startswith("xl/worksheets/"):
                target_to_name[target] = sname
                ordered_ws_targets.append((target, sname))

        # Fallback only if rels were absent or unmapped (e.g. legacy synthetic zips)
        if not ordered_ws_targets:
            sheet_files = sorted(
                (n for n in names if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)),
                key=lambda s: int(re.search(r"(\d+)", s).group(1)),
            )
            for idx, sf in enumerate(sheet_files):
                sname = sheets_meta[idx]["name"] if idx < len(sheets_meta) else Path(sf).stem
                target_to_name[sf] = sname
                ordered_ws_targets.append((sf, sname))

        # Map worksheet -> pivotTable files via worksheet rels
        pivot_sheet_map: dict[str, str] = {}
        for sf, sname in ordered_ws_targets:
            ws_rels_path = f"xl/worksheets/_rels/{Path(sf).name}.rels"
            for _, rel_t in _parse_rels(z, ws_rels_path).items():
                resolved_t = _resolve_package_target("xl/worksheets", rel_t)
                if resolved_t.startswith("xl/pivotTables/"):
                    pivot_sheet_map[resolved_t] = sname

        sheets_xml: list[dict[str, Any]] = []
        for sf, sname in ordered_ws_targets:
            s_root = ET.fromstring(z.read(sf))
            cfs = []
            for cf in s_root.findall(_q("x", "conditionalFormatting")):
                rules = []
                for r_el in cf.findall(_q("x", "cfRule")):
                    fmls = [f.text for f in r_el.findall(_q("x", "formula"))]
                    dxf_id_raw = r_el.attrib.get("dxfId")
                    dxf_bgr = None
                    if dxf_id_raw is not None and dxf_id_raw.isdigit():
                        didx = int(dxf_id_raw)
                        if 0 <= didx < len(dxf_colors):
                            dxf_bgr = dxf_colors[didx]
                    rules.append({
                        "type": r_el.attrib.get("type"),
                        "operator": r_el.attrib.get("operator"),
                        "dxfId": dxf_id_raw,
                        "dxf_color_bgr": dxf_bgr,
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
                "name": sname,
                "conditional_formatting": cfs,
                "sort_state": sort_state,
                "auto_filter": af_el.attrib.get("ref") if af_el is not None else None,
            })

        # Parse pivot caches for field names
        cache_fields_by_id: dict[str, list[str]] = {}
        for pc_el in wb_root.findall(f".//{_q('x', 'pivotCache')}"):
            cid = pc_el.attrib.get("cacheId", "")
            crid = pc_el.attrib.get(_q("r", "id"), "")
            ctarget = wb_rels.get(crid, "")
            centry = _resolve_package_target("xl", ctarget) if ctarget else ""
            if centry and centry in names:
                c_root = ET.fromstring(z.read(centry))
                cache_fields_by_id[cid] = [
                    cf.attrib.get("name", "")
                    for cf in c_root.findall(f".//{_q('x', 'cacheFields')}/{_q('x', 'cacheField')}")
                ]

        pivots = []
        for pn in sorted(n for n in names if n.startswith("xl/pivotTables/pivotTable") and n.endswith(".xml")):
            pr = ET.fromstring(z.read(pn))
            loc = pr.find(_q("x", "location"))
            cid = pr.attrib.get("cacheId", "")
            cfields = cache_fields_by_id.get(cid, [])
            row_fields = []
            for rf in pr.findall(f"{_q('x', 'rowFields')}/{_q('x', 'field')}"):
                x_idx = rf.attrib.get("x", "")
                if x_idx.lstrip("-").isdigit():
                    idx_i = int(x_idx)
                    if 0 <= idx_i < len(cfields):
                        row_fields.append(cfields[idx_i])
            data_fields = [
                df.attrib.get("name", "")
                for df in pr.findall(f"{_q('x', 'dataFields')}/{_q('x', 'dataField')}")
            ]
            pivots.append({
                "file": pn,
                "sheet": pivot_sheet_map.get(pn),
                "name": pr.attrib.get("name"),
                "cacheId": cid,
                "location": loc.attrib.get("ref") if loc is not None else None,
                "row_fields": row_fields,
                "data_fields": data_fields,
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


def read_excel_static(xlsx_path: str) -> dict[str, Any]:
    """Read-only static formula, cached-value, and number-format reader via openpyxl."""
    abs_path = os.path.abspath(xlsx_path)
    if openpyxl is None or not zipfile.is_zipfile(abs_path):
        return {}
    with zipfile.ZipFile(abs_path, "r") as z:
        if "xl/workbook.xml" not in z.namelist():
            return {}

    sheets: dict[str, dict[str, Any]] = {}
    sheet_names: list[str] = []

    with open(abs_path, "rb") as fp_f:
        wb_f = openpyxl.load_workbook(fp_f, read_only=True, data_only=False)
        try:
            sheet_names = list(wb_f.sheetnames)
            for sname in sheet_names:
                ws = wb_f[sname]
                formulas: dict[tuple[int, int], Any] = {}
                number_formats: dict[tuple[int, int], str] = {}
                for row in ws.iter_rows():
                    for cell in row:
                        r = getattr(cell, "row", None)
                        c = getattr(cell, "column", None)
                        if r is None or c is None:
                            continue
                        if cell.value is not None:
                            formulas[(int(r), int(c))] = cell.value
                        if cell.number_format and cell.number_format != "General":
                            number_formats[(int(r), int(c))] = str(cell.number_format)
                sheets[sname] = {
                    "formulas": formulas,
                    "values": {},
                    "number_formats": number_formats,
                    "max_row": ws.max_row or 0,
                    "max_col": ws.max_column or 0,
                }
        finally:
            wb_f.close()

    with open(abs_path, "rb") as fp_v:
        wb_v = openpyxl.load_workbook(fp_v, read_only=True, data_only=True)
        try:
            for sname in wb_v.sheetnames:
                if sname not in sheets:
                    continue
                ws = wb_v[sname]
                vals: dict[tuple[int, int], Any] = {}
                for row in ws.iter_rows():
                    for cell in row:
                        r = getattr(cell, "row", None)
                        c = getattr(cell, "column", None)
                        if r is None or c is None:
                            continue
                        if cell.value is not None:
                            vals[(int(r), int(c))] = cell.value
                sheets[sname]["values"] = vals
        finally:
            wb_v.close()

    return {"sheet_names": sheet_names, "sheets": sheets}


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

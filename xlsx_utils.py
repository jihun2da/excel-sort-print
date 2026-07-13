# -*- coding: utf-8 -*-
"""
엑셀 읽기/정렬/서식-보존 재분류 저장을 담당하는 유틸 모듈.

핵심 아이디어: 셀 값만 뽑아 새 시트를 만들면 서식(배경색, 테두리, 필터, 열너비 등)이
사라지므로, 원본 xlsx(zip) 안의 <row> XML 요소 자체를 그대로 재배치/재분류한다.
이렇게 하면 각 행이 갖고 있던 스타일 참조, 공유문자열 참조, 서식이 전혀 손상되지 않는다.
"""
import io
import re
import zipfile
from dataclasses import dataclass

import openpyxl
from lxml import etree

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

# 컬럼 인덱스 (0-based): A=0, E=4, F=5, G=6, H=7, I=8, J=9, K=10
COL = {"A": 0, "E": 4, "F": 5, "G": 6, "H": 7, "I": 8, "J": 9, "K": 10}

# 엑셀 레거시 인덱스 컬러(0-63) 중 자주 쓰이는 항목 (근사치)
INDEXED_COLORS = [
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
    "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
    "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
    "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
    "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
    "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
    "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
]


# ---------- 색상 유틸 ----------
def _hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip("#")
    if len(hex_str) == 8:
        hex_str = hex_str[2:]
    if len(hex_str) != 6:
        return (255, 255, 255)
    return tuple(int(hex_str[i:i + 2], 16) for i in (0, 2, 4))


def _rgb_to_hex(r, g, b):
    def c(v):
        return max(0, min(255, round(v)))
    return "{:02X}{:02X}{:02X}".format(c(r), c(g), c(b))


def _apply_tint(hex_str, tint):
    if not tint:
        return hex_str
    import colorsys
    r, g, b = _hex_to_rgb(hex_str)
    h, l, s = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
    if tint < 0:
        l = l * (1 + tint)
    else:
        l = l * (1 - tint) + tint
    r2, g2, b2 = colorsys.hls_to_rgb(h, l, s)
    return _rgb_to_hex(r2 * 255, g2 * 255, b2 * 255)


def get_theme_colors(zf):
    """실제 파일의 xl/theme/theme1.xml 을 읽어 셀 스타일 theme 인덱스(0~11) 순서로 매핑한다.
    (styles.xml의 <color theme="n">은 0=lt1,1=dk1,2=lt2,3=dk2,4-9=accent1-6,10=hlink,11=folHlink 순서)
    """
    try:
        theme_xml = zf.read("xl/theme/theme1.xml")
    except KeyError:
        return ["FFFFFF", "000000", "FFFFFF", "000000", "000000", "000000",
                "000000", "000000", "000000", "000000", "0000FF", "800080"]
    root = etree.fromstring(theme_xml)
    ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    scheme = root.find(".//a:clrScheme", ns)
    raw = {}
    if scheme is not None:
        for tag in ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
                    "accent4", "accent5", "accent6", "hlink", "folHlink"]:
            el = scheme.find(f"a:{tag}", ns)
            if el is None:
                continue
            srgb = el.find("a:srgbClr", ns)
            sysclr = el.find("a:sysClr", ns)
            if srgb is not None:
                raw[tag] = srgb.get("val")
            elif sysclr is not None:
                raw[tag] = sysclr.get("lastClr")
    return [
        raw.get("lt1", "FFFFFF"), raw.get("dk1", "000000"),
        raw.get("lt2", "FFFFFF"), raw.get("dk2", "000000"),
        raw.get("accent1", "000000"), raw.get("accent2", "000000"),
        raw.get("accent3", "000000"), raw.get("accent4", "000000"),
        raw.get("accent5", "000000"), raw.get("accent6", "000000"),
        raw.get("hlink", "0000FF"), raw.get("folHlink", "800080"),
    ]


def extract_fill_hex(cell, theme_colors):
    """openpyxl 셀의 배경색을 '#RRGGBB' 문자열로 추출. 색이 없으면 None."""
    fill = cell.fill
    if fill is None or fill.patternType != "solid":
        return None
    fg = fill.fgColor
    hexv = None
    try:
        if fg.type == "rgb" and isinstance(fg.rgb, str):
            hexv = fg.rgb
        elif fg.type == "theme":
            theme_idx = fg.theme
            base = theme_colors[theme_idx] if 0 <= theme_idx < len(theme_colors) else "FFFFFF"
            tint = fg.tint or 0.0
            hexv = _apply_tint(base, tint)
        elif fg.type == "indexed":
            idx = fg.indexed
            if idx in (64, 65):
                return None
            hexv = INDEXED_COLORS[idx] if 0 <= idx < len(INDEXED_COLORS) else None
    except Exception:
        return None
    if not hexv:
        return None
    hexv = hexv.lstrip("#")
    if len(hexv) == 8:
        hexv = hexv[2:]
    if hexv.upper() == "FFFFFF":
        return None
    return "#" + hexv.upper()


def text_color_for(bg_hex):
    if not bg_hex:
        return "#000000"
    r, g, b = _hex_to_rgb(bg_hex)
    lum = 0.299 * r + 0.587 * g + 0.114 * b
    return "#FFFFFF" if lum < 140 else "#000000"


# ---------- 정렬 유틸 ----------
def _tier_of(v):
    if v is None or v == "":
        return (0, "")
    if isinstance(v, (int, float)):
        return (1, v)
    s = str(v).strip()
    try:
        return (1, float(s))
    except ValueError:
        return (2, s.lower())


def _compare_val(a, b):
    ta, tb = _tier_of(a), _tier_of(b)
    if ta[0] != tb[0]:
        return -1 if ta[0] < tb[0] else 1
    if ta[0] == 1:
        return -1 if ta[1] < tb[1] else (1 if ta[1] > tb[1] else 0)
    sa, sb = str(ta[1]), str(tb[1])
    return -1 if sa < sb else (1 if sa > sb else 0)


def sort_sequential(rows, order):
    """order 의 열들을 Excel에서 하듯 각각 독립적으로, 차례대로(순차) 적용한다(각 단계는 안정 정렬).
    그 결과 가장 마지막에 적용한 열이 최종적으로 가장 우선순위가 높은 정렬 기준이 된다.
    (대량 분류 ['A','I','J','K'] 는 실제 사용자 로컬 정렬 결과와 1004행 전부 일치 검증됨)

    rows 의 각 원소는 {"cells": [...], "row_num": int, "a_color": str|None} 형태의 dict.
    """
    import functools
    result = list(rows)
    for key in order:
        idx = COL[key]

        def cmp(r1, r2, idx=idx):
            return _compare_val(r1["cells"][idx], r2["cells"][idx])

        result = sorted(result, key=functools.cmp_to_key(cmp))
    return result


# ---------- 데이터 로드 ----------
@dataclass
class LoadedSheet:
    header: list
    rows: list  # list of dict: {"cells": [...], "row_num": int, "a_color": str|None}


def load_sheet_rows(raw_bytes, sheet_name):
    wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
    ws = wb[sheet_name]
    with zipfile.ZipFile(io.BytesIO(raw_bytes)) as zf:
        theme_colors = get_theme_colors(zf)

    all_rows = list(ws.iter_rows(min_row=1))
    header = [c.value for c in all_rows[0]] if all_rows else []
    rows = []
    for r_idx, row_cells in enumerate(all_rows[1:], start=2):
        values = [c.value for c in row_cells]
        if not any(v is not None and v != "" for v in values):
            continue
        a_color = extract_fill_hex(row_cells[0], theme_colors) if row_cells else None
        rows.append({"cells": values, "row_num": r_idx, "a_color": a_color})
    return LoadedSheet(header=header, rows=rows)


# ---------- 서식 보존 다중 시트 저장 ----------
def _local(tag):
    return f"{{{MAIN_NS}}}{tag}"


def _get_sheet_path(zf, sheet_name):
    wb_xml = zf.read("xl/workbook.xml")
    rels_xml = zf.read("xl/_rels/workbook.xml.rels")
    wb_root = etree.fromstring(wb_xml)
    rels_root = etree.fromstring(rels_xml)
    sheets_el = wb_root.find(_local("sheets"))
    target = None
    for sh in sheets_el.findall(_local("sheet")):
        if sh.get("name") == sheet_name:
            target = sh
            break
    if target is None:
        raise ValueError(f"시트를 찾을 수 없습니다: {sheet_name}")
    rid = target.get(f"{{{R_NS}}}id")
    rel = None
    for r in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship"):
        if r.get("Id") == rid:
            rel = r
            break
    if rel is None:
        raise ValueError("시트 경로를 찾을 수 없습니다.")
    target_path = rel.get("Target").lstrip("/")
    if not target_path.startswith("xl/"):
        target_path = "xl/" + target_path
    return target_path


def _sanitize_sheet_name_factory(existing_names):
    used = set(existing_names)

    def sanitize(raw, fallback):
        n = re.sub(r'[\\/\?\*\[\]:]', "_", (raw if raw else fallback)).strip()
        if not n:
            n = fallback
        if len(n) > 31:
            n = n[:31]
        base, i = n, 2
        while n in used:
            suffix = f"_{i}"
            n = base[:max(1, 31 - len(suffix))] + suffix
            i += 1
        used.add(n)
        return n
    return sanitize


def build_multi_sheet_xlsx(raw_bytes, original_sheet_name, sheet_groups):
    """sheet_groups: [{"name": str, "row_nums": [int, ...]}] (원본 엑셀의 1-based 행 번호, 원하는 출력 순서)
    첫 번째 그룹은 기존 시트 자리를 재사용(이름만 변경)하고, 나머지는 새 시트 탭으로 추가한다.
    반환값: 새 xlsx 파일의 bytes
    """
    zin = zipfile.ZipFile(io.BytesIO(raw_bytes), "r")
    sheet_path = _get_sheet_path(zin, original_sheet_name)
    template_xml = zin.read(sheet_path)

    wb_root = etree.fromstring(zin.read("xl/workbook.xml"))
    rels_root = etree.fromstring(zin.read("xl/_rels/workbook.xml.rels"))
    ct_root = etree.fromstring(zin.read("[Content_Types].xml"))

    sheets_el = wb_root.find(_local("sheets"))
    sheet_els = sheets_el.findall(_local("sheet"))
    target_el = None
    for sh in sheet_els:
        if sh.get("name") == original_sheet_name:
            target_el = sh
            break
    if target_el is None:
        raise ValueError("원본 시트를 찾을 수 없습니다.")

    max_sheet_id = 0
    for sh in sheet_els:
        try:
            max_sheet_id = max(max_sheet_id, int(sh.get("sheetId")))
        except (TypeError, ValueError):
            pass

    rel_els = rels_root.findall(f"{{{PKG_REL_NS}}}Relationship")
    max_rid = 0
    for r in rel_els:
        m = re.match(r"^rId(\d+)$", r.get("Id") or "")
        if m:
            max_rid = max(max_rid, int(m.group(1)))

    max_sheet_file_idx = 0
    for r in rel_els:
        m = re.search(r"worksheets/sheet(\d+)\.xml$", r.get("Target") or "")
        if m:
            max_sheet_file_idx = max(max_sheet_file_idx, int(m.group(1)))

    existing_names = [sh.get("name") for sh in sheet_els if sh is not target_el]  # 자기 자신의 현재 이름은 충돌로 치지 않음
    sanitize = _sanitize_sheet_name_factory(existing_names)

    def build_filtered_sheet_xml(row_nums):
        doc = etree.fromstring(template_xml)
        sheet_data = doc.find(_local("sheetData"))
        row_els = sheet_data.findall(_local("row"))
        row_map = {}
        header_el = None
        for el in row_els:
            rn = int(el.get("r"))
            if rn == 1:
                header_el = el
            else:
                row_map[rn] = el
        for el in row_els:
            sheet_data.remove(el)
        if header_el is not None:
            sheet_data.append(header_el)
        for i, orig_row_num in enumerate(row_nums):
            el = row_map.get(orig_row_num)
            if el is None:
                continue
            new_row_num = i + 2
            el.set("r", str(new_row_num))
            for c in el.findall(_local("c")):
                ref = c.get("r")
                if ref:
                    c.set("r", re.sub(r"\d+$", str(new_row_num), ref))
            sheet_data.append(el)
        return etree.tostring(doc, xml_declaration=True, encoding="UTF-8", standalone=True)

    first = sheet_groups[0]
    first_name = sanitize(first["name"], original_sheet_name)
    target_el.set("name", first_name)
    new_files = {sheet_path: build_filtered_sheet_xml(first["row_nums"])}

    sheet_file_idx = max_sheet_file_idx
    rid_counter = max_rid
    sheet_id_counter = max_sheet_id
    insert_after = target_el
    ct_types_el = ct_root  # <Types> is root

    for g in sheet_groups[1:]:
        sheet_file_idx += 1
        rid_counter += 1
        sheet_id_counter += 1
        new_path = f"xl/worksheets/sheet{sheet_file_idx}.xml"
        new_rid = f"rId{rid_counter}"
        new_name = sanitize(g["name"], f"Sheet{sheet_id_counter}")

        new_files[new_path] = build_filtered_sheet_xml(g["row_nums"])

        new_sheet_el = etree.SubElement(sheets_el, _local("sheet"))
        new_sheet_el.set("name", new_name)
        new_sheet_el.set("sheetId", str(sheet_id_counter))
        new_sheet_el.set(f"{{{R_NS}}}id", new_rid)
        insert_after.addnext(new_sheet_el)
        insert_after = new_sheet_el

        new_rel_el = etree.SubElement(rels_root, f"{{{PKG_REL_NS}}}Relationship")
        new_rel_el.set("Id", new_rid)
        new_rel_el.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
        new_rel_el.set("Target", f"worksheets/sheet{sheet_file_idx}.xml")

        new_override = etree.SubElement(ct_types_el, f"{{{CT_NS}}}Override")
        new_override.set("PartName", f"/xl/worksheets/sheet{sheet_file_idx}.xml")
        new_override.set("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")

    new_files["xl/workbook.xml"] = etree.tostring(wb_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    new_files["xl/_rels/workbook.xml.rels"] = etree.tostring(rels_root, xml_declaration=True, encoding="UTF-8", standalone=True)
    new_files["[Content_Types].xml"] = etree.tostring(ct_root, xml_declaration=True, encoding="UTF-8", standalone=True)

    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename in new_files:
                continue
            zout.writestr(item, zin.read(item.filename))
        for path, content in new_files.items():
            zout.writestr(path, content)
    zin.close()
    out_buf.seek(0)
    return out_buf.getvalue()

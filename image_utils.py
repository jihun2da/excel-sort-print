# -*- coding: utf-8 -*-
"""
정렬된 행 데이터를 A4(300dpi) 인쇄용 이미지로 렌더링하는 유틸.
JS(캔버스) 버전과 동일한 레이아웃 규격을 사용한다.
"""
import glob
from PIL import Image, ImageDraw, ImageFont

from xlsx_utils import COL, text_color_for

PAGE_W, PAGE_H = 2480, 3508
MARGIN, GAP, ROWS = 80, 40, 15
COL_W = (PAGE_W - 2 * MARGIN - GAP) // 2
ROW_H = (PAGE_H - 2 * MARGIN) / ROWS
NO_W, A_W = 130, 130
TOP_H = ROW_H * 0.6
FONT_MAIN = 83   # 번호 / E열: 20pt @300dpi
FONT_A = 46      # A열 값: 11pt @300dpi
FONT_SUB = 46    # F열/G열 기준 크기: 11pt @300dpi (겹치면 자동 축소)


def _find_font_paths():
    bold_candidates = glob.glob("/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc") + \
        glob.glob("/usr/share/fonts/**/NotoSansCJK-Bold.ttc", recursive=True) + \
        glob.glob("/usr/share/fonts/**/*CJK*Bold*.ttc", recursive=True)
    reg_candidates = glob.glob("/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc") + \
        glob.glob("/usr/share/fonts/**/NotoSansCJK-Regular.ttc", recursive=True) + \
        glob.glob("/usr/share/fonts/**/*CJK*Regular*.ttc", recursive=True)
    bold = bold_candidates[0] if bold_candidates else None
    reg = reg_candidates[0] if reg_candidates else None
    return bold, reg


FONT_BOLD_PATH, FONT_REGULAR_PATH = _find_font_paths()
KR_INDEX = 2  # NotoSansCJK ttc 내 한국어(KR) 서브폰트 인덱스

_font_cache = {}


def get_font(bold, size):
    key = (bold, size)
    if key in _font_cache:
        return _font_cache[key]
    path = FONT_BOLD_PATH if bold else FONT_REGULAR_PATH
    try:
        font = ImageFont.truetype(path, size, index=KR_INDEX) if path else ImageFont.load_default()
    except Exception:
        font = ImageFont.load_default()
    _font_cache[key] = font
    return font


def _text_w(draw, text, font):
    bbox = draw.textbbox((0, 0), text, font=font)
    return bbox[2] - bbox[0]


def _fit_font(draw, text, max_width, base_size, bold, min_size=20):
    size = base_size
    font = get_font(bold, size)
    while size > min_size and _text_w(draw, text, font) > max_width:
        size -= 2
        font = get_font(bold, size)
    return font


def _draw_text_centered_v(draw, xy, text, font, fill, align="left"):
    x, y_center = xy
    bbox = draw.textbbox((0, 0), text, font=font)
    h = bbox[3] - bbox[1]
    w = bbox[2] - bbox[0]
    ty = y_center - h / 2 - bbox[1]
    if align == "center":
        draw.text((x - w / 2, ty), text, font=font, fill=fill)
    elif align == "right":
        draw.text((x - w, ty), text, font=font, fill=fill)
    else:
        draw.text((x, ty), text, font=font, fill=fill)


def draw_cell(draw, x0, y0, x1, y1, no, a_val, e_val, f_val, g_val, a_color, divider_style):
    if a_color:
        draw.rectangle([x0 + NO_W, y0, x0 + NO_W + A_W, y1], fill=a_color)

    draw.rectangle([x0, y0, x1, y1], outline="black", width=3)
    draw.line([x0 + NO_W, y0, x0 + NO_W, y1], fill="black", width=2)
    draw.line([x0 + NO_W + A_W, y0, x0 + NO_W + A_W, y1], fill="black", width=2)

    ycen = (y0 + y1) / 2
    if no is not None:
        t = str(no)
        f = _fit_font(draw, t, NO_W - 10, FONT_MAIN, True)
        _draw_text_centered_v(draw, (x0 + NO_W / 2, ycen), t, f, "black", "center")

    if a_val not in (None, ""):
        t = str(a_val)
        f = _fit_font(draw, t, A_W - 10, FONT_A, True)
        _draw_text_centered_v(draw, (x0 + NO_W + A_W / 2, ycen), t, f, text_color_for(a_color), "center")

    main_x0 = x0 + NO_W + A_W
    main_x1 = x1
    y_mid = y0 + TOP_H

    if divider_style:
        draw.line([main_x0, y_mid, main_x1, y_mid], fill=divider_style["color"], width=max(1, int(round(divider_style["width"]))))

    if e_val not in (None, ""):
        t = str(e_val)
        f = _fit_font(draw, t, (main_x1 - main_x0) - 30, FONT_MAIN, True)
        _draw_text_centered_v(draw, (main_x0 + 18, (y0 + y_mid) / 2), t, f, "black", "left")

    f_str = str(f_val) if f_val not in (None, "") else ""
    g_str = str(g_val) if g_val not in (None, "") else ""
    avail_w = (main_x1 - main_x0) - 36
    gap_min = 16
    sub_size = FONT_SUB
    while True:
        font = get_font(False, sub_size)
        fw = _text_w(draw, f_str, font) if f_str else 0
        gw = _text_w(draw, g_str, font) if g_str else 0
        gap = gap_min if (f_str and g_str) else 0
        if fw + gw + gap <= avail_w or sub_size <= 14:
            break
        sub_size -= 1
    font = get_font(False, sub_size)
    if f_str:
        _draw_text_centered_v(draw, (main_x0 + 18, (y_mid + y1) / 2), f_str, font, "black", "left")
    if g_str:
        _draw_text_centered_v(draw, (main_x1 - 18, (y_mid + y1) / 2), g_str, font, "black", "right")


def build_page_image(left_chunk, right_chunk, left_start, right_start, divider_style):
    img = Image.new("RGB", (PAGE_W, PAGE_H), "white")
    draw = ImageDraw.Draw(img)
    for col, (chunk, start) in enumerate([(left_chunk, left_start), (right_chunk, right_start)]):
        x0 = MARGIN + col * (COL_W + GAP)
        x1 = x0 + COL_W
        for i in range(ROWS):
            y0 = MARGIN + i * ROW_H
            y1 = y0 + ROW_H
            if chunk and i < len(chunk):
                row = chunk[i]
                cells = row["cells"]
                draw_cell(draw, x0, y0, x1, y1, start + i,
                          cells[COL["A"]], cells[COL["E"]], cells[COL["F"]], cells[COL["G"]],
                          row["a_color"], divider_style)
            else:
                draw.rectangle([x0, y0, x1, y1], outline="black", width=3)
    return img


def build_pages_for_rows(rows, divider_style):
    """행 리스트 하나(=시트/그룹 하나)를 A4 페이지 이미지 리스트로 변환. 번호는 항상 1번부터."""
    chunk_size = 15
    chunks = [rows[i:i + chunk_size] for i in range(0, len(rows), chunk_size)]
    num_chunks = len(chunks)
    num_pages = -(-num_chunks // 2)  # ceil
    pages = []
    for p in range(num_pages):
        left_chunk = chunks[p] if p < num_chunks else []
        right_idx = p + num_pages
        right_chunk = chunks[right_idx] if right_idx < num_chunks else []
        left_start = p * 15 + 1
        right_start = right_idx * 15 + 1
        img = build_page_image(left_chunk, right_chunk, left_start, right_start, divider_style)
        if right_chunk:
            range_text = f"{left_start}-{left_start+len(left_chunk)-1}, {right_start}-{right_start+len(right_chunk)-1}"
        else:
            range_text = f"{left_start}-{left_start+len(left_chunk)-1}"
        pages.append({"image": img, "page_num": p + 1, "range_text": range_text})
    return pages

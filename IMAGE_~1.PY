# -*- coding: utf-8 -*-
"""
정렬된 행 데이터를 A4 인쇄용 이미지로 렌더링하는 유틸.

- 해상도는 150dpi(2480x3508의 절반)로 낮춰 이미지 1장당 메모리 사용량을 약 1/4로 줄였다.
  (Streamlit Community Cloud 무료 서버의 메모리 한도 초과로 인한 "Oh no" 크래시를 줄이기 위함)
- 페이지는 생성 직후 PNG bytes 로 변환해 PIL Image 객체를 즉시 버린다(메모리 절약).
- "표지(0번)" 셀: 이름 + 배경색을 지정하면 첫 페이지 맨 앞에 표지 칸이 추가된다.
  기존 데이터는 순서·번호가 전혀 바뀌지 않고(밀려나지 않고) 그대로 1번부터 이어진다.
- build_combined_pdf: 생성된 모든 페이지를 순서대로 하나의 PDF로 합친다(인쇄용).
"""
import glob
import io
from PIL import Image, ImageDraw, ImageFont

from xlsx_utils import COL, text_color_for

# ---------- 레이아웃 상수 (150dpi 기준, 300dpi 대비 절반 크기로 메모리 절감) ----------
PAGE_W, PAGE_H = 1240, 1754
MARGIN, GAP, ROWS = 40, 20, 15
COL_W = (PAGE_W - 2 * MARGIN - GAP) // 2
ROW_H = (PAGE_H - 2 * MARGIN) / ROWS
NO_W, A_W = 65, 65
TOP_H = ROW_H * 0.6
FONT_MAIN = 42   # E열: 20pt @150dpi
FONT_NO = round(FONT_MAIN * 0.5)  # 번호(No.) 칸 전용: 요청에 따라 기존 크기의 50%
FONT_A = 23      # A열 값: 11pt @150dpi
FONT_SUB = 23    # F열/G열 기준 크기: 11pt @150dpi (겹치면 자동 축소)
FONT_COVER = 50  # 표지 이름 기준 크기


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
    size = max(int(size), 8)
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


def _fit_font(draw, text, max_width, base_size, bold, min_size=12):
    size = base_size
    font = get_font(bold, size)
    while size > min_size and _text_w(draw, text, font) > max_width:
        size -= 1
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

    draw.rectangle([x0, y0, x1, y1], outline="black", width=2)
    draw.line([x0 + NO_W, y0, x0 + NO_W, y1], fill="black", width=1)
    draw.line([x0 + NO_W + A_W, y0, x0 + NO_W + A_W, y1], fill="black", width=1)

    ycen = (y0 + y1) / 2
    if no is not None:
        t = str(no)
        f = _fit_font(draw, t, NO_W - 8, FONT_NO, True)
        _draw_text_centered_v(draw, (x0 + NO_W / 2, ycen), t, f, "black", "center")

    if a_val not in (None, ""):
        t = str(a_val)
        f = _fit_font(draw, t, A_W - 8, FONT_A, True)
        _draw_text_centered_v(draw, (x0 + NO_W + A_W / 2, ycen), t, f, text_color_for(a_color), "center")

    main_x0 = x0 + NO_W + A_W
    main_x1 = x1
    y_mid = y0 + TOP_H

    if divider_style:
        draw.line([main_x0, y_mid, main_x1, y_mid], fill=divider_style["color"], width=max(1, int(round(divider_style["width"] * 0.6))))

    if e_val not in (None, ""):
        t = str(e_val)
        f = _fit_font(draw, t, (main_x1 - main_x0) - 20, FONT_MAIN, True)
        _draw_text_centered_v(draw, (main_x0 + 10, (y0 + y_mid) / 2), t, f, "black", "left")

    f_str = str(f_val) if f_val not in (None, "") else ""
    g_str = str(g_val) if g_val not in (None, "") else ""
    avail_w = (main_x1 - main_x0) - 20
    gap_min = 10
    sub_size = FONT_SUB
    while True:
        font = get_font(False, sub_size)
        fw = _text_w(draw, f_str, font) if f_str else 0
        gw = _text_w(draw, g_str, font) if g_str else 0
        gap = gap_min if (f_str and g_str) else 0
        if fw + gw + gap <= avail_w or sub_size <= 8:
            break
        sub_size -= 1
    font = get_font(False, sub_size)
    if f_str:
        _draw_text_centered_v(draw, (main_x0 + 10, (y_mid + y1) / 2), f_str, font, "black", "left")
    if g_str:
        _draw_text_centered_v(draw, (main_x1 - 10, (y_mid + y1) / 2), g_str, font, "black", "right")


def draw_cover_cell(draw, x0, y0, x1, y1, name, color):
    bg = color or "#FFD54F"
    draw.rectangle([x0, y0, x1, y1], fill=bg, outline="black", width=2)
    tc = text_color_for(bg)

    # 좌상단에 작게 "0" 표시
    f_no = get_font(True, FONT_A)
    draw.text((x0 + 6, y0 + 4), "0", font=f_no, fill=tc)

    # 중앙에 크게 이름 표시
    t = name or ""
    if t:
        f = _fit_font(draw, t, (x1 - x0) - 16, FONT_COVER, True, min_size=14)
        _draw_text_centered_v(draw, ((x0 + x1) / 2, (y0 + y1) / 2), t, f, tc, "center")


def build_page_image(left_chunk, right_chunk, divider_style):
    img = Image.new("RGB", (PAGE_W, PAGE_H), "white")
    draw = ImageDraw.Draw(img)
    for col, chunk in enumerate([left_chunk, right_chunk]):
        x0 = MARGIN + col * (COL_W + GAP)
        x1 = x0 + COL_W
        for i in range(ROWS):
            y0 = MARGIN + i * ROW_H
            y1 = y0 + ROW_H
            if chunk and i < len(chunk):
                it = chunk[i]
                if it["type"] == "cover":
                    draw_cover_cell(draw, x0, y0, x1, y1, it["name"], it["color"])
                else:
                    row = it["row"]
                    cells = row["cells"]
                    draw_cell(draw, x0, y0, x1, y1, it["no"],
                              cells[COL["A"]], cells[COL["E"]], cells[COL["F"]], cells[COL["G"]],
                              row["a_color"], divider_style)
            else:
                draw.rectangle([x0, y0, x1, y1], outline="black", width=2)
    return img


def _label_range(chunk):
    nos = [it["no"] for it in chunk if it["type"] == "data"]
    has_cover = any(it["type"] == "cover" for it in chunk)
    if not nos:
        return "표지" if has_cover else ""
    r = f"{nos[0]}-{nos[-1]}"
    return ("표지," + r) if has_cover else r


def build_pages_for_rows(rows, divider_style, cover=None):
    """행 리스트 하나(=시트/그룹 하나)를 A4 페이지 이미지(PNG bytes) 리스트로 변환.

    cover: {"name": str, "color": "#RRGGBB"} 를 넘기면 첫 페이지 왼쪽 맨 위 칸 앞에
    "0번" 표지 칸이 추가된다. 기존 데이터는 순서나 번호가 전혀 바뀌지 않고(밀려나거나
    이동하지 않고) 그대로 1번부터 번호가 매겨진다. 표지가 추가되는 만큼 전체 칸 수가
    1개 늘어나므로 페이지 수가 늘어날 수 있다.
    """
    if not rows:
        return []

    if cover and (cover.get("name") or "").strip():
        items = [{"type": "cover", "name": cover.get("name", "").strip(), "color": cover.get("color") or "#FFD54F"}]
        items += [{"type": "data", "row": r, "no": i + 1} for i, r in enumerate(rows)]
    else:
        items = [{"type": "data", "row": r, "no": i + 1} for i, r in enumerate(rows)]

    chunk_size = 15
    chunks = [items[i:i + chunk_size] for i in range(0, len(items), chunk_size)]
    num_chunks = len(chunks)
    num_pages = -(-num_chunks // 2)  # ceil
    pages = []
    for p in range(num_pages):
        left_chunk = chunks[p] if p < num_chunks else []
        right_idx = p + num_pages
        right_chunk = chunks[right_idx] if right_idx < num_chunks else []
        img = build_page_image(left_chunk, right_chunk, divider_style)
        buf = io.BytesIO()
        img.save(buf, format="PNG", optimize=True)
        png_bytes = buf.getvalue()
        img.close()
        del img

        left_label = _label_range(left_chunk)
        right_label = _label_range(right_chunk)
        range_text = f"{left_label}, {right_label}" if right_chunk else left_label

        pages.append({"png_bytes": png_bytes, "page_num": p + 1, "range_text": range_text})
    return pages


def build_combined_pdf(all_pages):
    """여러 페이지(dict, "png_bytes" 포함)를 순서대로 이어붙인 PDF 하나의 bytes 를 만든다."""
    if not all_pages:
        return None
    imgs = []
    try:
        for pg in all_pages:
            im = Image.open(io.BytesIO(pg["png_bytes"]))
            imgs.append(im.convert("RGB"))
        buf = io.BytesIO()
        imgs[0].save(buf, format="PDF", save_all=True, append_images=imgs[1:])
        return buf.getvalue()
    finally:
        for im in imgs:
            try:
                im.close()
            except Exception:
                pass

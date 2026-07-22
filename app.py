# -*- coding: utf-8 -*-
"""
엑셀 정렬 & A4 인쇄용 이미지 생성 웹앱 (Streamlit 버전)

- 대량 분류 (이미지 전용): 엑셀 정렬은 하지 않고 업로드된 원본 순서 그대로, H열 배경색이 이어지는
  구간별로 그룹을 나눈다(중간에 다른 색이 1~2행만 섞이면 무시, 3행 이상이면 새 그룹). 그룹마다
  별도의 이미지 세트(1번부터 번호)를 생성하며, 엑셀 다운로드 단계 없이 바로 이미지 생성으로 이동한다.
- 소량 분류: F열 -> G열 -> A열 -> K열 순서로 순차 정렬 후, H열 값이 같은 행이
  기준 개수(기본 50, 40/30 선택 가능) 이상이면 그 그룹만 별도 시트로 분리한다.
  정렬된 엑셀 다운로드 + A4 인쇄용 이미지 생성을 모두 제공한다.
- 정렬된 엑셀 다운로드는 원본 파일의 행(XML row) 자체를 재배치하는 방식이라
  셀 서식·배경색·필터·열 너비 등이 손실되지 않는다.
- A4 인쇄용 이미지 생성 전, "표지(0번)" 이름/색상을 입력하면 해당 그룹(시트)의 첫 페이지 첫 칸이
  표지로 바뀌고 실제 데이터는 그대로 1번부터 이어진다(원래 1번 자리였던 데이터 1건은 맨 뒤로 밀림).
  대량 분류는 감지된 색상 그룹 개수만큼 표지 입력칸이 각각 표시된다.
- 생성된 이미지는 순서대로 하나의 PDF로 합쳐지며, 앱 안에서 "인쇄" 버튼으로 바로 인쇄할 수 있다.
- 메모리 사용량을 줄이기 위해 이미지는 150dpi로 생성하고, PIL Image 대신 PNG bytes 로만 보관한다.
- 번호(No.) 칸의 숫자 크기는 E열 값 폰트의 50% 크기로 표시된다.
"""
import base64
import gc
import io
import zipfile

import streamlit as st
import streamlit.components.v1 as components

from xlsx_utils import load_sheet_rows, sort_sequential, build_multi_sheet_xlsx, group_rows_by_h_color, COL
from image_utils import build_pages_for_rows, build_combined_pdf, FONT_BOLD_PATH, FONT_REGULAR_PATH

st.set_page_config(page_title="엑셀 정렬 & A4 인쇄용 이미지 생성", layout="wide")

if not FONT_BOLD_PATH or not FONT_REGULAR_PATH:
    st.warning(
        "한글 폰트(Noto Sans CJK)를 서버에서 찾지 못했습니다. 이미지의 한글이 깨져 보일 수 있습니다. "
        "배포 환경에 packages.txt(fonts-noto-cjk)가 함께 설치되었는지 확인해주세요."
    )

st.title("엑셀 정렬 & A4 인쇄용 이미지 생성 웹앱")
st.caption(
    "엑셀 업로드 → 분류 방식 선택 → (소량 분류만) 정렬된 엑셀 다운로드 → "
    "표지(선택) 입력 → A4 인쇄용 이미지 생성 → 인쇄"
)

# ---------- 세션 상태 초기화 ----------
_defaults = {
    "raw_bytes": None,
    "sheet_names": [],
    "current_sheet": None,
    "loaded": None,
    "sorted_sheets": [],
    "method_label": "",
    "mode": None,  # "bulk_image_only" | "small_batch"
    "pages_by_sheet": {},
    "combined_pdf_bytes": None,
}
for k, v in _defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


def _reset_downstream():
    st.session_state.sorted_sheets = []
    st.session_state.pages_by_sheet = {}
    st.session_state.combined_pdf_bytes = None


# ---------- 1. 파일 업로드 ----------
st.header("1. 엑셀 파일 업로드")
uploaded = st.file_uploader("엑셀 파일(.xlsx)을 선택하세요", type=["xlsx"])

if uploaded is not None:
    raw_bytes = uploaded.getvalue()
    if raw_bytes != st.session_state.raw_bytes:
        try:
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), read_only=True)
            sheet_names = wb.sheetnames
            first_loaded = load_sheet_rows(raw_bytes, sheet_names[0])
        except Exception as e:
            st.error(
                f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}\n\n"
                "파일이 손상되지 않았는지, 정상적인 .xlsx 파일인지 확인해주세요."
            )
        else:
            st.session_state.raw_bytes = raw_bytes
            st.session_state.sheet_names = sheet_names
            st.session_state.current_sheet = sheet_names[0]
            st.session_state.loaded = first_loaded
            st.session_state.mode = None
            _reset_downstream()

if st.session_state.raw_bytes is not None:
    sheet = st.selectbox(
        "시트 선택", st.session_state.sheet_names,
        index=st.session_state.sheet_names.index(st.session_state.current_sheet),
    )
    if sheet != st.session_state.current_sheet:
        try:
            loaded = load_sheet_rows(st.session_state.raw_bytes, sheet)
        except Exception as e:
            st.error(f"시트를 읽는 중 오류가 발생했습니다: {e}")
        else:
            st.session_state.current_sheet = sheet
            st.session_state.loaded = loaded
            st.session_state.mode = None
            _reset_downstream()

    n_rows = len(st.session_state.loaded.rows) if st.session_state.loaded else 0
    st.info(f'시트 "{st.session_state.current_sheet}" 불러옴 · 데이터 {n_rows}행 (헤더 제외)')
    if n_rows > 2000:
        st.warning(
            f"데이터가 {n_rows}행으로 많은 편입니다. 이미지 생성 시 서버 메모리를 많이 사용해 "
            "느려지거나 실패할 수 있습니다. 필요하다면 시트를 나눠서 처리해주세요."
        )

    # ---------- 2. 분류 방식 선택 ----------
    st.header("2. 분류(정렬) 방식 선택")
    col_bulk, col_small = st.columns(2)

    with col_bulk:
        st.subheader("대량 분류 (이미지 전용)")
        st.caption(
            "엑셀 정렬은 하지 않고 업로드된 원본 순서 그대로, H열 배경색이 이어지는 구간별로 "
            "그룹을 나눠 그룹마다 별도의 이미지 세트(1번부터 번호)를 생성합니다. 중간에 다른 색이 "
            "1~2행만 섞이면 무시하고 같은 그룹으로 처리하며, 3행 이상 이어지면 새 그룹으로 전환합니다."
        )
        if st.button("대량 분류 실행", type="primary", use_container_width=True):
            try:
                groups = group_rows_by_h_color(st.session_state.loaded.rows)
            except Exception as e:
                st.error(f"H열 색상 그룹을 나누는 중 오류가 발생했습니다: {e}")
            else:
                st.session_state.sorted_sheets = [
                    {"name": f"그룹{i + 1}", "rows": g["rows"], "group_color": g["color"]}
                    for i, g in enumerate(groups)
                ]
                st.session_state.method_label = "대량분류"
                st.session_state.mode = "bulk_image_only"
                st.session_state.pages_by_sheet = {}
                st.session_state.combined_pdf_bytes = None

    with col_small:
        st.subheader("소량 분류 (엑셀 + 이미지)")
        st.caption(
            "F열 정렬 → G열 정렬 → A열 정렬 → K열 정렬 순서로 차례대로 적용 후, "
            "H열 값이 같은 행이 기준 개수 이상이면 해당 그룹만 별도 시트로 분리합니다. "
            "정렬된 엑셀 다운로드와 이미지 생성을 모두 제공합니다."
        )
        threshold = st.selectbox("H열 동일값 그룹 분리 기준", [50, 40, 30], index=0, format_func=lambda v: f"{v}개 이상")
        if st.button("소량 분류 실행", use_container_width=True):
            try:
                sorted_rows = sort_sequential(st.session_state.loaded.rows, ["F", "G", "A", "K"])
                group_map = {}
                for row in sorted_rows:
                    hv = row["cells"][COL["H"]]
                    key = "" if hv is None else str(hv)
                    group_map.setdefault(key, []).append(row)
                large_groups = [(k, v) for k, v in group_map.items() if len(v) >= threshold]
                large_keys = {k for k, _ in large_groups}
                main_rows = [
                    row for row in sorted_rows
                    if ("" if row["cells"][COL["H"]] is None else str(row["cells"][COL["H"]])) not in large_keys
                ]
                sheets = []
                if main_rows:
                    sheets.append({"name": f"{st.session_state.current_sheet}_기본", "rows": main_rows})
                for k, v in large_groups:
                    sheets.append({"name": k or "(빈값)", "rows": v})
                if not sheets:
                    sheets = [{"name": st.session_state.current_sheet, "rows": sorted_rows}]
            except Exception as e:
                st.error(f"분류 중 오류가 발생했습니다: {e}")
            else:
                st.session_state.sorted_sheets = sheets
                st.session_state.method_label = "소량분류"
                st.session_state.mode = "small_batch"
                st.session_state.pages_by_sheet = {}
                st.session_state.combined_pdf_bytes = None

    # ---------- 3. 결과 ----------
    if st.session_state.sorted_sheets:
        st.header("3. 결과")
        summary_lines = []
        for s in st.session_state.sorted_sheets:
            extra = f" · H열 색상 {s['group_color']}" if s.get("group_color") else (" · H열 색상 없음" if "group_color" in s else "")
            summary_lines.append(f"- **{s['name']}**: {len(s['rows'])}개{extra}")
        st.markdown(
            f"**{st.session_state.method_label} 완료** · 시트 {len(st.session_state.sorted_sheets)}개 생성\n\n"
            + "\n".join(summary_lines)
        )

        if st.session_state.mode == "small_batch":
            col_dl, col_div = st.columns([1, 1])
            with col_dl:
                try:
                    sheet_groups = [
                        {"name": s["name"], "row_nums": [r["row_num"] for r in s["rows"]]}
                        for s in st.session_state.sorted_sheets
                    ]
                    out_bytes = build_multi_sheet_xlsx(
                        st.session_state.raw_bytes, st.session_state.current_sheet, sheet_groups
                    )
                except Exception as e:
                    st.error(f"정렬된 엑셀을 만드는 중 오류가 발생했습니다: {e}")
                else:
                    st.download_button(
                        "정렬된 엑셀 다운로드",
                        data=out_bytes,
                        file_name=f"정렬_{st.session_state.method_label}_{st.session_state.current_sheet}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                    st.caption(
                        "원본 파일의 행(XML row) 자체를 그대로 재배치/재분류하는 방식이라 "
                        "셀 서식·배경색·필터·열 너비 등 원본 정보가 전혀 손실되지 않습니다."
                    )
        else:
            col_div = st.container()

        with col_div:
            divider_choice = st.selectbox(
                "E열 값과 F·G열 값 사이 구분선",
                ["연하게", "진하게(기본)", "표시 안 함"],
                index=0,
            )
            if divider_choice == "표시 안 함":
                divider_style = None
            elif divider_choice == "연하게":
                divider_style = {"color": "#d4d4d4", "width": 1.5}
            else:
                divider_style = {"color": "#000000", "width": 2}

        cover_by_index = {}
        if st.session_state.mode == "bulk_image_only":
            st.subheader("표지(0번) 설정 - 그룹별 (선택 사항)")
            st.caption(
                "그룹마다 이름을 입력하면 해당 그룹의 첫 페이지 첫 칸이 '0번' 표지 칸으로 바뀌고, 원래 그 "
                "자리에 있던 데이터 1건은 맨 뒤로 밀려서 번호가 다시 매겨집니다. 비워두면 그 그룹은 표지 "
                "없이 기존처럼 1번부터 생성됩니다."
            )
            for i, s in enumerate(st.session_state.sorted_sheets):
                default_color = s.get("group_color") or "#FFD54F"
                col_cn, col_cc = st.columns([2, 1])
                with col_cn:
                    name = st.text_input(
                        f"{s['name']} 표지 이름 ({len(s['rows'])}개)",
                        key=f"cover_name_{i}", placeholder="예: 홍길동",
                    )
                with col_cc:
                    color = st.color_picker(f"{s['name']} 표지 배경색", default_color, key=f"cover_color_{i}")
                cover_by_index[i] = {"name": name, "color": color} if name and name.strip() else None
        else:
            st.subheader("표지(0번) 설정 (선택 사항)")
            st.caption(
                "이름을 입력하면 첫 페이지 첫 칸이 '0번' 표지 칸으로 바뀌고, 원래 그 자리에 있던 데이터 1건은 "
                "맨 뒤로 밀려서 번호가 다시 매겨집니다. 비워두면 표지 없이 기존처럼 1번부터 생성됩니다."
            )
            col_cn, col_cc = st.columns([2, 1])
            with col_cn:
                cover_name = st.text_input("표지에 표시할 이름", key="cover_name_input", placeholder="예: 홍길동")
            with col_cc:
                cover_color = st.color_picker("표지 배경색", "#FFD54F", key="cover_color_input")
            cover_by_index[0] = {"name": cover_name, "color": cover_color} if cover_name and cover_name.strip() else None

        if st.button("A4 인쇄용 이미지 생성", type="primary"):
            pages_by_sheet = {}
            progress = st.progress(0.0, text="이미지 생성 준비 중...")

            def _num_pages(n_items):
                num_chunks = -(-n_items // 15)  # ceil(n_items / 15)
                return -(-num_chunks // 2)       # ceil(num_chunks / 2)

            try:
                total = sum(
                    _num_pages(len(s["rows"]) + (1 if (cover_by_index.get(i) and s["rows"]) else 0))
                    for i, s in enumerate(st.session_state.sorted_sheets)
                ) or 1
                done = 0
                for i, s in enumerate(st.session_state.sorted_sheets):
                    if not s["rows"]:
                        continue
                    this_cover = cover_by_index.get(i)
                    pages = build_pages_for_rows(s["rows"], divider_style, cover=this_cover)
                    pages_by_sheet[s["name"]] = pages
                    done += len(pages)
                    progress.progress(min(1.0, done / max(total, 1)), text=f"이미지 생성 중... ({done}/{total} 페이지)")

                st.session_state.pages_by_sheet = pages_by_sheet

                all_pages_flat = [pg for pages in pages_by_sheet.values() for pg in pages]
                try:
                    st.session_state.combined_pdf_bytes = build_combined_pdf(all_pages_flat)
                except Exception as e:
                    st.session_state.combined_pdf_bytes = None
                    st.warning(f"PDF 합치기에 실패해 인쇄/PDF 다운로드는 사용할 수 없습니다: {e}")
            except MemoryError:
                st.error(
                    "이미지를 생성하는 중 서버 메모리가 부족했습니다. 데이터 양을 줄이거나, "
                    "시트를 나눠서 다시 시도해주세요."
                )
                st.session_state.pages_by_sheet = {}
            except Exception as e:
                st.error(f"이미지 생성 중 오류가 발생했습니다: {e}")
                st.session_state.pages_by_sheet = {}
            finally:
                progress.empty()
                gc.collect()

            if st.session_state.pages_by_sheet:
                st.success(
                    f"완료: 시트 {len(st.session_state.pages_by_sheet)}개, "
                    f"총 {sum(len(p) for p in st.session_state.pages_by_sheet.values())}페이지 생성됨"
                )

        st.info(
            "인쇄 시 프린터 설정에서 **실제 크기(100%) / 배율 없음**으로 인쇄해야 A4 용지에 정확히 맞습니다. "
            "(여백 채우기·자동맞춤 설정은 끄고 인쇄하세요)"
        )

        if st.session_state.pages_by_sheet:
            # ---------- 인쇄 / PDF ----------
            if st.session_state.combined_pdf_bytes:
                st.subheader("인쇄")
                b64 = base64.b64encode(st.session_state.combined_pdf_bytes).decode()
                html_template = """
<div style="display:flex;gap:12px;align-items:center;flex-wrap:wrap;">
  <button id="printBtn" style="padding:10px 18px;font-size:15px;background:#ff4b4b;color:white;border:none;border-radius:6px;cursor:pointer;">
    인쇄 (생성된 순서대로)
  </button>
  <span style="color:#666;font-size:13px;">클릭하면 새 탭이 열리고 인쇄 대화상자가 자동으로 뜹니다. 팝업이 차단되면 브라우저 팝업 차단을 해제해주세요.</span>
</div>
<script>
const B64_DATA = "__B64__";
function b64ToBlob(b64, mime) {
  const byteChars = atob(b64);
  const byteNumbers = new Array(byteChars.length);
  for (let i = 0; i < byteChars.length; i++) { byteNumbers[i] = byteChars.charCodeAt(i); }
  const byteArray = new Uint8Array(byteNumbers);
  return new Blob([byteArray], {type: mime});
}
document.getElementById('printBtn').addEventListener('click', function() {
  const blob = b64ToBlob(B64_DATA, 'application/pdf');
  const url = URL.createObjectURL(blob);
  const win = window.open(url, '_blank');
  if (win) {
    win.onload = function(){ try { win.print(); } catch(e) {} };
    setTimeout(function(){ try { win.print(); } catch(e) {} }, 800);
  } else {
    alert('팝업이 차단되었습니다. 브라우저의 팝업 차단을 해제한 뒤 다시 시도해주세요.');
  }
});
</script>
"""
                components.html(html_template.replace("__B64__", b64), height=70)
                st.download_button(
                    "PDF 다운로드",
                    data=st.session_state.combined_pdf_bytes,
                    file_name="인쇄용_전체.pdf",
                    mime="application/pdf",
                )

            # ---------- 전체 ZIP 다운로드 ----------
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for sheet_name, pages in st.session_state.pages_by_sheet.items():
                    safe_name = "".join(c if c not in '\\/?*[]:' else "_" for c in sheet_name)
                    for pg in pages:
                        zf.writestr(f"{safe_name}/인쇄_{pg['page_num']:02d}페이지.png", pg["png_bytes"])
            st.download_button(
                "전체 이미지 ZIP 다운로드",
                data=zip_buf.getvalue(),
                file_name="인쇄용_이미지_전체.zip",
                mime="application/zip",
            )

            for sheet_name, pages in st.session_state.pages_by_sheet.items():
                st.subheader(f"시트: {sheet_name} ({len(pages)}페이지)")
                cols = st.columns(4)
                for idx, pg in enumerate(pages):
                    with cols[idx % 4]:
                        st.image(pg["png_bytes"], caption=f"{pg['page_num']}페이지 ({pg['range_text']})")
                        st.download_button(
                            "다운로드",
                            data=pg["png_bytes"],
                            file_name=f"{sheet_name}_인쇄_{pg['page_num']}페이지.png",
                            mime="image/png",
                            key=f"dl_{sheet_name}_{pg['page_num']}",
                        )
else:
    st.info("먼저 엑셀 파일을 업로드해주세요.")

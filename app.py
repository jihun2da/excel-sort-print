# -*- coding: utf-8 -*-
"""
엑셀 정렬 & A4 인쇄용 이미지 생성 웹앱 (Streamlit 버전)

- 대량 분류: A열 -> I열 -> J열 -> K열 순서로 순차 정렬
- 소량 분류: F열 -> G열 -> A열 -> K열 순서로 순차 정렬 후, H열 값이 같은 행이
  기준 개수(기본 50, 40/30 선택 가능) 이상이면 그 그룹만 별도 시트로 분리
- 정렬된 엑셀 다운로드는 원본 파일의 행(XML row) 자체를 재배치하는 방식이라
  셀 서식·배경색·필터·열 너비 등이 손실되지 않음
- A4 인쇄용 이미지는 대량/소량 분류 모두 동일한 로직으로 생성 (시트별 15개씩 좌우 2단 배치)
"""
import io
import zipfile

import streamlit as st

from xlsx_utils import load_sheet_rows, sort_sequential, build_multi_sheet_xlsx, COL
from image_utils import build_pages_for_rows, FONT_BOLD_PATH, FONT_REGULAR_PATH

st.set_page_config(page_title="엑셀 정렬 & A4 인쇄용 이미지 생성", layout="wide")

if not FONT_BOLD_PATH or not FONT_REGULAR_PATH:
    st.warning(
        "한글 폰트(Noto Sans CJK)를 서버에서 찾지 못했습니다. 이미지의 한글이 깨져 보일 수 있습니다. "
        "배포 환경에 packages.txt(fonts-noto-cjk)가 함께 설치되었는지 확인해주세요."
    )

st.title("엑셀 정렬 & A4 인쇄용 이미지 생성 웹앱")
st.caption(
    "엑셀 업로드 → 분류 방식 선택 → 정렬된 엑셀 다운로드(서식·색상·필터 등 원본 그대로 유지) "
    "+ A4 인쇄용 이미지 생성"
)

# ---------- 세션 상태 초기화 ----------
if "raw_bytes" not in st.session_state:
    st.session_state.raw_bytes = None
if "sheet_names" not in st.session_state:
    st.session_state.sheet_names = []
if "current_sheet" not in st.session_state:
    st.session_state.current_sheet = None
if "loaded" not in st.session_state:
    st.session_state.loaded = None
if "sorted_sheets" not in st.session_state:
    st.session_state.sorted_sheets = []
if "method_label" not in st.session_state:
    st.session_state.method_label = ""
if "pages_by_sheet" not in st.session_state:
    st.session_state.pages_by_sheet = {}

# ---------- 1. 파일 업로드 ----------
st.header("1. 엑셀 파일 업로드")
uploaded = st.file_uploader("엑셀 파일(.xlsx)을 선택하세요", type=["xlsx"])

if uploaded is not None:
    raw_bytes = uploaded.getvalue()
    if raw_bytes != st.session_state.raw_bytes:
        st.session_state.raw_bytes = raw_bytes
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), read_only=True)
        st.session_state.sheet_names = wb.sheetnames
        st.session_state.current_sheet = wb.sheetnames[0]
        st.session_state.loaded = load_sheet_rows(raw_bytes, wb.sheetnames[0])
        st.session_state.sorted_sheets = []
        st.session_state.pages_by_sheet = {}

if st.session_state.raw_bytes is not None:
    sheet = st.selectbox(
        "시트 선택", st.session_state.sheet_names,
        index=st.session_state.sheet_names.index(st.session_state.current_sheet),
    )
    if sheet != st.session_state.current_sheet:
        st.session_state.current_sheet = sheet
        st.session_state.loaded = load_sheet_rows(st.session_state.raw_bytes, sheet)
        st.session_state.sorted_sheets = []
        st.session_state.pages_by_sheet = {}

    st.info(f'시트 "{sheet}" 불러옴 · 데이터 {len(st.session_state.loaded.rows)}행 (헤더 제외)')

    # ---------- 2. 분류 방식 선택 ----------
    st.header("2. 분류(정렬) 방식 선택")
    col_bulk, col_small = st.columns(2)

    with col_bulk:
        st.subheader("대량 분류 방식")
        st.caption("A열 정렬 → I열 정렬 → J열 정렬 → K열 정렬 순서로 차례대로(순차) 적용")
        if st.button("대량 분류 실행", type="primary", use_container_width=True):
            rows = sort_sequential(st.session_state.loaded.rows, ["A", "I", "J", "K"])
            st.session_state.sorted_sheets = [{"name": sheet, "rows": rows}]
            st.session_state.method_label = "대량분류"
            st.session_state.pages_by_sheet = {}

    with col_small:
        st.subheader("소량 분류 방식")
        st.caption(
            "F열 정렬 → G열 정렬 → A열 정렬 → K열 정렬 순서로 차례대로 적용 후, "
            "H열 값이 같은 행이 기준 개수 이상이면 해당 그룹만 별도 시트로 분리"
        )
        threshold = st.selectbox("H열 동일값 그룹 분리 기준", [50, 40, 30], index=0, format_func=lambda v: f"{v}개 이상")
        if st.button("소량 분류 실행", use_container_width=True):
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
                sheets.append({"name": f"{sheet}_기본", "rows": main_rows})
            for k, v in large_groups:
                sheets.append({"name": k or "(빈값)", "rows": v})
            if not sheets:
                sheets = [{"name": sheet, "rows": sorted_rows}]
            st.session_state.sorted_sheets = sheets
            st.session_state.method_label = "소량분류"
            st.session_state.pages_by_sheet = {}

    # ---------- 3. 결과 ----------
    if st.session_state.sorted_sheets:
        st.header("3. 결과")
        summary_lines = [f"- **{s['name']}**: {len(s['rows'])}개" for s in st.session_state.sorted_sheets]
        st.markdown(
            f"**{st.session_state.method_label} 완료** · 시트 {len(st.session_state.sorted_sheets)}개 생성\n\n"
            + "\n".join(summary_lines)
        )

        col_dl, col_div = st.columns([1, 1])
        with col_dl:
            sheet_groups = [
                {"name": s["name"], "row_nums": [r["row_num"] for r in s["rows"]]}
                for s in st.session_state.sorted_sheets
            ]
            out_bytes = build_multi_sheet_xlsx(
                st.session_state.raw_bytes, st.session_state.current_sheet, sheet_groups
            )
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

        if st.button("A4 인쇄용 이미지 생성"):
            pages_by_sheet = {}
            progress = st.progress(0.0, text="이미지 생성 준비 중...")

            def _num_pages(n_rows):
                num_chunks = -(-n_rows // 15)  # ceil(n_rows / 15)
                return -(-num_chunks // 2)     # ceil(num_chunks / 2)

            total = sum(_num_pages(len(s["rows"])) for s in st.session_state.sorted_sheets) or 1
            done = 0
            for s in st.session_state.sorted_sheets:
                if not s["rows"]:
                    continue
                pages = build_pages_for_rows(s["rows"], divider_style)
                pages_by_sheet[s["name"]] = pages
                done += len(pages)
                progress.progress(min(1.0, done / max(total, 1)), text=f"이미지 생성 중... ({done}/{total} 페이지)")
            progress.empty()
            st.session_state.pages_by_sheet = pages_by_sheet
            st.success(f"완료: 시트 {len(pages_by_sheet)}개, 총 {sum(len(p) for p in pages_by_sheet.values())}페이지 생성됨")

        st.info(
            "인쇄 시 프린터 설정에서 **실제 크기(100%) / 배율 없음**으로 인쇄해야 A4 용지에 정확히 맞습니다. "
            "(여백 채우기·자동맞춤 설정은 끄고 인쇄하세요)"
        )

        if st.session_state.pages_by_sheet:
            # 전체 ZIP 다운로드
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for sheet_name, pages in st.session_state.pages_by_sheet.items():
                    safe_name = "".join(c if c not in '\\/?*[]:' else "_" for c in sheet_name)
                    for pg in pages:
                        img_buf = io.BytesIO()
                        pg["image"].save(img_buf, format="PNG")
                        zf.writestr(f"{safe_name}/인쇄_{pg['page_num']:02d}페이지.png", img_buf.getvalue())
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
                        img_buf = io.BytesIO()
                        pg["image"].save(img_buf, format="PNG")
                        st.image(img_buf.getvalue(), caption=f"{pg['page_num']}페이지 ({pg['range_text']})")
                        st.download_button(
                            "다운로드",
                            data=img_buf.getvalue(),
                            file_name=f"{sheet_name}_인쇄_{pg['page_num']}페이지.png",
                            mime="image/png",
                            key=f"dl_{sheet_name}_{pg['page_num']}",
                        )
else:
    st.info("먼저 엑셀 파일을 업로드해주세요.")

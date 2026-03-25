"""건설 현장 동산보드판 OCR — Streamlit UI.

파이프라인: 이미지 업로드 → 흰 박스 크롭 → Upstage OCR → Gemini LLM 구조화 → 엑셀 + 작업로그
"""

import os
from datetime import datetime
from pathlib import Path

import streamlit as st

# Streamlit Cloud: st.secrets → 환경변수로 전달 (.env 없는 환경 대응)
for _key in ("UPSTAGE_API_KEY", "GEMINI_API_KEY", "UPSTAGE_SSL_VERIFY"):
    if _key not in os.environ:
        try:
            os.environ[_key] = st.secrets[_key]
        except (KeyError, FileNotFoundError):
            pass

from python_ocr import (
    REQUIRED_KEYS,
    RESULT_DIR,
    append_monthly,
    crop_white_box,
    llm_structure,
    ocr_extract,
    write_work_log,
)

st.set_page_config(
    page_title="동산보드판 OCR",
    page_icon="📋",
    layout="wide",
)

st.title("건설 현장 동산보드판 OCR")
st.caption(
    "「변환 시작」으로 흰 박스 크롭 → OCR → Gemini LLM 구조화까지 실행한 뒤, "
    "엑셀(.xlsx)을 내려받을 수 있습니다. 결과는 img/result/ 폴더에도 저장됩니다."
)

uploaded_files = st.file_uploader(
    "보드판 이미지를 선택하세요 (여러 장 가능)",
    type=["png", "jpg", "jpeg", "webp", "bmp", "gif", "tiff", "tif"],
    accept_multiple_files=True,
    help="PNG, JPEG, WebP 등 이미지 형식을 지원합니다.",
)

if uploaded_files:
    st.subheader("미리보기")
    ncols = min(4, len(uploaded_files))
    cols = st.columns(ncols)
    for idx, file in enumerate(uploaded_files):
        with cols[idx % ncols]:
            st.image(file, caption=file.name, use_container_width=True)

has_files = bool(uploaded_files)
if st.button("변환 시작", type="primary", disabled=not has_files):
    RESULT_DIR.mkdir(parents=True, exist_ok=True)

    results: list[dict] = []
    total = len(uploaded_files)
    total_ocr_pages = 0
    total_gemini_input = 0
    total_gemini_output = 0

    bar = st.progress(0.0, text="처리 준비 중…")
    for i, uf in enumerate(uploaded_files):
        raw = uf.getvalue()
        max_step = max(total * 3, 1)

        # 1) 흰 박스 크롭
        bar.progress(
            (i * 3) / max_step,
            text=f"크롭 ({i + 1}/{total}): {uf.name}",
        )
        try:
            cropped = crop_white_box(raw)
        except Exception:
            cropped = raw

        # 2) Upstage OCR
        bar.progress(
            (i * 3 + 1) / max_step,
            text=f"OCR ({i + 1}/{total}): {uf.name}",
        )
        try:
            ocr_text, pages = ocr_extract(cropped, uf.name)
            total_ocr_pages += pages
            ocr_error = None
        except Exception as e:
            ocr_text = ""
            ocr_error = str(e)
            results.append({
                "name": uf.name,
                "text": "",
                "error": ocr_error,
                "structured": None,
                "image_bytes": raw,
            })
            continue

        # 3) Gemini LLM 구조화
        bar.progress(
            (i * 3 + 2) / max_step,
            text=f"LLM 구조화 ({i + 1}/{total}): {uf.name}",
        )
        try:
            structured = llm_structure(ocr_text)
            total_gemini_input += structured.pop("_input_tokens", 0)
            total_gemini_output += structured.pop("_output_tokens", 0)
            s_err = None
        except Exception as e:
            structured = None
            s_err = str(e)

        results.append({
            "name": uf.name,
            "text": ocr_text,
            "error": None,
            "structured": structured,
            "structure_error": s_err,
            "image_bytes": raw,
        })

    bar.progress(1.0, text="처리 완료")

    # 월별 엑셀에 누적 저장 (img/result/YYYY-MM.xlsx)
    excel_rows = [
        {
            "filename": r["name"],
            "image_bytes": r.get("image_bytes"),
            "structured": r.get("structured") or {},
        }
        for r in results
    ]
    out_path = append_monthly(excel_rows)
    out_name = out_path.name

    # 작업 로그
    write_work_log(
        num_images=total,
        ocr_pages=total_ocr_pages,
        gemini_input_tokens=total_gemini_input,
        gemini_output_tokens=total_gemini_output,
        output_file=out_name,
    )

    st.session_state["ocr_results"] = results
    st.session_state["excel_path"] = str(out_path)
    st.session_state["excel_name"] = out_name

    ok = sum(1 for r in results if not r.get("error"))
    st_ok = sum(1 for r in results if not r.get("error") and r.get("structured"))
    st.session_state["ocr_last_message"] = (
        f"OCR 성공 {ok}건 / LLM 구조화 {st_ok}건 (전체 {len(results)}건) — "
        f"결과: img/result/{out_name}"
    )

# --- 결과 표시 ---
results = st.session_state.get("ocr_results")
if results:
    st.divider()
    st.subheader("결과")
    if st.session_state.get("ocr_last_message"):
        st.success(st.session_state["ocr_last_message"])

    excel_path = st.session_state.get("excel_path")
    if excel_path and Path(excel_path).exists():
        xlsx_bytes = Path(excel_path).read_bytes()
        st.download_button(
            label="엑셀 다운로드 (.xlsx)",
            data=xlsx_bytes,
            file_name=st.session_state.get("excel_name", "result.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    LABEL_MAP = dict(zip(REQUIRED_KEYS, ("공사명", "공종", "위치", "내용")))
    for idx, r in enumerate(results):
        if r.get("error"):
            st.error(f"**{r['name']}** — {r['error']}")
        else:
            with st.expander(f"📄 {r['name']}", expanded=(idx == 0)):
                if r.get("structured"):
                    cols = st.columns(len(REQUIRED_KEYS))
                    for ci, k in enumerate(REQUIRED_KEYS):
                        cols[ci].metric(LABEL_MAP[k], r["structured"].get(k, ""))
                st.text_area(
                    "OCR 원문",
                    value=r.get("text") or "(비어 있음)",
                    height=120,
                    key=f"ocr_text_{idx}",
                    label_visibility="collapsed",
                )
                if r.get("structure_error"):
                    st.warning(r["structure_error"])

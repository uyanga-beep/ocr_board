"""건설 현장 동산보드판 OCR — Streamlit UI.

파이프라인: 이미지 업로드 → 흰 박스 크롭 → Gemini Vision(OCR+구조화) → 엑셀 + 작업로그
"""

import os
from datetime import datetime
from pathlib import Path

import streamlit as st

for _key in ("GEMINI_API_KEY",):
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
    gemini_extract,
    write_work_log,
)

st.set_page_config(
    page_title="동산보드판 OCR",
    page_icon="📋",
    layout="wide",
)

st.title("건설 현장 동산보드판 OCR")
st.caption(
    "「변환 시작」으로 Gemini Vision을 통해 이미지에서 직접 데이터를 추출합니다. "
    "결과는 엑셀(.xlsx)로 다운로드할 수 있으며, img/result/ 폴더에 월별 누적 저장됩니다."
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
    total_gemini_input = 0
    total_gemini_output = 0

    bar = st.progress(0.0, text="처리 준비 중…")
    for i, uf in enumerate(uploaded_files):
        raw = uf.getvalue()
        max_step = max(total * 2, 1)

        bar.progress(
            (i * 2) / max_step,
            text=f"크롭 ({i + 1}/{total}): {uf.name}",
        )
        try:
            cropped = crop_white_box(raw)
        except Exception:
            cropped = raw

        bar.progress(
            (i * 2 + 1) / max_step,
            text=f"Gemini Vision ({i + 1}/{total}): {uf.name}",
        )
        try:
            structured = gemini_extract(cropped)
            total_gemini_input += structured.pop("_input_tokens", 0)
            total_gemini_output += structured.pop("_output_tokens", 0)
            s_err = None
        except Exception as e:
            structured = None
            s_err = str(e)

        results.append({
            "name": uf.name,
            "error": None if structured else s_err,
            "structured": structured,
            "structure_error": s_err,
            "image_bytes": raw,
        })

    bar.progress(1.0, text="처리 완료")

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

    write_work_log(
        num_images=total,
        gemini_input_tokens=total_gemini_input,
        gemini_output_tokens=total_gemini_output,
        output_file=out_name,
    )

    st.session_state["ocr_results"] = results
    st.session_state["excel_path"] = str(out_path)
    st.session_state["excel_name"] = out_name

    ok = sum(1 for r in results if r.get("structured"))
    st.session_state["ocr_last_message"] = (
        f"Gemini Vision 성공 {ok}건 (전체 {len(results)}건) — 결과: img/result/{out_name}"
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
                if r.get("structure_error"):
                    st.warning(r["structure_error"])

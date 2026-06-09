# -*- coding: utf-8 -*-
"""
독서활동기록 ISBN 검증 대시보드
--------------------------------
교사가 학생들의 독서기록 엑셀(.xlsx)을 업로드하면,
국립중앙도서관 서지정보 API와 대조하여 공식 등재 도서만 정제해
새로운 엑셀 파일로 반환하는 Streamlit 웹 애플리케이션입니다.

실행:  streamlit run app.py
"""

import io
import os
import re
import time

import pandas as pd
import requests
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 상수 / 설정
# ---------------------------------------------------------------------------
API_URL = "https://www.nl.go.kr/seoji/SearchApi.do"
# 참고: 문서상 엔드포인트는 http://seoji.nl.go.kr/landingPage/SearchApi.do 이며,
# 위 https 주소도 동일 API로 동작합니다. 환경에 따라 아래 FALLBACK_URL을 사용하세요.
FALLBACK_URL = "http://seoji.nl.go.kr/landingPage/SearchApi.do"

ERROR_MESSAGE = "오류: ISBN 미등재 또는 도서 정보 불일치"
REQUEST_TIMEOUT = 10          # API 요청 타임아웃(초)
REQUEST_DELAY = 0.2           # 호출 간 대기(초) - 과도한 요청 방지
PAGE_SIZE = 10                # API 검색 결과 페이지 크기

st.set_page_config(page_title="독서기록 ISBN 검증 대시보드", page_icon="📚", layout="wide")


# ---------------------------------------------------------------------------
# 유틸리티 함수
# ---------------------------------------------------------------------------
def normalize(text) -> str:
    """공백/특수문자를 제거하고 소문자화하여 비교용 문자열로 정규화."""
    if text is None:
        return ""
    s = str(text)
    s = re.sub(r"[\s\u200b]+", "", s)          # 공백 및 zero-width space 제거
    s = re.sub(r"[^\w가-힣]", "", s)           # 한글/영숫자 외 기호 제거
    return s.lower()


def get_stored_key() -> str:
    """
    숨겨둔 인증키를 안전하게 읽어온다.
    우선순위: st.secrets["NL_API_KEY"]  ->  환경변수 NL_API_KEY
    (둘 다 없으면 빈 문자열을 반환하고, 화면에서 직접 입력받는다.)
    """
    # st.secrets 는 secrets 파일이 없으면 예외를 던지므로 방어적으로 접근
    try:
        if "NL_API_KEY" in st.secrets:
            return str(st.secrets["NL_API_KEY"]).strip()
    except Exception:  # noqa: BLE001
        pass
    return os.environ.get("NL_API_KEY", "").strip()


def clean_author(author_raw: str) -> str:
    """API의 저자 필드에서 '지은이', '저', ';' 등 부가 표기를 정리."""
    if not author_raw:
        return ""
    s = str(author_raw)
    # 첫 번째 구분자 이전의 대표 저자만 사용
    s = re.split(r"[;,/]", s)[0]
    s = re.sub(r"(지은이|옮긴이|저자|저|글|그림|편|역|엮음)\s*[:：]?", "", s)
    return s.strip()


# ---------------------------------------------------------------------------
# 국립중앙도서관 API 연동
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def call_seoji_api(title: str, cert_key: str, use_fallback: bool = False) -> list:
    """
    제목으로 서지정보를 검색하여 docs(list)를 반환.
    네트워크/파싱 오류 시 빈 리스트 반환.
    cache_data 로 동일 (title, cert_key) 조합은 재호출하지 않음.
    """
    if not title or not str(title).strip():
        return []

    url = FALLBACK_URL if use_fallback else API_URL
    params = {
        "cert_key": cert_key,
        "result_style": "json",
        "page_no": 1,
        "page_size": PAGE_SIZE,
        "title": str(title).strip(),
    }
    try:
        resp = requests.get(url, params=params, timeout=REQUEST_TIMEOUT)
        resp.raise_for_status()
        data = resp.json()
    except (requests.RequestException, ValueError):
        return []

    docs = data.get("docs", [])
    return docs if isinstance(docs, list) else []


def match_book(title: str, author: str, docs: list):
    """
    검색 결과(docs) 중 입력 제목/저자와 가장 잘 맞는 항목을 선택.
    반환: (matched_doc | None, score)
      - 제목 일치(부분 포함)는 필수
      - 저자 일치 시 가중치 부여
    """
    n_title = normalize(title)
    n_author = normalize(author)
    if not n_title:
        return None, 0

    best_doc, best_score = None, 0
    for doc in docs:
        doc_title = normalize(doc.get("TITLE", ""))
        doc_author = normalize(clean_author(doc.get("AUTHOR", "")))
        if not doc_title:
            continue

        # 제목 일치 판정 (양방향 부분 포함)
        title_hit = n_title in doc_title or doc_title in n_title
        if not title_hit:
            continue

        score = 1
        if n_title == doc_title:
            score += 2  # 완전 일치 가산점

        # 저자 일치 판정 (입력 저자가 있을 때만 평가)
        if n_author:
            if n_author in doc_author or doc_author in n_author:
                score += 2
            else:
                score -= 1  # 저자 불일치 감점(오타/다른 책 방지)

        # ISBN이 있어야 정식 등재로 인정
        if not str(doc.get("EA_ISBN", "")).strip():
            score -= 2

        if score > best_score:
            best_doc, best_score = doc, score

    return best_doc, best_score


def verify_row(title: str, author: str, cert_key: str, use_fallback: bool) -> dict:
    """한 행(도서)을 검증하여 결과 dict 반환."""
    docs = call_seoji_api(title, cert_key, use_fallback)
    doc, score = match_book(title, author, docs)

    # 저자가 입력된 경우 저자까지 어느 정도 맞아야 성공으로 인정(score 기준)
    success = doc is not None and score >= 1 and str(doc.get("EA_ISBN", "")).strip()

    if success:
        clean_title = str(doc.get("TITLE", title)).strip()
        clean_auth = clean_author(doc.get("AUTHOR", author)) or str(author).strip()
        isbn = str(doc.get("EA_ISBN", "")).strip()
        publisher = str(doc.get("PUBLISHER", "")).strip()
        return {
            "검증결과": "성공",
            "정제된 도서정보": f"{clean_title} ({clean_auth})",
            "ISBN": isbn,
            "출판사": publisher,
        }
    return {
        "검증결과": "실패",
        "정제된 도서정보": ERROR_MESSAGE,
        "ISBN": "",
        "출판사": "",
    }


# ---------------------------------------------------------------------------
# 엑셀 생성(openpyxl 스타일링)
# ---------------------------------------------------------------------------
def build_excel(df: pd.DataFrame) -> bytes:
    """결과 DataFrame을 스타일이 적용된 엑셀 바이트로 변환."""
    wb = Workbook()
    ws = wb.active
    ws.title = "검증결과"

    header_fill = PatternFill("solid", fgColor="2F5496")
    header_font = Font(bold=True, color="FFFFFF")
    success_fill = PatternFill("solid", fgColor="E2EFDA")
    fail_fill = PatternFill("solid", fgColor="FCE4E4")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    columns = list(df.columns)

    # 헤더
    for c_idx, col in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=c_idx, value=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

    # 데이터
    result_col = columns.index("검증결과") if "검증결과" in columns else None
    for r_idx, (_, row) in enumerate(df.iterrows(), start=2):
        is_fail = result_col is not None and str(row.iloc[result_col]) == "실패"
        row_fill = fail_fill if is_fail else success_fill
        for c_idx, col in enumerate(columns, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=row[col])
            cell.fill = row_fill
            cell.alignment = left
            cell.border = border

    # 열 너비 자동 조정
    for c_idx, col in enumerate(columns, start=1):
        max_len = max(
            [len(str(col))] + [len(str(v)) for v in df[col].astype(str).tolist()]
        )
        ws.column_dimensions[get_column_letter(c_idx)].width = min(max(12, max_len + 4), 55)

    ws.freeze_panes = "A2"

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------
def main():
    st.title("📚 독서활동기록 ISBN 검증 대시보드")
    st.caption(
        "학생 독서기록 엑셀을 업로드하면 국립중앙도서관 서지정보 API와 대조하여 "
        "공식 등재 도서만 정제된 형태로 반환합니다."
    )

    # --- 사이드바: API 키 및 옵션 ---
    with st.sidebar:
        st.header("⚙️ 설정")
        stored_key = get_stored_key()
        if stored_key:
            # 안전 저장소(secrets/환경변수)에 키가 있으면 화면에 노출하지 않음
            cert_key = stored_key
            st.success("🔐 인증키가 안전하게 설정되어 있습니다.")
        else:
            # 저장된 키가 없을 때만 직접 입력받음(비밀번호 형태로 가려짐)
            cert_key = st.text_input(
                "국립중앙도서관 인증키 (cert_key)",
                value="",
                type="password",
                help="배포 시에는 Secrets에 등록하면 이 칸이 사라집니다.",
            ).strip()
        use_fallback = st.checkbox(
            "대체 엔드포인트(landingPage) 사용", value=False,
            help="기본 주소로 접속이 안 될 때 체크하세요.",
        )
        st.markdown("---")
        st.markdown(
            "**인증키 발급:** 국립중앙도서관 서지정보 "
            "[오픈API](https://www.nl.go.kr/NL/contents/N31101030700.do) 신청"
        )

    # --- 파일 업로드 ---
    uploaded = st.file_uploader(
        "독서기록 엑셀 파일 업로드 (.xlsx)", type=["xlsx"], accept_multiple_files=False
    )

    if uploaded is None:
        st.info("좌측 설정에서 인증키를 입력하고, 위에 엑셀 파일을 업로드해 주세요.")
        return

    # 엑셀 읽기
    try:
        df = pd.read_excel(uploaded)
    except Exception as e:  # noqa: BLE001
        st.error(f"엑셀을 읽는 중 오류가 발생했습니다: {e}")
        return

    if df.empty:
        st.warning("업로드한 파일에 데이터가 없습니다.")
        return

    st.subheader("📄 업로드 데이터 미리보기")
    st.dataframe(df.head(10), use_container_width=True)

    # --- 열 매핑 ---
    cols = list(df.columns)

    def guess(candidates):
        for i, c in enumerate(cols):
            if any(k in str(c) for k in candidates):
                return i
        return 0

    c1, c2 = st.columns(2)
    with c1:
        title_col = st.selectbox(
            "‘책 제목’ 열 선택", cols, index=guess(["제목", "책", "도서", "title"])
        )
    with c2:
        author_options = ["(없음)"] + cols
        a_guess = guess(["저자", "작가", "author"])
        author_col = st.selectbox(
            "‘저자’ 열 선택 (선택)",
            author_options,
            index=(a_guess + 1) if a_guess else 0,
        )

    # --- 검증 실행 ---
    if st.button("🔍 ISBN 검증 시작", type="primary"):
        if not cert_key:
            st.error("먼저 좌측 사이드바에서 인증키(cert_key)를 입력해 주세요.")
            return

        total = len(df)
        progress = st.progress(0.0, text="검증을 시작합니다...")
        results = []

        for idx, row in df.iterrows():
            title = row[title_col]
            author = "" if author_col == "(없음)" else row[author_col]
            res = verify_row(title, author, cert_key, use_fallback)
            results.append(res)

            done = (len(results)) / total
            progress.progress(
                done, text=f"검증 중... ({len(results)}/{total})  현재: {str(title)[:20]}"
            )
            time.sleep(REQUEST_DELAY)

        progress.empty()

        # 결과 병합
        result_df = pd.DataFrame(results, index=df.index)
        merged = pd.concat([df, result_df], axis=1)

        success_cnt = int((result_df["검증결과"] == "성공").sum())
        fail_cnt = int((result_df["검증결과"] == "실패").sum())

        # --- 통계 ---
        st.subheader("📊 검증 통계")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("전체", f"{total} 건")
        m2.metric("성공 ✅", f"{success_cnt} 건")
        m3.metric("실패 ❌", f"{fail_cnt} 건")
        rate = (success_cnt / total * 100) if total else 0
        m4.metric("성공률", f"{rate:.1f}%")

        # --- 결과 표 ---
        st.subheader("📋 검증 결과")
        st.dataframe(merged, use_container_width=True)

        # --- 엑셀 다운로드 ---
        excel_bytes = build_excel(merged)
        st.download_button(
            label="⬇️ 검증 결과 엑셀 다운로드",
            data=excel_bytes,
            file_name="독서기록_검증결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
        st.success("검증이 완료되었습니다. 위 버튼으로 결과 파일을 내려받으세요.")


if __name__ == "__main__":
    main()

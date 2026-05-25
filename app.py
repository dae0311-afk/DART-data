"""
DART 재무정보 추출 에이전트 v2 (API 기반)
- 회사 검색 → 동명회사 구분 → 보고서 조회 → 재무 추출 → 검증용 메타정보 표시
- 추후 PDF 다운로드 기반 검증 단계 추가 예정
"""

import io
from datetime import datetime
from typing import Optional

import pandas as pd
import streamlit as st
import OpenDartReader

# ====================================================================
# 0) 페이지 설정
# ====================================================================
st.set_page_config(
    page_title="DART 재무정보 추출 에이전트",
    page_icon="📊",
    layout="wide",
)

st.title("📊 DART 재무정보 추출 에이전트")
st.caption("v2 (API 기반) — 추후 PDF 자동 검증 단계 추가 예정")

# ====================================================================
# 1) 상수 정의
# ====================================================================
# 보고서 종류 코드 (DART API 표준)
REPRT_CODE = {
    "사업보고서": "11011",
    "반기보고서": "11012",
    "1분기보고서": "11013",
    "3분기보고서": "11014",
}

# 재무제표 구분 (CFS=연결, OFS=별도)
FS_DIV_OPTIONS = {
    "연결재무제표(CFS)": "CFS",
    "별도재무제표(OFS)": "OFS",
}

# 추출 대상 계정 - 키워드 매칭 (회사마다 표기 차이 흡수)
# 확장 시 이 딕셔너리에만 추가하면 됨
TARGET_ACCOUNTS = {
    "매출액": {
        "keywords": ["매출액", "수익(매출액)", "영업수익", "매출", "수익"],
        "sj_div": "IS",  # 손익계산서
        "sj_div_alt": "CIS",  # 포괄손익계산서도 허용
    },
    "영업이익": {
        "keywords": ["영업이익", "영업이익(손실)", "영업손실"],
        "sj_div": "IS",
        "sj_div_alt": "CIS",
    },
}

# ====================================================================
# 2) API 키 입력
# ====================================================================
api_key = ""
try:
    api_key = st.secrets.get("DART_API_KEY", "")
except (FileNotFoundError, Exception):
    api_key = ""

if not api_key:
    api_key = st.sidebar.text_input(
        "DART OpenAPI 인증키",
        type="password",
        help="https://opendart.fss.or.kr 에서 발급받은 40자리 인증키",
    )

if not api_key:
    st.info("좌측 사이드바에 DART OpenAPI 인증키를 입력하면 시작할 수 있습니다.")
    st.markdown(
        "**키 발급 절차**\n"
        "1. [opendart.fss.or.kr](https://opendart.fss.or.kr/) 회원가입\n"
        "2. 인증키 신청/관리 → 인증키 신청\n"
        "3. 발급된 40자리 키 입력"
    )
    st.stop()

@st.cache_resource(show_spinner=False)
def get_dart(key: str) -> OpenDartReader:
    return OpenDartReader(key)

try:
    dart = get_dart(api_key)
except Exception as e:
    st.error(f"API 키 초기화 실패: {e}")
    st.stop()

# ====================================================================
# 3) 유틸리티 함수
# ====================================================================
def to_number(x) -> Optional[int]:
    """문자열 금액('1,234,567' 또는 '-1,234')을 정수로 변환."""
    if x is None or pd.isna(x):
        return None
    s = str(x).strip().replace(",", "").replace(" ", "")
    if s in ("", "-", "nan", "None"):
        return None
    try:
        # 괄호로 음수 표시되는 경우 처리: (1,234) → -1234
        if s.startswith("(") and s.endswith(")"):
            s = "-" + s[1:-1]
        return int(float(s))
    except (ValueError, TypeError):
        return None


def format_amount(v: Optional[int], unit_divisor: int = 1) -> str:
    """금액 표시 포맷. unit_divisor=1000000이면 백만원 단위."""
    if v is None:
        return "-"
    return f"{v // unit_divisor:,}" if unit_divisor > 1 else f"{v:,}"


def find_account(df: pd.DataFrame, keywords: list, sj_div_filter: list = None) -> dict:
    """
    재무제표 DataFrame에서 keywords와 일치하는 계정 행을 찾음.
    sj_div_filter가 주어지면 해당 재무제표 구분만 검색.
    """
    work_df = df.copy()
    if sj_div_filter and "sj_div" in work_df.columns:
        work_df = work_df[work_df["sj_div"].isin(sj_div_filter)]

    if work_df.empty:
        return {"matched": False, "row": None, "match_type": None}

    # 1) 정확 일치
    exact = work_df[work_df["account_nm"].isin(keywords)]
    if not exact.empty:
        return {"matched": True, "row": exact.iloc[0], "match_type": "정확일치"}

    # 2) 부분 일치 (정규식)
    pattern = "|".join([k.replace("(", r"\(").replace(")", r"\)") for k in keywords])
    partial = work_df[work_df["account_nm"].astype(str).str.contains(pattern, na=False, regex=True)]
    if not partial.empty:
        return {"matched": True, "row": partial.iloc[0], "match_type": "부분일치"}

    return {"matched": False, "row": None, "match_type": None}


@st.cache_data(ttl=3600, show_spinner=False)
def search_companies(_dart: OpenDartReader, query: str) -> pd.DataFrame:
    """회사명으로 검색 (동명회사 모두 반환)."""
    try:
        # 종목코드(6자리 숫자) 또는 고유번호(8자리)면 단일 조회
        q = query.strip()
        if q.isdigit() and len(q) in (6, 8):
            info = _dart.company(q)
            if info:
                return pd.DataFrame([info])
            return pd.DataFrame()
        # 회사명 검색
        result = _dart.company_by_name(q)
        if result is None or (hasattr(result, "empty") and result.empty):
            return pd.DataFrame()
        return result if isinstance(result, pd.DataFrame) else pd.DataFrame(result)
    except Exception as e:
        st.warning(f"회사 검색 중 오류: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=600, show_spinner=False)
def fetch_finstate_all(_dart: OpenDartReader, corp_code: str, year: int,
                       reprt_code: str, fs_div: str) -> Optional[pd.DataFrame]:
    """전체 재무제표 조회."""
    try:
        df = _dart.finstate_all(corp_code, year, reprt_code=reprt_code, fs_div=fs_div)
        if df is None or (hasattr(df, "empty") and df.empty):
            return None
        return df
    except Exception:
        return None


@st.cache_data(ttl=600, show_spinner=False)
def fetch_finstate_simple(_dart: OpenDartReader, corp_code: str, year: int,
                          reprt_code: str) -> Optional[pd.DataFrame]:
    """주요계정 조회 (상장사만, fallback용)."""
    try:
        df = _dart.finstate(corp_code, year, reprt_code=reprt_code)
        if df is None or (hasattr(df, "empty") and df.empty):
            return None
        return df
    except Exception:
        return None


@st.cache_data(ttl=600, show_spinner=False)
def fetch_report_list(_dart: OpenDartReader, corp_code: str, year: int) -> pd.DataFrame:
    """해당 회사의 해당 연도 정기공시 목록 조회 (메타정보 수집용)."""
    try:
        start = f"{year}0101"
        end = f"{year+1}0630"  # 사업보고서는 다음해 상반기 제출
        reports = _dart.list(corp_code, start=start, end=end, kind="A")  # A: 정기공시
        if reports is None or (hasattr(reports, "empty") and reports.empty):
            return pd.DataFrame()
        return reports
    except Exception:
        return pd.DataFrame()


# ====================================================================
# 4) 사이드바 - 검색 조건
# ====================================================================
st.sidebar.header("🔍 검색 조건")

company_input = st.sidebar.text_input(
    "회사명 또는 코드",
    value="",
    placeholder="예: 삼성전자 / 005930",
    help="회사명, 종목코드(6자리), 고유번호(8자리) 모두 가능",
)

year = st.sidebar.number_input(
    "사업연도",
    min_value=2015,
    max_value=datetime.now().year,
    value=datetime.now().year - 1,
    step=1,
)

report_label = st.sidebar.selectbox(
    "보고서 종류",
    options=list(REPRT_CODE.keys()),
    index=0,
)

# 연결/별도 모두 조회 옵션
fetch_both_fs = st.sidebar.checkbox(
    "연결·별도 모두 조회 (권장)",
    value=True,
    help="둘을 비교 표시하여 혼동을 방지합니다",
)

if not fetch_both_fs:
    fs_label = st.sidebar.radio("재무제표 구분", list(FS_DIV_OPTIONS.keys()), index=0)
    fs_div_targets = [FS_DIV_OPTIONS[fs_label]]
else:
    fs_div_targets = ["CFS", "OFS"]

unit_label = st.sidebar.selectbox(
    "표시 단위",
    options=["원", "천원", "백만원", "억원"],
    index=2,
)
UNIT_DIVISOR = {"원": 1, "천원": 1_000, "백만원": 1_000_000, "억원": 100_000_000}[unit_label]

st.sidebar.divider()
search_btn = st.sidebar.button("1️⃣ 회사 검색", use_container_width=True)

# ====================================================================
# 5) 단계 1: 회사 검색
# ====================================================================
if search_btn:
    if not company_input.strip():
        st.warning("회사명 또는 코드를 입력하세요.")
        st.stop()

    with st.spinner(f"'{company_input}' 검색 중..."):
        companies = search_companies(dart, company_input)

    if companies.empty:
        st.error(
            f"'{company_input}'에 해당하는 회사를 찾을 수 없습니다.\n"
            "- 정확한 법인명을 입력했는지 확인하세요 (예: '주식회사' 제외 가능)\n"
            "- 종목코드(6자리) 또는 DART 고유번호(8자리)로도 검색 가능합니다"
        )
        st.stop()

    st.session_state["companies"] = companies
    st.session_state["selected_corp"] = None

# 검색 결과 표시 및 선택
if "companies" in st.session_state and not st.session_state["companies"].empty:
    companies = st.session_state["companies"]
    st.subheader(f"1️⃣ 검색 결과 ({len(companies)}건)")

    # 표시용 컬럼 선택
    display_cols = []
    for c in ["corp_name", "corp_code", "stock_code", "ceo_nm", "corp_cls",
              "est_dt", "adres", "induty_code"]:
        if c in companies.columns:
            display_cols.append(c)

    st.dataframe(
        companies[display_cols] if display_cols else companies,
        use_container_width=True,
        hide_index=True,
    )

    # 동명회사 선택
    options = []
    for _, row in companies.iterrows():
        label = f"{row.get('corp_name', '?')} (고유번호: {row.get('corp_code', '?')}"
        if row.get("stock_code") and str(row.get("stock_code")).strip():
            label += f", 종목코드: {row.get('stock_code')}"
        label += f", 구분: {row.get('corp_cls', '?')})"
        options.append(label)

    selected_idx = st.selectbox(
        "조회할 회사를 선택하세요",
        options=range(len(options)),
        format_func=lambda i: options[i],
        key="company_select",
    )
    selected_corp = companies.iloc[selected_idx]
    st.session_state["selected_corp"] = selected_corp

    st.info(
        f"선택: **{selected_corp.get('corp_name')}** | "
        f"고유번호 `{selected_corp.get('corp_code')}` | "
        f"법인구분 `{selected_corp.get('corp_cls', '?')}` "
        f"(Y=유가증권, K=코스닥, N=코넥스, E=기타외감)"
    )

    extract_btn = st.button("2️⃣ 재무 데이터 추출", type="primary", use_container_width=True)

    # ================================================================
    # 6) 단계 2: 재무 데이터 추출
    # ================================================================
    if extract_btn:
        corp_code = selected_corp.get("corp_code")
        corp_cls = selected_corp.get("corp_cls", "")
        corp_name = selected_corp.get("corp_name", "")

        # 비상장 외감기업 경고
        if corp_cls == "E":
            st.warning(
                "⚠️ **외감 비상장기업(E)입니다.** DART OpenAPI는 비상장사 재무를 "
                "전체 또는 일부만 제공할 수 있습니다. 결과가 비어있으면 PDF 본문 검증 단계가 필수입니다."
            )

        # 공시 목록 메타정보
        with st.spinner("공시 목록 조회 중..."):
            report_list = fetch_report_list(dart, corp_code, year)

        # 두 재무제표 구분 모두 시도
        all_results = {}
        raw_dfs = {}

        for fs_div in fs_div_targets:
            with st.spinner(f"{fs_div} 재무제표 조회 중..."):
                df = fetch_finstate_all(
                    dart, corp_code, year,
                    reprt_code=REPRT_CODE[report_label],
                    fs_div=fs_div,
                )

            if df is None or df.empty:
                # 상장사면 finstate fallback 시도
                if corp_cls in ("Y", "K"):
                    df = fetch_finstate_simple(
                        dart, corp_code, year, reprt_code=REPRT_CODE[report_label]
                    )
                    # finstate는 fs_div를 자체 컬럼에서 필터링 필요
                    if df is not None and not df.empty and "fs_div" in df.columns:
                        df = df[df["fs_div"] == fs_div]

            if df is None or df.empty:
                all_results[fs_div] = None
                continue

            raw_dfs[fs_div] = df

            # 추출 대상 계정별 매칭
            extracted = []
            for label, spec in TARGET_ACCOUNTS.items():
                sj_filter = [spec["sj_div"], spec["sj_div_alt"]]
                result = find_account(df, spec["keywords"], sj_div_filter=sj_filter)

                if not result["matched"]:
                    extracted.append({
                        "항목": label,
                        "원문계정명": "(미발견)",
                        "매칭": "실패",
                        "당기": None,
                        "전기": None,
                        "전전기": None,
                        "통화": None,
                    })
                    continue

                row = result["row"]
                extracted.append({
                    "항목": label,
                    "원문계정명": row.get("account_nm"),
                    "매칭": result["match_type"],
                    "당기": to_number(row.get("thstrm_amount")),
                    "전기": to_number(row.get("frmtrm_amount")),
                    "전전기": to_number(row.get("bfefrmtrm_amount")),
                    "통화": row.get("currency", "KRW"),
                    "재무제표구분": row.get("sj_div"),
                })

            all_results[fs_div] = pd.DataFrame(extracted)

        # ============================================================
        # 7) 결과 표시
        # ============================================================
        st.subheader(f"2️⃣ {corp_name} | {year}년 {report_label}")

        # 공시 메타정보
        if not report_list.empty:
            with st.expander("📋 해당 연도 공시 목록 (PDF 검증용 메타정보)", expanded=False):
                meta_cols = [c for c in ["rcept_no", "report_nm", "rcept_dt", "flr_nm"]
                             if c in report_list.columns]
                st.dataframe(
                    report_list[meta_cols] if meta_cols else report_list,
                    use_container_width=True,
                    hide_index=True,
                )
                st.caption(
                    "각 rcept_no는 DART 공시 원문 링크의 핵심 식별자입니다. "
                    "예: `https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcept_no}`"
                )

        # 연결·별도 나란히 표시
        if len(fs_div_targets) == 2:
            col1, col2 = st.columns(2)
            for col, fs_div, title in [
                (col1, "CFS", "🔵 연결재무제표(CFS)"),
                (col2, "OFS", "🟢 별도재무제표(OFS)"),
            ]:
                with col:
                    st.markdown(f"### {title}")
                    res = all_results.get(fs_div)
                    if res is None or res.empty:
                        st.warning(f"{fs_div} 데이터를 가져오지 못했습니다.")
                        continue
                    display = res.copy()
                    for c in ["당기", "전기", "전전기"]:
                        display[c] = display[c].apply(lambda v: format_amount(v, UNIT_DIVISOR))
                    display.columns = [
                        "항목", "원문계정명", "매칭",
                        f"당기({year})", f"전기({year-1})", f"전전기({year-2})",
                        "통화", "재무제표구분",
                    ][:len(display.columns)]
                    st.dataframe(display, use_container_width=True, hide_index=True)
                    st.caption(f"단위: {unit_label}")
        else:
            fs_div = fs_div_targets[0]
            res = all_results.get(fs_div)
            if res is None or res.empty:
                st.error(f"{fs_div} 재무 데이터를 가져오지 못했습니다.")
            else:
                display = res.copy()
                for c in ["당기", "전기", "전전기"]:
                    display[c] = display[c].apply(lambda v: format_amount(v, UNIT_DIVISOR))
                st.dataframe(display, use_container_width=True, hide_index=True)
                st.caption(f"단위: {unit_label}")

        # 증감 분석 (연결 우선, 없으면 별도)
        primary_fs = "CFS" if all_results.get("CFS") is not None else "OFS"
        primary_df = all_results.get(primary_fs)
        if primary_df is not None and not primary_df.empty:
            st.markdown(f"### 📈 전년 대비 증감 ({primary_fs})")
            growth_rows = []
            for _, r in primary_df.iterrows():
                curr, prev = r["당기"], r["전기"]
                if curr is None or prev is None or prev == 0:
                    continue
                growth_rows.append({
                    "항목": r["항목"],
                    f"당기({year})": format_amount(curr, UNIT_DIVISOR),
                    f"전기({year-1})": format_amount(prev, UNIT_DIVISOR),
                    "증감액": format_amount(curr - prev, UNIT_DIVISOR),
                    "증감률(%)": f"{(curr - prev) / abs(prev) * 100:+.2f}%",
                })
            if growth_rows:
                st.dataframe(pd.DataFrame(growth_rows), use_container_width=True, hide_index=True)

        # 엑셀 다운로드
        st.markdown("### 📥 엑셀 다운로드")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            # 회사정보
            pd.DataFrame([selected_corp]).to_excel(writer, sheet_name="회사정보", index=False)
            # 추출결과
            for fs_div, res in all_results.items():
                if res is not None and not res.empty:
                    res.to_excel(writer, sheet_name=f"추출_{fs_div}", index=False)
            # 원본 전체 재무제표
            for fs_div, raw in raw_dfs.items():
                raw.to_excel(writer, sheet_name=f"원본_{fs_div}", index=False)
            # 공시 목록
            if not report_list.empty:
                report_list.to_excel(writer, sheet_name="공시목록", index=False)

        st.download_button(
            label="📥 전체 결과 엑셀 다운로드",
            data=output.getvalue(),
            file_name=f"{corp_name}_{year}_{report_label}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        # 디버그용 원본 표시
        with st.expander("🔬 원본 재무제표 전체 행 보기 (검증용)"):
            for fs_div, raw in raw_dfs.items():
                st.markdown(f"**{fs_div}**")
                st.dataframe(raw, use_container_width=True)

        # 향후 PDF 검증 단계 placeholder
        st.divider()
        st.info(
            "🔜 **다음 단계 (개발 예정): PDF 자동 검증**\n\n"
            "위 공시 목록의 접수번호(rcept_no)로 감사보고서/연결감사보고서 PDF를 "
            "자동 다운로드하고, 본문 표에서 동일 항목을 재추출하여 API 값과 대조합니다. "
            "불일치 시 경고를 표시합니다."
        )

# ====================================================================
# 8) 하단 안내
# ====================================================================
st.sidebar.divider()
with st.sidebar.expander("ℹ️ 사용 가이드"):
    st.markdown(
        "**순서**\n"
        "1. 회사명/코드 입력 → `회사 검색`\n"
        "2. 동명회사 중 정확한 회사 선택\n"
        "3. `재무 데이터 추출` 클릭\n"
        "4. 연결·별도 결과 확인 → 엑셀 다운로드\n\n"
        "**주의**\n"
        "- 비상장 외감기업(E)은 API에 데이터가 없을 수 있음\n"
        "- 금융업은 `finstate_all`에서 제외됨\n"
        "- 계정 매칭은 '키워드 기반'이므로 추후 PDF 본문 대조 필수"
    )

"""
DART 재무정보 추출 에이전트 v3
- PE 실무 5개년 재무 요약 템플릿 자동 생성
- 추출 항목: 매출액, EBITDA, 영업이익, 당기순이익, 자산총계, 현금성자산, 부채총계, 총차입금, 자본총계
- Growth/Margin 자동 계산
- 추후 PDF 검증 단계 연결 예정
"""

import io
from datetime import datetime
from typing import Optional, Dict, List

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
st.caption("v3 — PE 5개년 재무요약 템플릿 (API 기반, PDF 검증 단계 추가 예정)")

# ====================================================================
# 1) 상수 정의
# ====================================================================
REPRT_CODE_ANNUAL = "11011"  # 사업보고서

# 추출 대상 계정 정의
# - sj_filter: 재무제표 구분 (BS=재무상태표, IS=손익계산서, CIS=포괄손익계산서, CF=현금흐름표)
# - keywords: 매칭할 계정명 후보 (정확일치 → 부분일치 순서)
# - sign: 부호 (+1 또는 -1, 비용/차감 항목은 합산 시 부호 조정)
TARGET_ACCOUNTS = {
    # 손익계산서
    "매출액": {
        "sj": ["IS", "CIS"],
        "keywords": ["매출액", "수익(매출액)", "영업수익", "매출"],
    },
    "영업이익": {
        "sj": ["IS", "CIS"],
        "keywords": ["영업이익", "영업이익(손실)"],
    },
    "당기순이익": {
        "sj": ["IS", "CIS"],
        "keywords": ["당기순이익", "당기순이익(손실)", "분기순이익", "반기순이익", "당기순손익"],
    },
    # 재무상태표
    "자산총계": {
        "sj": ["BS"],
        "keywords": ["자산총계"],
    },
    "현금성자산": {
        "sj": ["BS"],
        "keywords": ["현금및현금성자산", "현금 및 현금성자산"],
    },
    "부채총계": {
        "sj": ["BS"],
        "keywords": ["부채총계"],
    },
    "자본총계": {
        "sj": ["BS"],
        "keywords": ["자본총계", "지배기업소유주지분", "자본 총계"],
    },
    # 차입금 (합산 대상)
    "_단기차입금": {
        "sj": ["BS"],
        "keywords": ["단기차입금"],
    },
    "_유동성장기부채": {
        "sj": ["BS"],
        "keywords": ["유동성장기부채", "유동성장기차입금", "유동성사채"],
    },
    "_장기차입금": {
        "sj": ["BS"],
        "keywords": ["장기차입금"],
    },
    "_사채": {
        "sj": ["BS"],
        "keywords": ["사채", "회사채"],
    },
    # EBITDA 산출용 (현금흐름표 가산항목)
    "_유형자산감가상각비": {
        "sj": ["CF"],
        "keywords": ["감가상각비", "유형자산감가상각비"],
    },
    "_무형자산상각비": {
        "sj": ["CF"],
        "keywords": ["무형자산상각비", "무형자산상각"],
    },
    "_사용권자산상각비": {
        "sj": ["CF"],
        "keywords": ["사용권자산상각비"],
    },
}

# ====================================================================
# 2) API 키
# ====================================================================
api_key = ""
try:
    api_key = st.secrets.get("DART_API_KEY", "")
except Exception:
    api_key = ""

if not api_key:
    api_key = st.sidebar.text_input("DART OpenAPI 인증키", type="password")

if not api_key:
    st.info("좌측 사이드바에 DART OpenAPI 인증키를 입력하세요.")
    st.stop()

@st.cache_resource(show_spinner=False)
def get_dart(key: str):
    return OpenDartReader(key)

try:
    dart = get_dart(api_key)
except Exception as e:
    st.error(f"API 키 초기화 실패: {e}")
    st.stop()

# ====================================================================
# 3) 유틸리티
# ====================================================================
def to_number(x) -> Optional[int]:
    if x is None or pd.isna(x):
        return None
    s = str(x).strip().replace(",", "").replace(" ", "")
    if s in ("", "-", "nan", "None"):
        return None
    try:
        if s.startswith("(") and s.endswith(")"):
            s = "-" + s[1:-1]
        return int(float(s))
    except (ValueError, TypeError):
        return None


def to_eokwon(v: Optional[int]) -> Optional[float]:
    """원 단위 → 억원 단위 (소수점 없이 반올림)."""
    if v is None:
        return None
    return round(v / 100_000_000)


def format_eokwon(v: Optional[float]) -> str:
    if v is None:
        return "N/A"
    return f"{int(v):,}" if v >= 0 else f"({int(abs(v)):,})"


def format_pct(v: Optional[float]) -> str:
    if v is None:
        return "N/A"
    if v < 0:
        return f"({abs(v):.1f}%)"
    return f"{v:.1f}%"


def find_account_in_df(df: pd.DataFrame, keywords: list, sj_filter: list) -> Optional[pd.Series]:
    """단일 재무제표 DataFrame에서 계정 찾기."""
    if df is None or df.empty:
        return None
    work = df.copy()
    if "sj_div" in work.columns:
        work = work[work["sj_div"].isin(sj_filter)]
    if work.empty:
        return None

    # 정확 일치 우선
    exact = work[work["account_nm"].isin(keywords)]
    if not exact.empty:
        return exact.iloc[0]

    # 부분 일치
    pattern = "|".join([k.replace("(", r"\(").replace(")", r"\)") for k in keywords])
    partial = work[work["account_nm"].astype(str).str.contains(pattern, na=False, regex=True)]
    if not partial.empty:
        # 가장 짧은 계정명 우선 (예: "매출액"이 "매출액(영업수익)"보다 우선)
        partial = partial.copy()
        partial["_len"] = partial["account_nm"].astype(str).str.len()
        partial = partial.sort_values("_len")
        return partial.iloc[0]

    return None


# ====================================================================
# 4) 데이터 수집
# ====================================================================
@st.cache_data(ttl=3600, show_spinner=False)
def search_companies(_dart, query: str) -> pd.DataFrame:
    try:
        q = query.strip()
        if q.isdigit() and len(q) in (6, 8):
            info = _dart.company(q)
            if info is None:
                return pd.DataFrame()
            return pd.DataFrame([info]) if isinstance(info, dict) else pd.DataFrame(info)
        result = _dart.company_by_name(q)
        if result is None:
            return pd.DataFrame()
        return result if isinstance(result, pd.DataFrame) else pd.DataFrame(result)
    except Exception as e:
        st.warning(f"검색 오류: {e}")
        return pd.DataFrame()


@st.cache_data(ttl=600, show_spinner=False)
def fetch_finstate(_dart, corp_code: str, year: int, fs_div: str) -> Optional[pd.DataFrame]:
    """단일 연도 전체 재무제표 조회."""
    try:
        df = _dart.finstate_all(corp_code, year, reprt_code=REPRT_CODE_ANNUAL, fs_div=fs_div)
        if df is None or (hasattr(df, "empty") and df.empty):
            return None
        return df
    except Exception:
        return None


def extract_year_data(df: pd.DataFrame, year_offset: int = 0) -> Dict[str, Optional[int]]:
    """
    하나의 finstate_all 결과에서 모든 타겟 계정 값을 추출.
    year_offset: 0=당기, 1=전기, 2=전전기
    """
    if df is None or df.empty:
        return {key: None for key in TARGET_ACCOUNTS.keys()}

    amount_col = {0: "thstrm_amount", 1: "frmtrm_amount", 2: "bfefrmtrm_amount"}[year_offset]

    result = {}
    for key, spec in TARGET_ACCOUNTS.items():
        row = find_account_in_df(df, spec["keywords"], spec["sj"])
        if row is None:
            result[key] = None
        else:
            result[key] = to_number(row.get(amount_col))
    return result


def collect_multi_year_data(_dart, corp_code: str, years: List[int], fs_div: str,
                            progress_callback=None) -> Dict[int, Dict[str, Optional[int]]]:
    """
    여러 연도 데이터를 효율적으로 수집.
    각 사업보고서는 당기/전기/전전기 3개년을 제공하므로 3년마다 호출.
    """
    yearly_data = {y: {} for y in years}
    missing_reasons = {y: None for y in years}

    # 호출할 연도들 선정: 가장 최근 연도부터 3년씩 묶어서
    sorted_years = sorted(years, reverse=True)
    fetch_years = []  # 실제로 API 호출할 기준 연도

    covered = set()
    for y in sorted_years:
        if y in covered:
            continue
        fetch_years.append(y)
        # 이 호출이 커버하는 3개년 = y, y-1, y-2
        covered.update([y, y - 1, y - 2])

    total = len(fetch_years)
    for idx, base_year in enumerate(fetch_years):
        if progress_callback:
            progress_callback(idx, total, base_year)

        df = fetch_finstate(_dart, corp_code, base_year, fs_div)
        if df is None:
            # 해당 연도 보고서 자체가 없음
            for offset in [0, 1, 2]:
                ty = base_year - offset
                if ty in years and not yearly_data[ty]:
                    missing_reasons[ty] = "보고서 미공시 또는 API 데이터 없음"
            continue

        # 이 호출에서 3개년 추출
        for offset in [0, 1, 2]:
            target_year = base_year - offset
            if target_year in years:
                # 이미 채워진 연도는 건너뜀 (더 최근 호출이 우선)
                if yearly_data[target_year]:
                    continue
                vals = extract_year_data(df, year_offset=offset)
                yearly_data[target_year] = vals

    if progress_callback:
        progress_callback(total, total, None)

    return yearly_data, missing_reasons


# ====================================================================
# 5) 템플릿 구성
# ====================================================================
def build_template_table(yearly_data: Dict[int, Dict[str, Optional[int]]],
                         missing_reasons: Dict[int, str],
                         years: List[int]) -> pd.DataFrame:
    """
    이미지의 PE 5개년 템플릿 구조로 변환 (단위: 억원).
    """
    rows = []

    def get_val(year, key):
        d = yearly_data.get(year, {})
        return d.get(key)

    def safe_sum(year, keys):
        d = yearly_data.get(year, {})
        vals = [d.get(k) for k in keys if d.get(k) is not None]
        if not vals:
            return None
        return sum(vals)

    # 매출액
    revenues = {y: get_val(y, "매출액") for y in years}
    rows.append(["매출액"] + [format_eokwon(to_eokwon(revenues[y])) for y in years])

    # Growth
    growth_row = ["  Growth"]
    sorted_years = sorted(years)
    for i, y in enumerate(sorted_years):
        if i == 0:
            growth_row.append("N/A")
        else:
            curr = revenues[y]
            prev = revenues[sorted_years[i - 1]]
            if curr is None or prev is None or prev == 0:
                growth_row.append("N/A")
            else:
                growth = (curr - prev) / abs(prev) * 100
                growth_row.append(format_pct(growth))
    rows.append(growth_row)

    # EBITDA = 영업이익 + 감가상각비 + 무형자산상각비 + 사용권자산상각비
    ebitda = {}
    for y in years:
        op = get_val(y, "영업이익")
        if op is None:
            ebitda[y] = None
            continue
        da = safe_sum(y, ["_유형자산감가상각비", "_무형자산상각비", "_사용권자산상각비"])
        if da is None:
            # 감가상각비가 전혀 안 잡히면 영업이익만으로는 EBITDA 산출 부정확
            ebitda[y] = None
        else:
            ebitda[y] = op + da

    rows.append(["EBITDA"] + [format_eokwon(to_eokwon(ebitda[y])) for y in years])
    # EBITDA Margin
    margin_row = ["  Margin"]
    for y in years:
        if ebitda[y] is None or revenues[y] is None or revenues[y] == 0:
            margin_row.append("N/A")
        else:
            margin_row.append(format_pct(ebitda[y] / revenues[y] * 100))
    rows.append(margin_row)

    # 영업이익
    op_inc = {y: get_val(y, "영업이익") for y in years}
    rows.append(["영업이익"] + [format_eokwon(to_eokwon(op_inc[y])) for y in years])
    margin_row = ["  Margin"]
    for y in years:
        if op_inc[y] is None or revenues[y] is None or revenues[y] == 0:
            margin_row.append("N/A")
        else:
            margin_row.append(format_pct(op_inc[y] / revenues[y] * 100))
    rows.append(margin_row)

    # 당기순이익
    net_inc = {y: get_val(y, "당기순이익") for y in years}
    rows.append(["당기순이익"] + [format_eokwon(to_eokwon(net_inc[y])) for y in years])
    margin_row = ["  Margin"]
    for y in years:
        if net_inc[y] is None or revenues[y] is None or revenues[y] == 0:
            margin_row.append("N/A")
        else:
            margin_row.append(format_pct(net_inc[y] / revenues[y] * 100))
    rows.append(margin_row)

    # 자산총계
    rows.append(["자산총계"] + [format_eokwon(to_eokwon(get_val(y, "자산총계"))) for y in years])
    # 현금성자산
    rows.append(["  현금성자산"] + [format_eokwon(to_eokwon(get_val(y, "현금성자산"))) for y in years])
    # 부채총계
    rows.append(["부채총계"] + [format_eokwon(to_eokwon(get_val(y, "부채총계"))) for y in years])
    # 총차입금
    total_debt = {}
    for y in years:
        total_debt[y] = safe_sum(y, ["_단기차입금", "_유동성장기부채", "_장기차입금", "_사채"])
    rows.append(["  총차입금"] + [format_eokwon(to_eokwon(total_debt[y])) for y in years])
    # 자본총계
    rows.append(["자본총계"] + [format_eokwon(to_eokwon(get_val(y, "자본총계"))) for y in years])

    columns = ["(단위: 억원)"] + [str(y) for y in years]
    return pd.DataFrame(rows, columns=columns)


# ====================================================================
# 6) 사이드바 - 검색
# ====================================================================
st.sidebar.header("🔍 검색 조건")

company_input = st.sidebar.text_input(
    "회사명 또는 코드",
    placeholder="예: 삼성전자 / 005930",
)

period_label = st.sidebar.selectbox(
    "조회 기간",
    options=["최근 5년", "최근 10년", "최근 20년", "최대 (2015~)"],
    index=0,
)
period_map = {"최근 5년": 5, "최근 10년": 10, "최근 20년": 20, "최대 (2015~)": 99}

fs_label = st.sidebar.radio(
    "재무제표 구분",
    options=["연결재무제표(CFS)", "별도재무제표(OFS)"],
    index=0,
)
fs_div_target = "CFS" if "연결" in fs_label else "OFS"

# 종료 연도 자동 결정 (직전 사업연도)
current_year = datetime.now().year
default_end_year = current_year - 1

end_year = st.sidebar.number_input(
    "종료 연도",
    min_value=2015,
    max_value=current_year,
    value=default_end_year,
    step=1,
    help="기본값은 직전 사업연도. 최신 사업보고서가 미공시면 자동으로 N/A 처리됨.",
)

search_btn = st.sidebar.button("1️⃣ 회사 검색", use_container_width=True)

# ====================================================================
# 7) 회사 검색 단계
# ====================================================================
if search_btn:
    if not company_input.strip():
        st.warning("회사명 또는 코드를 입력하세요.")
        st.stop()
    with st.spinner(f"'{company_input}' 검색 중..."):
        companies = search_companies(dart, company_input)
    if companies.empty:
        st.error(f"'{company_input}'에 해당하는 회사를 찾을 수 없습니다.")
        st.stop()
    st.session_state["companies"] = companies

if "companies" in st.session_state and not st.session_state["companies"].empty:
    companies = st.session_state["companies"]
    st.subheader(f"1️⃣ 검색 결과 ({len(companies)}건)")

    display_cols = [c for c in ["corp_name", "corp_code", "stock_code", "ceo_nm",
                                "corp_cls", "est_dt", "adres", "induty_code"]
                    if c in companies.columns]
    st.dataframe(
        companies[display_cols] if display_cols else companies,
        use_container_width=True, hide_index=True,
    )

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
    corp_code = selected_corp.get("corp_code")
    corp_cls = selected_corp.get("corp_cls", "")
    corp_name = selected_corp.get("corp_name", "")

    st.info(
        f"선택: **{corp_name}** | 고유번호 `{corp_code}` | "
        f"법인구분 `{corp_cls}` (Y=유가증권, K=코스닥, N=코넥스, E=기타외감)"
    )

    if corp_cls == "E":
        st.warning(
            "⚠️ **외감 비상장기업**입니다. DART OpenAPI는 비상장사 재무를 "
            "전체 또는 일부만 제공할 수 있습니다. 결과는 PDF 본문 검증이 필수입니다."
        )

    extract_btn = st.button("2️⃣ 5개년 템플릿 추출", type="primary", use_container_width=True)

    # ================================================================
    # 8) 데이터 추출 단계
    # ================================================================
    if extract_btn:
        period_n = period_map[period_label]
        if period_n == 99:
            start_year = 2015
        else:
            start_year = max(2015, end_year - period_n + 1)
        years = list(range(start_year, end_year + 1))

        st.subheader(f"2️⃣ {corp_name} | {start_year}~{end_year} | {fs_div_target}")

        progress = st.progress(0.0)
        status = st.empty()

        def update_progress(idx, total, year):
            if total > 0:
                progress.progress(idx / total)
            if year:
                status.text(f"📡 DART API 호출: {year}년 사업보고서 ({idx+1}/{total})")

        yearly_data, missing_reasons = collect_multi_year_data(
            dart, corp_code, years, fs_div_target,
            progress_callback=update_progress,
        )
        progress.empty()
        status.empty()

        # 템플릿 표 생성
        template_df = build_template_table(yearly_data, missing_reasons, years)

        # 화면 표시
        st.dataframe(
            template_df,
            use_container_width=True,
            hide_index=True,
        )

        # 결측 원인 안내
        empty_years = [y for y in years if not yearly_data.get(y)
                       or all(v is None for v in yearly_data[y].values())]
        if empty_years:
            st.warning(
                f"⚠️ 다음 연도는 데이터를 가져오지 못했습니다: {', '.join(map(str, empty_years))}\n\n"
                "**가능한 원인:**\n"
                "- 해당 연도 사업보고서 미공시 (특히 가장 최근 연도)\n"
                "- 외감 비상장사로 API 미제공\n"
                "- 결산월 변경 등으로 보고서 형식이 비표준\n"
                "- 회사가 해당 연도에 존재하지 않음 (설립 이전)"
            )

        # EBITDA 산출 한계 명시
        st.caption(
            "⚠️ **EBITDA 산출 방식**: 영업이익 + 현금흐름표상 (유형자산감가상각비 + 무형자산상각비 + 사용권자산상각비). "
            "주석에만 표기된 상각비는 누락되며, 일회성/영업외 손익 조정은 반영되지 않습니다. "
            "정확한 EBITDA는 PDF 검증 단계에서 확정 필요."
        )

        # 엑셀 다운로드
        st.markdown("### 📥 엑셀 다운로드")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            template_df.to_excel(writer, sheet_name="5개년_요약", index=False)
            # 원본 데이터 (원 단위)
            raw_rows = []
            for y in years:
                d = yearly_data.get(y, {})
                row = {"연도": y}
                for key in TARGET_ACCOUNTS.keys():
                    row[key] = d.get(key)
                raw_rows.append(row)
            pd.DataFrame(raw_rows).to_excel(writer, sheet_name="원본_원단위", index=False)
            # 회사정보
            pd.DataFrame([selected_corp]).to_excel(writer, sheet_name="회사정보", index=False)

        st.download_button(
            label="📥 엑셀 다운로드 (5개년 요약 + 원본)",
            data=output.getvalue(),
            file_name=f"{corp_name}_{start_year}_{end_year}_{fs_div_target}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        # 디버그용
        with st.expander("🔬 연도별 추출 원본값 (원 단위, 검증용)"):
            debug_rows = []
            for y in years:
                d = yearly_data.get(y, {})
                row = {"연도": y}
                for key in TARGET_ACCOUNTS.keys():
                    v = d.get(key)
                    row[key] = f"{v:,}" if v is not None else "N/A"
                debug_rows.append(row)
            st.dataframe(pd.DataFrame(debug_rows), use_container_width=True, hide_index=True)

        st.divider()
        st.info(
            "🔜 **다음 단계 (개발 예정): PDF 자동 검증**\n\n"
            "각 연도 사업보고서의 감사보고서/연결감사보고서 PDF를 자동 다운로드하고, "
            "본문 표에서 동일 항목을 재추출하여 API 값과 대조합니다. "
            "불일치 시 경고와 함께 PDF 페이지 미리보기를 제공합니다."
        )

# ====================================================================
# 9) 사이드바 가이드
# ====================================================================
st.sidebar.divider()
with st.sidebar.expander("ℹ️ 사용 가이드"):
    st.markdown(
        "**순서**\n"
        "1. 회사명/코드 입력 → 회사 검색\n"
        "2. 동명회사 중 정확한 회사 선택\n"
        "3. 5개년 템플릿 추출 클릭\n"
        "4. 결과 확인 → 엑셀 다운로드\n\n"
        "**주의사항**\n"
        "- 비상장 외감기업(E)은 API 데이터 부재 가능\n"
        "- 금융업은 finstate_all 제외\n"
        "- EBITDA는 추정치, PDF 검증 필수\n"
        "- 총차입금은 BS의 4개 계정 합산"
    )

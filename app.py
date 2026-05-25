"""
DART 재무정보 추출 에이전트 v8
- PE 5개년 재무요약 템플릿
- 상장사: finstate_all (XBRL) 우선
- 외감 비상장사: 감사보고서 sub_docs HTML viewer 직접 파싱
- 회사 검색: corp_code.xml 직접 조회 fallback (사명 변경 대응)
- v7: D&A 주석 fallback (CF 통합 표시 케이스 대응)
- v8: 첫 연도 매출 성장률 계산용 직전년 매출 추출 + 총차입금 None/0 구분
"""

import io
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from typing import Optional, Dict, List, Tuple

import pandas as pd
import requests
import streamlit as st
import OpenDartReader

# ====================================================================
# 0) 페이지 설정
# ====================================================================
st.set_page_config(
    page_title="DART 재무분석",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ----- 전역 스타일 -----
st.markdown(
    """
    <style>
    /* 상단 헤더 */
    .hpe-header h2 {
        margin: 0 0 2px 0;
        font-weight: 700;
        color: #1E3D6B;
    }
    .hpe-sub {
        color: #888;
        font-size: 0.9rem;
        margin-bottom: 14px;
    }
    /* 섹션 타이틀 */
    .hpe-section {
        font-size: 1.0rem;
        font-weight: 700;
        color: #1E3D6B;
        margin: 18px 0 6px 0;
        padding-bottom: 4px;
        border-bottom: 2px solid #d0d7e2;
    }
    /* 결과 컨테이너 카드 틀 */
    .hpe-card {
        border: 1px solid #d0d7e2;
        border-radius: 10px;
        padding: 14px 18px;
        background: #f7f9fc;
        margin-bottom: 12px;
    }
    /* 하단 디버그 영역 차분 */
    .hpe-debug-header {
        color: #888;
        font-size: 0.85rem;
        font-weight: 600;
        margin-top: 24px;
        padding-top: 14px;
        border-top: 1px dashed #d0d7e2;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ----- 헤더 -----
st.markdown(
    """
    <div class="hpe-header">
        <h2>📈 DART 재무분석 툴</h2>
        <div class="hpe-sub">
            Highland PE · 내부 전용 | 출처: DART(dart.fss.or.kr)
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)
st.divider()

# ====================================================================
# 1) 상수
# ====================================================================
REPRT_CODE_ANNUAL = "11011"  # 사업보고서

# 손익/재무상태/CF 계정 키워드
# 키는 `_` 접두사일 경우 내부 계산용 보조 항목 (구성표에는 노출 가능)
ACCOUNT_KEYWORDS = {
    "매출액": ["매출액", "수익(매출액)", "영업수익", "매출", "수익"],
    "영업이익": ["영업이익", "영업이익(손실)", "영업손실"],
    "당기순이익": ["당기순이익", "당기순이익(손실)", "당기순손익", "분기순이익", "반기순이익"],
    "자산총계": ["자산총계", "자산 총계"],
    "현금성자산": ["현금및현금성자산", "현금 및 현금성자산"],
    "부채총계": ["부채총계", "부채 총계"],
    "자본총계": ["자본총계", "자본 총계"],

    # ----- Cash & Cash Equivalents 구성 (valuation Net Debt 차감항) -----
    "_단기금융상품": ["단기금융상품"],
    "_단기투자자산": ["단기투자자산", "단기투자증권"],
    "_당기손익공정가치_유동": [
        "당기손익-공정가치측정금융자산",  # 명칭 정확 매칭용
    ],
    "_매도가능금융자산_유동": ["매도가능금융자산", "매도가능증권"],
    "_기타포괄손익공정가치_유동": [
        "기타포괄손익-공정가치측정금융자산",
    ],
    "_장기금융상품": ["장기금융상품", "장기성예금"],
    # K-IFRS 외감사 통합 계정 "기타유동금융자산"은 통째로 합산하지 않는다 (v14).
    # 정기예금/단기금융상품 등 cash-like 항목만 주석 분해(parse_other_fa_breakdown)로
    # 별도 추출해 Cash 합계에 포함. 비유동 통합 계정(기타비유동금융자산)은
    # 통상 보증금·대여금이 대부분이므로 valuation 기준 cash 후보로 다루지 않는다.
    # 본 항목은 BS 디버그 표시용으로만 유지.
    "_기타유동금융자산_표시용": ["기타유동금융자산"],

    # ----- Gross Debt 구성 (valuation 가산항) -----
    "_단기차입금": ["단기차입금"],
    "_유동성장기부채": ["유동성장기부채", "유동성장기차입금", "유동성사채"],
    "_장기차입금": ["장기차입금"],
    "_사채": ["사채"],
    "_유동리스부채": ["유동리스부채", "유동성리스부채"],
    "_비유동리스부채": ["비유동리스부채", "장기리스부채"],
    "_전환사채": ["전환사채"],
    "_신주인수권부사채": ["신주인수권부사채"],
    "_교환사채": ["교환사채"],

    # ----- EBITDA 가산 (추수설명: BS 아닌 CF/주석) -----
    "_유형자산감가상각비": ["감가상각비", "유형자산감가상각비"],
    "_무형자산상각비": ["무형자산상각비", "무형자산상각"],
    "_사용권자산상각비": ["사용권자산상각비"],
}

# XBRL용 재무제표 구분
SJ_DIV_MAP = {
    "매출액": ["IS", "CIS"],
    "영업이익": ["IS", "CIS"],
    "당기순이익": ["IS", "CIS"],
    "자산총계": ["BS"],
    "현금성자산": ["BS"],
    "부채총계": ["BS"],
    "자본총계": ["BS"],
    "_단기금융상품": ["BS"],
    "_단기투자자산": ["BS"],
    "_당기손익공정가치_유동": ["BS"],
    "_매도가능금융자산_유동": ["BS"],
    "_기타포괄손익공정가치_유동": ["BS"],
    "_장기금융상품": ["BS"],
    "_기타유동금융자산_표시용": ["BS"],
    "_단기차입금": ["BS"],
    "_유동성장기부채": ["BS"],
    "_장기차입금": ["BS"],
    "_사채": ["BS"],
    "_유동리스부채": ["BS"],
    "_비유동리스부채": ["BS"],
    "_전환사채": ["BS"],
    "_신주인수권부사채": ["BS"],
    "_교환사채": ["BS"],
    "_유형자산감가상각비": ["CF"],
    "_무형자산상각비": ["CF"],
    "_사용권자산상각비": ["CF"],
}

# HTML 파싱용 재무제표 매핑 (어느 표에서 찾을지)
STATEMENT_OF = {
    "매출액": "IS",
    "영업이익": "IS",
    "당기순이익": "IS",
    "자산총계": "BS",
    "현금성자산": "BS",
    "부채총계": "BS",
    "자본총계": "BS",
    "_단기금융상품": "BS",
    "_단기투자자산": "BS",
    "_당기손익공정가치_유동": "BS",
    "_매도가능금융자산_유동": "BS",
    "_기타포괄손익공정가치_유동": "BS",
    "_장기금융상품": "BS",
    "_기타유동금융자산_표시용": "BS",
    "_단기차입금": "BS",
    "_유동성장기부채": "BS",
    "_장기차입금": "BS",
    "_사채": "BS",
    "_유동리스부채": "BS",
    "_비유동리스부채": "BS",
    "_전환사채": "BS",
    "_신주인수권부사채": "BS",
    "_교환사채": "BS",
    "_유형자산감가상각비": "CF",
    "_무형자산상각비": "CF",
    "_사용권자산상각비": "CF",
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
# 3) 유틸
# ====================================================================
# DART 공시 관행상 "-" / "–" / "—" 등은 "해당사항 없음 = 0"을 뜻함.
# 빈 셀(NaN, "")과 구분해서 빈 셀은 None, 대시는 0으로 처리.
# (DART HTML 파싱 시 은산 2024 장기차입금 같은 "일부 연도만 0" 케이스를 정확히 는기 위함)
_DASH_TOKENS = {"-", "–", "—", "−", "ー"}


def to_number(x) -> Optional[int]:
    if x is None or pd.isna(x):
        return None
    s = str(x).strip().replace(",", "").replace(" ", "").replace("\xa0", "")
    if s in ("", "nan", "None"):
        return None
    # 대시 표기(“-” 외 유니코드 변종 포함)는 0으로 해석
    if s in _DASH_TOKENS:
        return 0
    try:
        if s.startswith("(") and s.endswith(")"):
            s = "-" + s[1:-1]
        return int(float(s))
    except (ValueError, TypeError):
        return None


def to_eokwon(v: Optional[int], unit_scale: int = 1) -> Optional[float]:
    if v is None:
        return None
    return round(v * unit_scale / 100_000_000)


def format_eokwon(v: Optional[float]) -> str:
    if v is None:
        return "N/A"
    return f"{int(v):,}" if v >= 0 else f"({int(abs(v)):,})"


# ----- v15: 일반 단위 시스템 (백만원/억원/십억원) -----
# UNIT_SPECS: 표시 단위별 나눐수(divisor in 원 단위) 정의.
# all값은 원 점수(int)으로 계산 후 divisor로 나눠 표시.
# 소수점 자리수: 십억원은 1자리, 나머지 정수.
UNIT_SPECS = {
    "백만원": {"divisor": 1_000_000,    "decimals": 0, "suffix": "백만원"},
    "억원":   {"divisor": 100_000_000,  "decimals": 0, "suffix": "억원"},
    "십억원": {"divisor": 1_000_000_000,"decimals": 1, "suffix": "십억원"},
}


def to_unit(v_won: Optional[float], unit_label: str = "억원") -> Optional[float]:
    """원 단위 값을 대상 표시 단위로 환산.
    내부 추출값 v·scale 곱에 이미 단위가 적용된 원 값을 입력으로 기대.
    """
    if v_won is None:
        return None
    spec = UNIT_SPECS.get(unit_label, UNIT_SPECS["억원"])
    out = v_won / spec["divisor"]
    if spec["decimals"] == 0:
        return round(out)
    return round(out, spec["decimals"])


def format_unit(v: Optional[float], unit_label: str = "억원") -> str:
    """단위 환산된 값을 표시 문자열로 포맷. 음수는 괄호로."""
    if v is None:
        return "N/A"
    spec = UNIT_SPECS.get(unit_label, UNIT_SPECS["억원"])
    dec = spec["decimals"]
    if dec == 0:
        body = f"{int(round(v)):,}"
        return body if v >= 0 else f"({int(round(abs(v))):,})"
    fmt = f"{{:,.{dec}f}}"
    body = fmt.format(v)
    return body if v >= 0 else "(" + fmt.format(abs(v)) + ")"


def won_value(v_raw: Optional[int], unit_scale: int = 1) -> Optional[int]:
    """raw값·scale→원 단위. None은 유지."""
    if v_raw is None:
        return None
    return int(v_raw) * int(unit_scale)


def format_pct(v: Optional[float]) -> str:
    if v is None:
        return "N/A"
    if v < 0:
        return f"({abs(v):.1f}%)"
    return f"{v:.1f}%"


def normalize_account_name(name: str) -> str:
    """계정명 정규화: 로마숫자/번호/공백/주석표시 제거."""
    if name is None:
        return ""
    s = str(name).strip()
    s = re.sub(r"^[\dIVXⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+\.\s*", "", s)
    s = re.sub(r"^[\(\[]\s*\d+\s*[\)\]]\s*", "", s)
    s = re.sub(r"\(주석.*?\)", "", s)
    s = s.replace(" ", "").replace("　", "")
    return s


# ====================================================================
# 4) XBRL 기반 추출 (상장사)
# ====================================================================
def find_in_xbrl(df: pd.DataFrame, keywords: list, sj_filter: list) -> Optional[pd.Series]:
    if df is None or df.empty:
        return None
    work = df.copy()
    if "sj_div" in work.columns:
        work = work[work["sj_div"].isin(sj_filter)]
    if work.empty:
        return None

    for kw in keywords:
        m = work[work["account_nm"] == kw]
        if not m.empty:
            return m.iloc[0]

    pattern = "|".join([re.escape(k) for k in keywords])
    partial = work[work["account_nm"].astype(str).str.contains(pattern, na=False, regex=True)]
    if not partial.empty:
        partial = partial.copy()
        partial["_len"] = partial["account_nm"].astype(str).str.len()
        return partial.sort_values("_len").iloc[0]
    return None


@st.cache_data(ttl=600, show_spinner=False)
def fetch_xbrl_finstate(_dart, corp_code: str, year: int, fs_div: str) -> Optional[pd.DataFrame]:
    try:
        df = _dart.finstate_all(corp_code, year, reprt_code=REPRT_CODE_ANNUAL, fs_div=fs_div)
        if df is None or (hasattr(df, "empty") and df.empty):
            return None
        return df
    except Exception:
        return None


def extract_from_xbrl(df: pd.DataFrame, year_offset: int = 0) -> Dict[str, Optional[int]]:
    if df is None or df.empty:
        return {key: None for key in ACCOUNT_KEYWORDS.keys()}
    amount_col = {0: "thstrm_amount", 1: "frmtrm_amount", 2: "bfefrmtrm_amount"}[year_offset]
    result = {}
    for key, keywords in ACCOUNT_KEYWORDS.items():
        sj = SJ_DIV_MAP.get(key, ["IS", "BS", "CIS", "CF"])
        row = find_in_xbrl(df, keywords, sj)
        result[key] = to_number(row.get(amount_col)) if row is not None else None
    return result


# ====================================================================
# 5) 외감 감사보고서 HTML 파싱 (sub_docs viewer)
# ====================================================================
@st.cache_data(ttl=3600, show_spinner=False)
def find_external_audit_reports(_dart, corp_code: str, year: int) -> List[Dict]:
    """
    외감사: kind='F'에서 해당 회계연도 감사보고서/연결감사보고서 검색.
    감사보고서는 다음 해 1~6월에 주로 제출됨.
    반환: 우선순위 정렬된 리스트 (연결감사보고서 → 감사보고서)
    """
    try:
        # 회계연도 다음 해 전체 + 그 다음 해 상반기까지 (수정 보고 대비)
        start = f"{year+1}-01-01"
        end = f"{year+2}-06-30"

        reports = _dart.list(corp_code, start=start, end=end, kind="F")
        if reports is None or reports.empty:
            return []

        # 연도 매칭: report_nm에 "(YYYY.MM)" 패턴이 있고 YYYY가 해당 연도여야 함
        def matches_year(report_nm):
            nm = str(report_nm)
            m = re.search(r"\((\d{4})\.\d{2}\)", nm)
            if m:
                return int(m.group(1)) == year
            return False

        reports = reports.copy()
        reports["_matches"] = reports["report_nm"].apply(matches_year)
        reports = reports[reports["_matches"]]
        if reports.empty:
            return []

        def priority(report_nm):
            nm = str(report_nm)
            if "연결감사보고서" in nm:
                return 0
            if "감사보고서" in nm:
                return 1
            return 99

        reports["_prio"] = reports["report_nm"].apply(priority)
        reports = reports[reports["_prio"] < 99]
        reports = reports.sort_values(["_prio", "rcept_dt"], ascending=[True, False])

        return [
            {
                "rcept_no": r["rcept_no"],
                "report_nm": r["report_nm"],
                "rcept_dt": r.get("rcept_dt"),
                "is_consolidated": "연결" in str(r["report_nm"]),
            }
            for _, r in reports.iterrows()
        ]
    except Exception:
        return []


@st.cache_data(ttl=3600, show_spinner=False)
def get_sub_docs(_dart, rcept_no: str) -> Optional[pd.DataFrame]:
    try:
        df = _dart.sub_docs(rcept_no)
        if df is None or df.empty:
            return None
        return df
    except Exception:
        return None


def find_statement_url(sub_docs: pd.DataFrame, keyword: str) -> Optional[str]:
    """sub_docs에서 'BS/IS/CF' 키워드를 포함한 첫 viewer URL."""
    if sub_docs is None or sub_docs.empty:
        return None
    norm = sub_docs["title"].astype(str).str.replace(" ", "").str.replace("　", "")
    matched = sub_docs[norm.str.contains(keyword, na=False)]
    if matched.empty:
        return None
    return matched.iloc[0]["url"]


@st.cache_data(ttl=3600, show_spinner=False)
def parse_html_statement(url: str) -> Optional[pd.DataFrame]:
    """DART viewer URL을 pd.read_html로 파싱. 표 1(본문)을 반환."""
    try:
        tables = pd.read_html(url)
        if len(tables) < 2:
            return None
        return tables[1]
    except Exception:
        return None


def detect_unit_from_header(url: str) -> int:
    """표 0의 헤더에서 '단위' 텍스트를 찾아 스케일 결정. 1=원, 1000=천원, 1000000=백만원."""
    try:
        tables = pd.read_html(url)
        if not tables:
            return 1
        header_text = " ".join(tables[0].fillna("").astype(str).values.flatten().tolist())
        if re.search(r"단위\s*[:：]?\s*백만\s*원", header_text):
            return 1_000_000
        if re.search(r"단위\s*[:：]?\s*천\s*원", header_text):
            return 1_000
        return 1
    except Exception:
        return 1


def extract_year_columns(df: pd.DataFrame) -> Dict[int, List[str]]:
    """
    당기/전기/전전기 컬럼 매핑.
    DART HTML 표는 컬럼명에 '(당)/(전)/(전전)' 표기를 사용하며,
    같은 컬럼이 2개씩 나타남 (원본 + .1) — 둘 다 후보로 둔다.
    """
    cols = [str(c) for c in df.columns]
    result = {0: [], 1: [], 2: []}
    for c in cols:
        cn = c.replace(" ", "")
        if "(전전)" in cn or "전전기" in cn:
            result[2].append(c)
        elif "(당)" in cn or "당기" in cn or "(당기)" in cn:
            result[0].append(c)
        elif "(전)" in cn or "전기" in cn or "(전기)" in cn:
            result[1].append(c)
    return result


def extract_value_from_row(df: pd.DataFrame, row_idx: int, target_cols: List[str]) -> Optional[int]:
    """한 행의 target_cols 중 첫 유효 숫자를 반환 (원본 → .1 순)."""
    for col in target_cols:
        v = to_number(df.iloc[row_idx][col])
        if v is not None:
            return v
    return None


# 구체 항목이 더 일반 항목 키워드에 오인식 되는 것을 차단.
# 예: "사채" 키워드가 "전환사채" 행에 부분매칭되지 않도록.
KEYWORD_EXCLUDES = {
    "사채": ["전환사채", "신주인수권부사채", "교환사채", "유동성사채"],
    "장기차입금": ["유동성장기차입금"],
    "장기리스부채": [],
    "단기금융상품": ["단기금융상품(사용제한)"],  # 필요시
    "장기금융상품": ["장기금융상품(사용제한)"],
}


def find_account_in_html(df: pd.DataFrame, keywords: list, year_offset: int = 0) -> Optional[int]:
    """첫 컬럼(과목)에서 매칭 행 검색 → 해당 연도 값 반환.
    정확 매칭 우선, 실패 시 부분 매칭.
    KEYWORD_EXCLUDES에 등록된 키워드는 해당 제외 항목을 포함한 행에 부분 매칭되지 않는다.
    """
    if df is None or df.empty:
        return None
    year_cols = extract_year_columns(df)
    target_cols = year_cols.get(year_offset, [])
    if not target_cols:
        return None
    first_col = df.columns[0]

    # 1차: 정확 매칭
    for idx in range(len(df)):
        cell = df.iloc[idx][first_col]
        if pd.isna(cell):
            continue
        norm = normalize_account_name(str(cell))
        for kw in keywords:
            if norm == normalize_account_name(kw):
                v = extract_value_from_row(df, idx, target_cols)
                if v is not None:
                    return v

    # 2차: 부분 매칭 (가장 짧은 계정명 우선) + 제외 규칙 적용
    candidates = []
    for idx in range(len(df)):
        cell = df.iloc[idx][first_col]
        if pd.isna(cell):
            continue
        norm = normalize_account_name(str(cell))
        for kw in keywords:
            nkw = normalize_account_name(kw)
            if not nkw or nkw not in norm:
                continue
            # 제외 규칙 체크: 키워드별 exclude 항목이 행 이름에 있으면 스킵
            excludes = KEYWORD_EXCLUDES.get(kw, [])
            if any(normalize_account_name(ex) in norm for ex in excludes):
                continue
            candidates.append((len(norm), idx, str(cell)))
            break
    candidates.sort()
    for _, idx, _ in candidates:
        v = extract_value_from_row(df, idx, target_cols)
        if v is not None:
            return v
    return None


def classify_statement_table(t):
    """표 내용 지문으로 BS/IS/CF/EQ 중 어떤 재무제표인지 판별."""
    if t is None or t.shape[0] < 3 or t.shape[1] < 3:
        return None
    first_col = t.iloc[:, 0].astype(str).fillna("")
    text = " ".join(first_col.tolist())
    text_norm = text.replace(" ", "").replace("　", "").lower()

    has_revenue = bool(re.search(r"매출액|영업수익", text_norm))
    has_op_inc = "영업이익" in text_norm or "영업손실" in text_norm
    if has_revenue and has_op_inc:
        return "IS"

    has_assets = ("자산총계" in text_norm) or bool(re.search(r"자\s*산\s*총\s*계", text))
    has_liab = ("부채총계" in text_norm) or bool(re.search(r"부\s*채\s*총\s*계", text))
    has_equity = ("자본총계" in text_norm) or bool(re.search(r"자\s*본\s*총\s*계", text))
    if has_assets and has_liab and has_equity:
        return "BS"

    has_op_cf = "영업활동" in text_norm and "현금흐름" in text_norm
    has_inv_cf = "투자활동" in text_norm
    has_fin_cf = "재무활동" in text_norm
    if has_op_cf and (has_inv_cf or has_fin_cf):
        return "CF"
    return None


@st.cache_data(ttl=3600, show_spinner=False)
def parse_combined_statements(url: str) -> Dict[str, Optional[pd.DataFrame]]:
    """
    (첨부)재무제표 같은 통합 페이지에서 모든 표를 읽고 BS/IS/CF로 자동 분류.
    반환: {'BS': df, 'IS': df, 'CF': df}
    """
    result = {"BS": None, "IS": None, "CF": None}
    try:
        tables = pd.read_html(url)
    except Exception:
        return result
    for t in tables:
        cls = classify_statement_table(t)
        if cls and result[cls] is None:
            result[cls] = t
    return result


def detect_unit_from_tables(tables: list) -> int:
    """파싱된 표 리스트의 상단 헤더에서 단위 감지."""
    text_blob = ""
    for t in tables[:5]:
        try:
            text_blob += " ".join(t.fillna("").astype(str).values.flatten().tolist())
        except Exception:
            continue
    if re.search(r"단위\s*[:：]?\s*백만\s*원", text_blob):
        return 1_000_000
    if re.search(r"단위\s*[:：]?\s*천\s*원", text_blob):
        return 1_000
    return 1


def find_da_from_notes(tables: list) -> Dict[str, Optional[Dict]]:
    """
    감사보고서의 모든 표(통합 페이지/주석 페이지)에서 D&A 후보를 찾는다.
    '비용성격별 분류' 등의 주석 표는 통상 '구분 | 당기 | 전기' 형태이고,
    유형자산 변동표(기초/증가/감소/기말)와는 구분된다.
    동일 키에 여러 후보가 있으면 (판관비 분석 vs 비용성격별 분류) 가장 큰 값을 선택.
    반환: {유형자산감가상각비/무형자산상각비/사용권자산상각비/감가상각비_합산: {value_in_won, ...}}
    """
    targets = {
        "감가상각비": ["감가상각비"],
        "감가상각비_합산": ["감가상각비와무형자산상각비", "감가상각비및무형자산상각비"],
        "무형자산상각비": ["무형자산상각비"],
        "사용권자산상각비": ["사용권자산상각비", "사용권자산감가상각비"],
    }

    # 표별 단위 추정 (직전 3개 표 텍스트에서 '단위' 탐색)
    unit_per_table = {}
    for i in range(len(tables)):
        local_text = ""
        for j in range(max(0, i - 3), i + 1):
            if j >= len(tables):
                continue
            try:
                local_text += " ".join(tables[j].fillna("").astype(str).values.flatten().tolist())
            except Exception:
                continue
        if re.search(r"단위\s*[:：]?\s*백만\s*원", local_text):
            unit_per_table[i] = 1_000_000
        elif re.search(r"단위\s*[:：]?\s*천\s*원", local_text):
            unit_per_table[i] = 1_000
        else:
            unit_per_table[i] = 1

    def col_norm(c):
        return str(c).replace(" ", "").replace("　", "")

    candidates = {k: [] for k in targets.keys()}

    for i, t in enumerate(tables):
        if t is None or t.shape[0] < 2 or t.shape[1] < 2:
            continue
        cols = [col_norm(c) for c in t.columns]
        # 당기/전기 컬럼이 있어야 함
        has_period = any(("당기" in c or ("당" in c and "기" in c)) for c in cols)
        if not has_period:
            continue
        # 변동 분석 표 (기초/증가/감소/기말) 제외
        col_text = " ".join(cols)
        if "기초" in col_text and "기말" in col_text:
            continue

        first_col = t.columns[0]
        for idx in range(len(t)):
            cell = t.iloc[idx][first_col]
            if pd.isna(cell):
                continue
            cell_norm = str(cell).replace(" ", "").replace("　", "")

            for key, patterns in targets.items():
                if cell_norm in patterns:
                    # 당기 컬럼에서 값 추출
                    for c in t.columns:
                        cn = col_norm(c)
                        if "당기" in cn or "당)기" in cn or ("당" in cn and "기" in cn and "전" not in cn):
                            v = to_number(t.iloc[idx][c])
                            if v is not None:
                                candidates[key].append({
                                    "table_idx": i,
                                    "row_idx": idx,
                                    "value": v,
                                    "unit_scale": unit_per_table[i],
                                    "value_in_won": v * unit_per_table[i],
                                    "column": str(c),
                                })
                                break
                    break

    # 각 키별 최대값 (가장 포괄적인 합계) 선택
    result = {}
    for key, items in candidates.items():
        if items:
            result[key] = max(items, key=lambda x: x["value_in_won"])
        else:
            result[key] = None
    return result


def fetch_all_tables_from_audit(sub_docs: pd.DataFrame) -> list:
    """감사보고서의 통합 재무제표 페이지 + 주석 페이지의 모든 표 합집합."""
    if sub_docs is None or sub_docs.empty:
        return []
    tables = []
    norm = sub_docs["title"].astype(str).str.replace(" ", "").str.replace("　", "")
    # 통합 재무제표
    matched = sub_docs[norm.str.contains("재무제표", na=False)]
    for _, row in matched.head(2).iterrows():
        try:
            tables.extend(pd.read_html(row["url"]))
        except Exception:
            continue
    # 주석
    matched = sub_docs[norm == "주석"]
    for _, row in matched.head(2).iterrows():
        try:
            tables.extend(pd.read_html(row["url"]))
        except Exception:
            continue
    return tables


# ====================================================================
# v14: "기타유동금융자산" 주석 분해
# K-IFRS 외감사의 통합 계정 "기타유동금융자산" 안에 정기예금/단기금융상품/MMF 등
# cash-like 항목과 단기대여금/보증금 등 비-cash 항목이 섞일 수 있음.
# v13 이전: 통째 Cash 합산 (대여금/보증금 포함) → 잘못된 현금성자산 과대계산.
# v14: 주석 표에서 항목별 분해, cash-like 화이트리스트만 Cash에 합산.
# "기타비유동금융자산"은 사용자 지시에 따라 처리 대상에서 제외 (항상 보증금 위주로 간주).
# ====================================================================
_CASH_LIKE_PATTERNS = [
    "정기예금", "단기금융상품", "단기금융자산",
    "MMF", "머니마켓펀드", "양도성예금증서", "환매조건부채권",
    "금융기관예치금", "정기예적금", "특정금전신탁",
]
# 비-cash 대상: 자산성 항목만 잔의. 예수보증금(부채)은 제외.
_NON_CASH_PATTERNS = [
    "임차보증금", "임대보증금", "보증금",
    "단기대여금", "장기대여금", "장단기대여금", "대여금",
    "파생상품", "파생금융자산",
    "출자금", "지분증권",
]
# 부채/비금융자산 표로 주적되는 키워드 — 자산 분해표가 아닌지 배제
_LIABILITY_INDICATORS = [
    "리스부채", "차입금", "사채", "미지급금", "미지급비용",
    "매입채무", "예수금", "예수보증금",
    "충당부채", "판매보증충당부채",
    "원/부재료", "급여", "퇴직급여", "복리후생",
]


def _classify_other_fa_item(name: str) -> str:
    """항목명 → 'cash_like' / 'non_cash' / 'unknown'"""
    n = str(name).replace(" ", "").replace("　", "")
    for p in _CASH_LIKE_PATTERNS:
        if p in n:
            return "cash_like"
    for p in _NON_CASH_PATTERNS:
        if p in n:
            return "non_cash"
    return "unknown"


def _norm_cell(s) -> str:
    return str(s).replace(" ", "").replace("　", "").replace("\xa0", "")


def _detect_table_unit_local(tables: list, idx: int) -> int:
    """표 idx 주변(이전 3개) 텍스트에서 '단위: 천원/백만원' 감지."""
    blob = ""
    for j in range(max(0, idx - 3), idx + 1):
        if j >= len(tables):
            continue
        try:
            blob += " ".join(tables[j].fillna("").astype(str).values.flatten().tolist())
        except Exception:
            continue
    if re.search(r"단위\s*[:：]?\s*백만\s*원", blob):
        return 1_000_000
    if re.search(r"단위\s*[:：]?\s*천\s*원", blob):
        return 1_000
    return 1


def _identify_breakdown_table(t: pd.DataFrame):
    """
    주석 표가 "기타유동금융자산 분해표"인지 판정 + 컬럼 매핑 반환.
    지원 패턴:
      A) 심플 헤더 4컬럼: [구분(유동/비유동), 구분(항목명), 당기말, 전기말]
         - col0에 "유동" 및 "비유동" 모두 등장
         - val_col_idx = 2 (당기말)
         - liquidity는 col0에서, name은 col1에서
      B) 멀티헤더 5컬럼: [구분, (당기말,유동), (당기말,비유동), (전기말,유동), (전기말,비유동)]
         - val_col_curr_idx, val_col_non_idx = 1, 2
         - liquidity는 컴럼 헤더에서, name은 col0에서
    반환: dict {pattern, name_col_idx, liq_col_idx_or_value, val_curr_idx, val_non_idx}
         또는 None.
    """
    if t is None or t.shape[0] < 1 or t.shape[1] < 2:
        return None

    # 패턴 B: 멀티헤더 + "당기말" + "유동/비유동"
    if any(isinstance(c, tuple) for c in t.columns):
        curr_curr = None  # 당기말 유동
        curr_non = None   # 당기말 비유동
        for i, c in enumerate(t.columns):
            if not isinstance(c, tuple):
                continue
            joined = "".join(str(x) for x in c).replace(" ", "")
            if "당기말" in joined and "비유동" in joined and curr_non is None:
                curr_non = i
            elif "당기말" in joined and "유동" in joined and curr_curr is None:
                curr_curr = i
        if curr_curr is not None or curr_non is not None:
            return {
                "pattern": "B",
                "name_col": 0,
                "val_curr_idx": curr_curr,
                "val_non_idx": curr_non,
            }

    # 패턴 A: col0에 유동/비유동 모두 등장, 컬럼 4개 이상
    try:
        col0 = t.iloc[:, 0].fillna("").astype(str).tolist()
    except Exception:
        return None
    col0_blob = " ".join(col0)
    if ("유동" in col0_blob) and ("비유동" in col0_blob) and t.shape[1] >= 3:
        # val_col은 세번째 컬럼 (당기말). 멀티헤더 아닄 단순헤더.
        # 안전: 컬럼 이름에 "당기" 포함된 컬럼 이동
        val_idx = 2  # default
        for i, c in enumerate(t.columns):
            cn = str(c).replace(" ", "")
            if "당기" in cn and i >= 2:
                val_idx = i
                break
        return {
            "pattern": "A",
            "name_col": 1,
            "liq_col": 0,
            "val_curr_idx": val_idx,  # 당기말 (유동이면 유동값, 비유동이면 비유동값)
            "val_non_idx": None,
        }
    return None


def parse_other_fa_breakdown(sub_docs: pd.DataFrame, debug: bool = False) -> Dict:
    """
    감사보고서 주석 페이지의 모든 표에서 "기타유동금융자산" 분해표를 수집·분류.
    단위는 표별 감지(통상 천원). 반환값은 원 단위.

    반환:
      {
        "cash_like":     {항목명: 원단위값(int)},  # 유동 항목만 집계 (Cash 합산용)
        "non_cash_like": {항목명: 원단위값(int)},  # 참고용 (표시만 함)
        "details":       [행별 메타...],
        "source_tables": [표인덱스...],
      }
    """
    out = {"cash_like": {}, "non_cash_like": {}, "details": [], "source_tables": []}
    if sub_docs is None or sub_docs.empty:
        return out
    norm = sub_docs["title"].astype(str).str.replace(" ", "").str.replace("　", "")
    notes = sub_docs[norm == "주석"]
    if notes.empty:
        return out

    try:
        tables = pd.read_html(notes.iloc[0]["url"])
    except Exception:
        return out

    for i, t in enumerate(tables):
        info = _identify_breakdown_table(t)
        if not info:
            continue
        # 부채 표 제외
        try:
            first_col_vals = t.iloc[:, info.get("name_col", 0)].fillna("").astype(str).tolist()
        except Exception:
            continue
        liability = False
        for v in first_col_vals:
            vn = _norm_cell(v)
            if any(p in vn for p in _LIABILITY_INDICATORS):
                liability = True
                break
        if liability:
            continue

        unit = _detect_table_unit_local(tables, i)

        # 행 순회
        table_matched = False
        for ridx in range(len(t)):
            try:
                name_cell = t.iloc[ridx, info["name_col"]]
            except Exception:
                continue
            if pd.isna(name_cell):
                continue
            name = str(name_cell).strip()
            if not name:
                continue
            nn = _norm_cell(name)
            if nn in ("합계", "총계", "소계", "구분"):
                continue
            if "합" in nn and "계" in nn:
                continue

            cat = _classify_other_fa_item(name)
            if cat == "unknown":
                continue

            # 유동성 판단
            liquidity = None
            if info["pattern"] == "A":
                try:
                    liq_cell = t.iloc[ridx, info["liq_col"]]
                    lq = _norm_cell(liq_cell) if not pd.isna(liq_cell) else ""
                except Exception:
                    lq = ""
                if "비유동" in lq:
                    liquidity = "non_current"
                elif "유동" in lq:
                    liquidity = "current"
                # 값: val_curr_idx 에서 동일
                v_curr_raw = t.iloc[ridx, info["val_curr_idx"]] if info["val_curr_idx"] is not None else None
                v_curr = to_number(v_curr_raw) if v_curr_raw is not None else None
                v_non = None
            else:  # 패턴 B
                v_curr_raw = t.iloc[ridx, info["val_curr_idx"]] if info["val_curr_idx"] is not None else None
                v_non_raw = t.iloc[ridx, info["val_non_idx"]] if info["val_non_idx"] is not None else None
                v_curr = to_number(v_curr_raw) if v_curr_raw is not None else None
                v_non = to_number(v_non_raw) if v_non_raw is not None else None
                if v_curr and not v_non:
                    liquidity = "current"
                elif v_non and not v_curr:
                    liquidity = "non_current"
                else:
                    liquidity = "both_or_none"

            # 원 단위 값 계산 (패턴 A는 v_curr가 해당 유동성 값)
            if info["pattern"] == "A":
                total = (v_curr or 0) * unit
            else:
                total = ((v_curr or 0) + (v_non or 0)) * unit

            # Cash 누적: 유동 cash_like만 Cash 합계에 포함.
            # (비유동 정기예금 같은 건 드물지만 존재 → 별도 처리 안 함.
            #  보수적으로 유동만 Cash, 비유동 cash_like는 non_cash_like로 차입)
            if cat == "cash_like":
                if info["pattern"] == "A":
                    if liquidity == "current":
                        out["cash_like"][name] = out["cash_like"].get(name, 0) + total
                    else:
                        # 비유동 cash_like는 보수적으로 참고만
                        out["non_cash_like"][name + " (비유동)"] = out["non_cash_like"].get(name + " (비유동)", 0) + total
                else:
                    # 패턴 B: 유동값·비유동값 분리
                    cur_total = (v_curr or 0) * unit
                    non_total = (v_non or 0) * unit
                    if cur_total > 0:
                        out["cash_like"][name] = out["cash_like"].get(name, 0) + cur_total
                    if non_total > 0:
                        out["non_cash_like"][name + " (비유동)"] = out["non_cash_like"].get(name + " (비유동)", 0) + non_total
            else:  # non_cash
                out["non_cash_like"][name] = out["non_cash_like"].get(name, 0) + total

            out["details"].append({
                "table_idx": i,
                "pattern": info["pattern"],
                "name": name,
                "category": cat,
                "liquidity": liquidity,
                "value_won": total,
                "v_curr": (v_curr or 0) * unit,
                "v_non": (v_non or 0) * unit if v_non is not None else 0,
                "unit": unit,
            })
            table_matched = True
        if table_matched and i not in out["source_tables"]:
            out["source_tables"].append(i)

    return out


def extract_external_audit_data(_dart, corp_code: str, year: int,
                                prefer_consolidated: bool = True) -> Dict:
    """
    외감 비상장사: 감사보고서 sub_docs에서 BS/IS/CF 직접 파싱.
    두 가지 sub_docs 구조 모두 대응:
      - 패턴 A: 재무상태표/손익계산서/현금흐름표가 각각 별도 항목
      - 패턴 B: '(첨부)재 무 제 표' 한 덩어리에 통합
    """
    reports = find_external_audit_reports(_dart, corp_code, year)
    if not reports:
        return {"data": None, "error": f"{year}년 외부감사 감사보고서 미발견"}

    if prefer_consolidated:
        chosen = next((r for r in reports if r["is_consolidated"]), reports[0])
    else:
        chosen = next((r for r in reports if not r["is_consolidated"]), reports[0])

    rcept_no = chosen["rcept_no"]
    sub_docs = get_sub_docs(_dart, rcept_no)
    if sub_docs is None:
        return {"data": None, "error": f"sub_docs 조회 실패 (rcept_no={rcept_no})",
                "rcept_no": rcept_no, "report_nm": chosen["report_nm"]}

    # 1차: 패턴 A - 항목 분리형
    urls = {
        "BS": find_statement_url(sub_docs, "재무상태표"),
        "IS": find_statement_url(sub_docs, "손익계산서"),
        "CF": find_statement_url(sub_docs, "현금흐름표"),
    }
    if not urls["IS"]:
        urls["IS"] = find_statement_url(sub_docs, "포괄손익계산서")

    dfs = {k: parse_html_statement(u) if u else None for k, u in urls.items()}
    unit_scale = detect_unit_from_header(urls["BS"]) if urls["BS"] else 1

    # 2차: 패턴 B - 통합 페이지 fallback
    used_combined = False
    if all(df is None for df in dfs.values()):
        combined_url = find_statement_url(sub_docs, "재무제표")
        if combined_url:
            used_combined = True
            classified = parse_combined_statements(combined_url)
            for k in ["BS", "IS", "CF"]:
                if dfs[k] is None and classified.get(k) is not None:
                    dfs[k] = classified[k]
            # 단위 감지 (전체 페이지 표에서)
            try:
                all_tables = pd.read_html(combined_url)
                unit_scale = detect_unit_from_tables(all_tables)
            except Exception:
                pass

    # 데이터 추출
    result_data = {}
    debug = []
    for key, keywords in ACCOUNT_KEYWORDS.items():
        stmt = STATEMENT_OF[key]
        df = dfs.get(stmt)
        val = find_account_in_html(df, keywords, year_offset=0)
        result_data[key] = val
        debug.append({
            "항목": key,
            "재무제표": stmt,
            "값(원단위)": val,
            "발견여부": "✓" if val is not None else "—",
        })

    # ============================================================
    # D&A fallback: CF에서 감가상각비/무형자산상각비/사용권자산상각비가 누락되면
    # 주석 페이지를 포함한 모든 표에서 자동 탐색.
    # CF에 통합 표시(예: "현금의 유출이 없는 비용등의 가산")만 있는 경우 대응.
    # 주의: 주석 표의 값은 통상 천원 단위 → unit_scale 별도 적용.
    # ============================================================
    da_keys = ["_유형자산감가상각비", "_무형자산상각비", "_사용권자산상각비"]
    da_missing = [k for k in da_keys if result_data.get(k) is None]
    da_fallback_used = False
    da_fallback_info = {}
    if da_missing:
        # 통합/주석 페이지 표 전체 수집
        all_audit_tables = fetch_all_tables_from_audit(sub_docs)
        if all_audit_tables:
            notes_da = find_da_from_notes(all_audit_tables)
            # 매핑: 주석 키 → result_data 키
            mapping = {
                "_유형자산감가상각비": "감가상각비",
                "_무형자산상각비": "무형자산상각비",
                "_사용권자산상각비": "사용권자산상각비",
            }
            for r_key, n_key in mapping.items():
                if result_data.get(r_key) is None:
                    info = notes_da.get(n_key)
                    if info is not None:
                        # 주석에서 가져온 값은 원 단위(value_in_won).
                        # result_data는 메인 unit_scale 기준으로 저장되어야
                        # 후속 to_eokwon 계산에서 일관성 유지됨.
                        # 따라서 원 단위값을 unit_scale로 나눠서 저장.
                        won_value = info["value_in_won"]
                        adjusted = int(won_value / unit_scale) if unit_scale else int(won_value)
                        result_data[r_key] = adjusted
                        da_fallback_info[r_key] = {
                            "source": f"주석 표 {info['table_idx']}",
                            "raw_value": info["value"],
                            "table_unit": info["unit_scale"],
                            "won_value": won_value,
                            "adjusted_for_main_scale": adjusted,
                        }
                        da_fallback_used = True
                        # debug 업데이트
                        for d in debug:
                            if d["항목"] == r_key:
                                d["값(원단위)"] = won_value
                                d["재무제표"] = f"주석(표{info['table_idx']})"
                                d["발견여부"] = "✓ (주석 fallback)"
                                break
            # 만약 개별 항목이 다 누락이고 '감가상각비_합산'만 있으면
            if all(result_data.get(k) is None for k in ["_유형자산감가상각비", "_무형자산상각비"]):
                combined = notes_da.get("감가상각비_합산")
                if combined is not None:
                    won_value = combined["value_in_won"]
                    adjusted = int(won_value / unit_scale) if unit_scale else int(won_value)
                    result_data["_유형자산감가상각비"] = adjusted
                    da_fallback_info["_유형자산감가상각비"] = {
                        "source": f"주석 표 {combined['table_idx']} (감가상각비+무형자산상각비 합산)",
                        "raw_value": combined["value"],
                        "table_unit": combined["unit_scale"],
                        "won_value": won_value,
                        "adjusted_for_main_scale": adjusted,
                    }
                    da_fallback_used = True

    # ============================================================
    # v14: "기타유동금융자산" 주석 분해 (cash-like / non-cash 자동 분류)
    # ============================================================
    other_fa_breakdown = {}
    try:
        other_fa_breakdown = parse_other_fa_breakdown(sub_docs)
    except Exception as e:
        other_fa_breakdown = {"cash_like": {}, "non_cash_like": {}, "details": [], "error": str(e)}

    return {
        "data": result_data,
        "unit_scale": unit_scale,
        "rcept_no": rcept_no,
        "report_nm": chosen["report_nm"],
        "is_consolidated": chosen["is_consolidated"],
        "debug": debug,
        "statement_urls": urls,
        "parsing_mode": "통합페이지" if used_combined else "항목분리",
        "da_fallback_used": da_fallback_used,
        "da_fallback_info": da_fallback_info,
        "other_fa_breakdown": other_fa_breakdown,
    }


# ====================================================================
# 6) 통합 수집 로직
# ====================================================================
def collect_multi_year(_dart, corp_code: str, corp_cls: str,
                       years: List[int], fs_div: str,
                       progress_callback=None) -> Tuple[Dict, Dict]:
    """
    상장사(Y/K/N): XBRL 3년 묶음 호출 → 누락 시 HTML 폴백 없음 (현재)
    외감(E): 매년 sub_docs HTML 파싱
    """
    yearly_data = {y: {} for y in years}
    yearly_meta = {y: {} for y in years}
    prefer_consolidated = (fs_div == "CFS")

    # ============ 상장사 ============
    if corp_cls in ("Y", "K", "N"):
        sorted_years = sorted(years, reverse=True)
        fetch_years = []
        covered = set()
        for y in sorted_years:
            if y in covered:
                continue
            fetch_years.append(y)
            covered.update([y, y - 1, y - 2])

        total = len(fetch_years)
        for idx, base_year in enumerate(fetch_years):
            if progress_callback:
                progress_callback(idx, total, f"XBRL {base_year}년")
            df = fetch_xbrl_finstate(_dart, corp_code, base_year, fs_div)
            if df is None:
                continue
            for offset in [0, 1, 2]:
                ty = base_year - offset
                if ty in years and not yearly_data[ty]:
                    yearly_data[ty] = extract_from_xbrl(df, year_offset=offset)
                    yearly_meta[ty] = {"source": "XBRL", "unit_scale": 1, "fs_div": fs_div}

        # 누락 연도는 외감 HTML 폴백 시도 (상장 → 외감으로 강등된 케이스)
        empty_years = [y for y in years
                       if not yearly_data[y] or all(v is None for v in yearly_data[y].values())]
        for i, y in enumerate(empty_years):
            if progress_callback:
                progress_callback(i, len(empty_years), f"HTML 폴백 {y}년")
            result = extract_external_audit_data(_dart, corp_code, y, prefer_consolidated)
            if result.get("data") and any(v is not None for v in result["data"].values()):
                yearly_data[y] = result["data"]
                yearly_meta[y] = {
                    "source": "HTML(외감 폴백)",
                    "unit_scale": result.get("unit_scale", 1),
                    "rcept_no": result.get("rcept_no"),
                    "report_nm": result.get("report_nm"),
                    "debug": result.get("debug"),
                    "statement_urls": result.get("statement_urls"),
                    "other_fa_breakdown": result.get("other_fa_breakdown", {}),
                }
            else:
                yearly_meta[y] = {"source": "FAILED", "error": result.get("error")}

    # ============ 외감 비상장사 ============
    else:  # E 등
        # 외감이라도 XBRL 제출했는지 한번 시도
        if progress_callback:
            progress_callback(0, len(years) + 1, "외감 XBRL 시도")
        df_test = fetch_xbrl_finstate(_dart, corp_code, max(years), fs_div)
        xbrl_works = df_test is not None and not df_test.empty

        if xbrl_works:
            sorted_years = sorted(years, reverse=True)
            fetch_years = []
            covered = set()
            for y in sorted_years:
                if y in covered:
                    continue
                fetch_years.append(y)
                covered.update([y, y - 1, y - 2])
            for idx, base_year in enumerate(fetch_years):
                if progress_callback:
                    progress_callback(idx, len(fetch_years), f"XBRL {base_year}년 (외감 XBRL)")
                df = fetch_xbrl_finstate(_dart, corp_code, base_year, fs_div)
                if df is None:
                    continue
                for offset in [0, 1, 2]:
                    ty = base_year - offset
                    if ty in years and not yearly_data[ty]:
                        yearly_data[ty] = extract_from_xbrl(df, year_offset=offset)
                        yearly_meta[ty] = {"source": "XBRL", "unit_scale": 1, "fs_div": fs_div}

        # 빈 연도는 HTML 파싱
        empty_years = [y for y in years
                       if not yearly_data[y] or all(v is None for v in yearly_data[y].values())]
        for i, y in enumerate(empty_years):
            if progress_callback:
                progress_callback(i, len(empty_years), f"외감 HTML 파싱 {y}년")
            result = extract_external_audit_data(_dart, corp_code, y, prefer_consolidated)
            if result.get("data") and any(v is not None for v in result["data"].values()):
                yearly_data[y] = result["data"]
                yearly_meta[y] = {
                    "source": "HTML(외감)",
                    "unit_scale": result.get("unit_scale", 1),
                    "rcept_no": result.get("rcept_no"),
                    "report_nm": result.get("report_nm"),
                    "debug": result.get("debug"),
                    "statement_urls": result.get("statement_urls"),
                    "other_fa_breakdown": result.get("other_fa_breakdown", {}),
                }
            else:
                yearly_meta[y] = {"source": "FAILED", "error": result.get("error")}

    # ============================================================
    # 첫 연도 매출 성장률 계산용 직전년 매출 추출
    # 조회 범위 첫 연도(예: 2021)의 재무제표에 이미 전기(2020) 열이 존재하므로,
    # 이를 재활용해 prev_year_revenue로 별도 저장.
    # XBRL: extract_from_xbrl(year_offset=1), HTML: find_account_in_html(year_offset=1)
    # ============================================================
    first_year = min(years)
    prev_rev_won = None  # 원 단위로 정규화해 저장
    try:
        meta = yearly_meta.get(first_year, {})
        src = meta.get("source", "")
        if src == "XBRL":
            # base_year = first_year+1 의 XBRL이 first_year 데이터 제공했으므로
            # first_year 자체의 base_year=first_year 호출에서 year_offset=1이 first_year-1
            df = fetch_xbrl_finstate(_dart, corp_code, first_year, fs_div)
            if df is not None and not df.empty:
                prev = extract_from_xbrl(df, year_offset=1)
                if prev.get("매출액") is not None:
                    prev_rev_won = prev["매출액"]  # XBRL은 unit_scale=1
        elif src in ("HTML(외감)", "HTML(외감 폴백)"):
            # 외감사: first_year의 보고서에서 전기 열 읽기
            reports = find_external_audit_reports(_dart, corp_code, first_year)
            if reports:
                chosen = next((r for r in reports if r["is_consolidated"]), reports[0]) \
                    if prefer_consolidated else \
                    next((r for r in reports if not r["is_consolidated"]), reports[0])
                sub_docs = get_sub_docs(_dart, chosen["rcept_no"])
                if sub_docs is not None:
                    # IS 우선, 없으면 통합 페이지
                    is_url = find_statement_url(sub_docs, "손익계산서") \
                        or find_statement_url(sub_docs, "포괄손익계산서")
                    is_df = parse_html_statement(is_url) if is_url else None
                    if is_df is None:
                        combined_url = find_statement_url(sub_docs, "재무제표")
                        if combined_url:
                            classified = parse_combined_statements(combined_url)
                            is_df = classified.get("IS")
                    if is_df is not None:
                        prev_val = find_account_in_html(is_df, ACCOUNT_KEYWORDS["매출액"], year_offset=1)
                        if prev_val is not None:
                            # 메인 unit_scale 적용 (HTML 외감은 unit_scale이 단위에 따라서 이미 반영됨)
                            us = yearly_meta[first_year].get("unit_scale", 1)
                            prev_rev_won = prev_val * us
        # 메타에 저장 (원 단위)
        if prev_rev_won is not None:
            yearly_meta[first_year]["prev_year_revenue_won"] = prev_rev_won
    except Exception as e:
        yearly_meta[first_year]["prev_year_revenue_error"] = str(e)

    return yearly_data, yearly_meta


# ====================================================================
# 7) 템플릿 표 생성
# ====================================================================
# Valuation 정의: 요약표의 "현금성자산" / "총차입금"은 아래 구성 항목 합으로 정의됨.
# 구성표와 동일한 소스이므로 합계가 일치해야 함.
_CASH_COMPOSITION_ITEMS = [
    ("현금및현금성자산", "현금성자산"),
    ("단기금융상품", "_단기금융상품"),
    ("단기투자자산", "_단기투자자산"),
    ("당기손익-공정가치측정금융자산(유동)", "_당기손익공정가치_유동"),
    ("매도가능금융자산/매도가능증권(유동)", "_매도가능금융자산_유동"),
    ("기타포괄손익-공정가치측정금융자산(유동)", "_기타포괄손익공정가치_유동"),
    ("장기금융상품/장기성예금", "_장기금융상품"),
    # v14: "기타유동금융자산" 통째 논입을 제거. 주석에서 분해한 cash-like 항목만
    # 동적으로 추가됨 (parse_other_fa_breakdown → yearly_meta[year]["other_fa_breakdown"]).
]

_DEBT_COMPOSITION_ITEMS = [
    ("단기차입금", "_단기차입금"),
    ("유동성장기부채(차입금/사채)", "_유동성장기부채"),
    ("장기차입금", "_장기차입금"),
    ("사채(일반)", "_사채"),
    ("유동리스부채", "_유동리스부채"),
    ("비유동리스부채", "_비유동리스부채"),
    ("전환사채(CB)", "_전환사채"),
    ("신주인수권부사채(BW)", "_신주인수권부사채"),
    ("교환사채(EB)", "_교환사채"),
]

_CASH_KEYS = [k for _, k in _CASH_COMPOSITION_ITEMS]
_DEBT_KEYS = [k for _, k in _DEBT_COMPOSITION_ITEMS]


# --------------------------------------------------------------------
# v14: 주석 분해 cash-like 합산 헬퍼
# yearly_meta[y]["other_fa_breakdown"]["cash_like"]의 원단위값을 억원으로 환산.
# 주석 분해 단계에서 이미 원 단위로 환산되었으므로 unit_scale 재적용 안 함.
# --------------------------------------------------------------------
def _other_fa_cash_like_eokwon(yearly_meta: Dict, year: int) -> Optional[float]:
    """주석 분해 cash_like 합을 억원으로 반환. 분해 결과 없으면 None."""
    bd = yearly_meta.get(year, {}).get("other_fa_breakdown")
    if not bd or not isinstance(bd, dict):
        return None
    cl = bd.get("cash_like", {}) or {}
    if not cl:
        return 0.0
    total_won = sum(v for v in cl.values() if isinstance(v, (int, float)))
    return round(total_won / 1e8)


# --------------------------------------------------------------------
# v15: 계정·연도별 원 단위 계산 헬퍼 (표·차트 공용 데이터 소스)
# --------------------------------------------------------------------
def _val_won(yearly_data: Dict, yearly_meta: Dict, year: int, key: str) -> Optional[int]:
    v = yearly_data.get(year, {}).get(key)
    sc = yearly_meta.get(year, {}).get("unit_scale", 1)
    return won_value(v, sc)


def _sum_won(yearly_data: Dict, yearly_meta: Dict, year: int, keys: List[str]) -> Optional[int]:
    sc = yearly_meta.get(year, {}).get("unit_scale", 1)
    d = yearly_data.get(year, {})
    vals = [d.get(k) for k in keys if d.get(k) is not None]
    if not vals:
        return None
    return won_value(sum(vals), sc)


def _other_fa_cash_like_won(yearly_meta: Dict, year: int) -> Optional[int]:
    """주석 분해 cash_like 합을 원 단위로."""
    bd = yearly_meta.get(year, {}).get("other_fa_breakdown")
    if not bd or not isinstance(bd, dict):
        return None
    cl = bd.get("cash_like", {}) or {}
    if not cl:
        return 0
    return int(sum(v for v in cl.values() if isinstance(v, (int, float))))


def compute_yearly_metrics(yearly_data: Dict, yearly_meta: Dict, years: List[int]) -> Dict:
    """연도별 원 단위 지표 종합. 차트·표 공용.
    반환:
      {
        "revenue":   {year: 원값},
        "growth_pct":{year: 성장률%},   # 첫 연도는 prev_year_revenue_won 활용
        "op_income": {year: 원값},
        "op_margin_pct":  {year: %},
        "ebitda":    {year: 원값},
        "ebitda_margin_pct": {year: %},
        "net_income":{year: 원값},
        "net_margin_pct":{year: %},
        "total_assets":      {year: 원값},
        "total_liabilities": {year: 원값},
        "total_equity":      {year: 원값},
        "cash_equiv":   {year: 원값},   # _CASH_KEYS 합 + 주석 분해 cash_like
        "total_borrow": {year: 원값},   # _DEBT_KEYS 합
      }
    """
    out = {k: {} for k in [
        "revenue", "growth_pct", "op_income", "op_margin_pct",
        "ebitda", "ebitda_margin_pct", "net_income", "net_margin_pct",
        "total_assets", "total_liabilities", "total_equity",
        "cash_equiv", "total_borrow",
    ]}
    sorted_years = sorted(years)

    for y in sorted_years:
        rev = _val_won(yearly_data, yearly_meta, y, "매출액")
        op = _val_won(yearly_data, yearly_meta, y, "영업이익")
        ni = _val_won(yearly_data, yearly_meta, y, "당기순이익")
        ta = _val_won(yearly_data, yearly_meta, y, "자산총계")
        tl = _val_won(yearly_data, yearly_meta, y, "부채총계")
        te = _val_won(yearly_data, yearly_meta, y, "자본총계")
        da = _sum_won(yearly_data, yearly_meta, y,
                      ["_유형자산감가상각비", "_무형자산상각비", "_사용권자산상각비"])
        cash_static = _sum_won(yearly_data, yearly_meta, y, _CASH_KEYS)
        cash_bd = _other_fa_cash_like_won(yearly_meta, y)
        if cash_static is None and cash_bd is None:
            cash = None
        else:
            cash = (cash_static or 0) + (cash_bd or 0)
        borrow = _sum_won(yearly_data, yearly_meta, y, _DEBT_KEYS)
        # 부채총계 있으나 부채 구성 전부 None이면 0(실제 차입 0)
        if borrow is None and tl is not None:
            borrow = 0

        ebitda = (op + da) if (op is not None and da is not None) else None

        out["revenue"][y] = rev
        out["op_income"][y] = op
        out["net_income"][y] = ni
        out["total_assets"][y] = ta
        out["total_liabilities"][y] = tl
        out["total_equity"][y] = te
        out["cash_equiv"][y] = cash
        out["total_borrow"][y] = borrow
        out["ebitda"][y] = ebitda

        # 이익률
        out["op_margin_pct"][y] = (op / rev * 100) if (op is not None and rev not in (None, 0)) else None
        out["net_margin_pct"][y] = (ni / rev * 100) if (ni is not None and rev not in (None, 0)) else None
        out["ebitda_margin_pct"][y] = (ebitda / rev * 100) if (ebitda is not None and rev not in (None, 0)) else None

    # 성장률: 첫 연도는 prev_year_revenue_won, 이후는 전년 대비
    for i, y in enumerate(sorted_years):
        curr = out["revenue"][y]
        if i == 0:
            prev = yearly_meta.get(y, {}).get("prev_year_revenue_won")
        else:
            prev = out["revenue"][sorted_years[i - 1]]
        if curr is None or prev is None or prev == 0:
            out["growth_pct"][y] = None
        else:
            out["growth_pct"][y] = (curr - prev) / abs(prev) * 100

    return out


def build_template_table(yearly_data: Dict, yearly_meta: Dict, years: List[int],
                         unit_label: str = "억원") -> pd.DataFrame:
    """v15: unit_label에 따라 단위 일괄 적용."""
    m = compute_yearly_metrics(yearly_data, yearly_meta, years)
    sorted_years = sorted(years)

    def fu(v_won):
        return format_unit(to_unit(v_won, unit_label), unit_label)

    rows = []
    # 매출액 + Growth
    rows.append(["매출액"] + [fu(m["revenue"][y]) for y in sorted_years])
    rows.append(["  Growth"] + [format_pct(m["growth_pct"][y]) for y in sorted_years])

    # EBITDA + Margin
    rows.append(["EBITDA"] + [fu(m["ebitda"][y]) for y in sorted_years])
    rows.append(["  Margin"] + [format_pct(m["ebitda_margin_pct"][y]) for y in sorted_years])

    # 영업이익 + Margin
    rows.append(["영업이익"] + [fu(m["op_income"][y]) for y in sorted_years])
    rows.append(["  Margin"] + [format_pct(m["op_margin_pct"][y]) for y in sorted_years])

    # 당기순이익 + Margin
    rows.append(["당기순이익"] + [fu(m["net_income"][y]) for y in sorted_years])
    rows.append(["  Margin"] + [format_pct(m["net_margin_pct"][y]) for y in sorted_years])

    # 자산총계 · 현금성 · 부채총계 · 총차입 · 자본총계
    rows.append(["자산총계"] + [fu(m["total_assets"][y]) for y in sorted_years])
    rows.append(["  현금성자산"] + [fu(m["cash_equiv"][y]) for y in sorted_years])
    rows.append(["부채총계"] + [fu(m["total_liabilities"][y]) for y in sorted_years])
    rows.append(["  총차입금"] + [fu(m["total_borrow"][y]) for y in sorted_years])
    rows.append(["자본총계"] + [fu(m["total_equity"][y]) for y in sorted_years])

    columns = [f"(단위: {unit_label})"] + [str(y) for y in sorted_years]
    return pd.DataFrame(rows, columns=columns)


# --------------------------------------------------------------------
# Valuation 관점 현금성자산 / 차입금 구성표
# --------------------------------------------------------------------
# 항목 리스트는 위 _CASH_COMPOSITION_ITEMS / _DEBT_COMPOSITION_ITEMS를 공유.
# 요약표의 현금성자산·총차입금과 구성표 합계는 동일한 소스 → 값 일치.


def _build_composition_table(
    items: List[tuple],
    yearly_data: Dict,
    yearly_meta: Dict,
    years: List[int],
    total_label: str,
    drop_empty_rows: bool = True,
    unit_label: str = "억원",
) -> pd.DataFrame:
    """항목×연도 구성표 생성. v15: 단위 선택가능."""
    sorted_years = sorted(years)

    def get_won(year, key):
        v = yearly_data.get(year, {}).get(key)
        scale = yearly_meta.get(year, {}).get("unit_scale", 1)
        return won_value(v, scale)

    def fu(v_won):
        return format_unit(to_unit(v_won, unit_label), unit_label)

    rows = []
    totals_won = {y: 0 for y in sorted_years}
    any_total = {y: False for y in sorted_years}

    for label, key in items:
        vals = {y: get_won(y, key) for y in sorted_years}
        if drop_empty_rows and all(v is None for v in vals.values()):
            continue
        rows.append([label] + [fu(vals[y]) for y in sorted_years])
        for y in sorted_years:
            if vals[y] is not None:
                totals_won[y] += vals[y]
                any_total[y] = True

    # 합계 행
    total_row = [total_label]
    for y in sorted_years:
        total_row.append(fu(totals_won[y]) if any_total[y] else "N/A")
    rows.append(total_row)

    columns = [f"(단위: {unit_label})"] + [str(y) for y in sorted_years]
    return pd.DataFrame(rows, columns=columns)


def build_cash_composition_table(yearly_data: Dict, yearly_meta: Dict, years: List[int],
                                 unit_label: str = "억원") -> pd.DataFrame:
    """현금성자산 구성표 (valuation Net Debt 차감항 후보).

    v14 변경:
    - 정적 BS 계정(_CASH_COMPOSITION_ITEMS) 행 + 주석 분해 cash-like 행("기타금융자산 분해")
      을 함께 노출. 합계는 두 부분의 합.
    - 별도 섹션에 non_cash_like(보증금/대여금 등) 행을 "[비포함 — 참고]"로 표기.
      합계에는 포함되지 않음.
    """
    sorted_years = sorted(years)

    def get_won(year, key):
        v = yearly_data.get(year, {}).get(key)
        scale = yearly_meta.get(year, {}).get("unit_scale", 1)
        return won_value(v, scale)

    def fu(v_won):
        return format_unit(to_unit(v_won, unit_label), unit_label)

    rows = []
    totals_won = {y: 0 for y in sorted_years}
    any_total = {y: False for y in sorted_years}

    # ----- (1) 정적 BS 계정 -----
    for label, key in _CASH_COMPOSITION_ITEMS:
        vals = {y: get_won(y, key) for y in sorted_years}
        if all(v is None for v in vals.values()):
            continue
        rows.append([label] + [fu(vals[y]) for y in sorted_years])
        for y in sorted_years:
            if vals[y] is not None:
                totals_won[y] += vals[y]
                any_total[y] = True

    # ----- (2) 주석 분해 cash-like (기타유동금융자산 풀어쓴 항목) -----
    cash_like_by_year: Dict[int, Dict[str, float]] = {}
    for y in sorted_years:
        bd = yearly_meta.get(y, {}).get("other_fa_breakdown") or {}
        cl = bd.get("cash_like", {}) if isinstance(bd, dict) else {}
        cash_like_by_year[y] = cl if isinstance(cl, dict) else {}

    all_cash_names: List[str] = []
    seen = set()
    for y in sorted_years:
        for name in cash_like_by_year[y].keys():
            if name not in seen:
                seen.add(name)
                all_cash_names.append(name)

    if all_cash_names:
        rows.append(["[기타금융자산 주석 분해 — 포함]"] + ["" for _ in sorted_years])
        for name in all_cash_names:
            row = [f"　{name}"]
            for y in sorted_years:
                won_v = cash_like_by_year[y].get(name)
                row.append(fu(won_v))
                if won_v is not None:
                    totals_won[y] += won_v
                    any_total[y] = True
            rows.append(row)

    # 합계
    total_row = ["　합계 (현금성자산 후보 총합)"]
    for y in sorted_years:
        total_row.append(fu(totals_won[y]) if any_total[y] else "N/A")
    rows.append(total_row)

    # ----- (3) non_cash_like (참고용, 합계 비포함) -----
    non_cash_by_year: Dict[int, Dict[str, float]] = {}
    for y in sorted_years:
        bd = yearly_meta.get(y, {}).get("other_fa_breakdown") or {}
        ncl = bd.get("non_cash_like", {}) if isinstance(bd, dict) else {}
        non_cash_by_year[y] = ncl if isinstance(ncl, dict) else {}

    all_non_cash_names: List[str] = []
    seen2 = set()
    for y in sorted_years:
        for name in non_cash_by_year[y].keys():
            if name not in seen2:
                seen2.add(name)
                all_non_cash_names.append(name)

    if all_non_cash_names:
        rows.append(["[기타금융자산 주석 분해 — 비포함 · 참고]"] + ["" for _ in sorted_years])
        for name in all_non_cash_names:
            row = [f"　{name}"]
            for y in sorted_years:
                won_v = non_cash_by_year[y].get(name)
                row.append(fu(won_v))
            rows.append(row)

    columns = [f"(단위: {unit_label})"] + [str(y) for y in sorted_years]
    return pd.DataFrame(rows, columns=columns)


def build_debt_composition_table(yearly_data: Dict, yearly_meta: Dict, years: List[int],
                                 unit_label: str = "억원") -> pd.DataFrame:
    """차입금 구성표 (valuation Gross Debt 가산항 후보)."""
    return _build_composition_table(
        _DEBT_COMPOSITION_ITEMS, yearly_data, yearly_meta, years,
        total_label="　합계 (총차입금 후보 합)",
        unit_label=unit_label,
    )


# ====================================================================
# v15: 차트 헬퍼 (Plotly)
# ====================================================================
# 4개 콤보 차트 (막대+라인)·1개 다중 라인 차트
# - 원 단위 metrics dict를 입력으로 받고, unit_label에 따라 시각적 단위 표시.
# - 원 단위 값이 있는 연도만 그림 → N/A는 구멍으로 넘어감.
def _import_plotly():
    """Plotly는 선택적 import (Streamlit Cloud 다운 시간 절약)."""
    try:
        import plotly.graph_objects as go
        return go
    except Exception:
        return None


def _years_with_unit_values(metric_dict: Dict, sorted_years: List[int], unit_label: str):
    """연도 순서대로 (year_label, value_in_unit) 투플 몄서 반환. None은 건너뛄."""
    xs, ys = [], []
    for y in sorted_years:
        v = metric_dict.get(y)
        if v is None:
            continue
        xs.append(str(y))
        ys.append(to_unit(v, unit_label))
    return xs, ys


def _format_pct_label(v):
    if v is None:
        return ""
    if v < 0:
        return f"({abs(v):.1f}%)"
    return f"{v:.1f}%"


def _format_unit_label(v, unit_label):
    if v is None:
        return ""
    spec = UNIT_SPECS.get(unit_label, UNIT_SPECS["억원"])
    dec = spec["decimals"]
    if dec == 0:
        return f"{int(round(v)):,}" if v >= 0 else f"({int(round(abs(v))):,})"
    fmt = f"{{:,.{dec}f}}"
    return fmt.format(v) if v >= 0 else "(" + fmt.format(abs(v)) + ")"


def _make_combo_chart(title: str, sorted_years: List[int],
                      bar_values_won: Dict, line_pct: Dict,
                      unit_label: str,
                      bar_name: str, line_name: str,
                      bar_color: str = "#1E3D6B", line_color: str = "#E5862E"):
    """막대(금액)+라인(%) 콤보 차트."""
    go = _import_plotly()
    if go is None:
        return None
    xs_bar = [str(y) for y in sorted_years]
    ys_bar = [to_unit(bar_values_won.get(y), unit_label) for y in sorted_years]
    ys_line = [line_pct.get(y) for y in sorted_years]
    bar_text = [_format_unit_label(v, unit_label) if v is not None else "" for v in ys_bar]
    line_text = [_format_pct_label(v) for v in ys_line]

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=xs_bar, y=ys_bar, name=bar_name,
        marker_color=bar_color,
        text=bar_text, textposition="outside",
        textfont=dict(size=13, color=bar_color),
        cliponaxis=False,
        hovertemplate=f"%{{x}}<br>{bar_name}: %{{text}} {unit_label}<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=xs_bar, y=ys_line, name=line_name,
        mode="lines+markers+text",
        line=dict(color=line_color, width=2.5),
        marker=dict(size=9, color=line_color),
        text=line_text, textposition="top center",
        textfont=dict(size=12, color=line_color),
        yaxis="y2",
        hovertemplate=f"%{{x}}<br>{line_name}: %{{text}}<extra></extra>",
    ))
    fig.update_layout(
        title=dict(text=f"<b>{title}</b>", x=0.5, xanchor="center",
                   font=dict(size=18, color="#1F2937")),
        height=460,
        margin=dict(l=40, r=40, t=80, b=40),
        plot_bgcolor="white",
        paper_bgcolor="white",
        showlegend=True,
        legend=dict(orientation="h", yanchor="top", y=1.08, xanchor="center", x=0.5,
                    bgcolor="rgba(0,0,0,0)"),
        xaxis=dict(showgrid=False, showline=True, linecolor="#1F2937", linewidth=1,
                   tickfont=dict(size=13)),
        yaxis=dict(visible=False, showgrid=False, zeroline=False),
        yaxis2=dict(visible=False, overlaying="y", side="right", showgrid=False, zeroline=False),
        bargap=0.45,
    )
    return fig


def _make_balance_chart(title: str, sorted_years: List[int], metrics: Dict, unit_label: str):
    """재무상태 5라인 차트."""
    go = _import_plotly()
    if go is None:
        return None
    series_def = [
        ("자산총계", "total_assets",      "#1E3D6B"),
        ("현금성자산", "cash_equiv",      "#7DB8E8"),
        ("부채총계", "total_liabilities", "#C7383C"),
        ("총차입금", "total_borrow",      "#F2A6A8"),
        ("자본총계", "total_equity",      "#2E9F5C"),
    ]
    fig = go.Figure()
    for label, key, color in series_def:
        xs, ys = _years_with_unit_values(metrics[key], sorted_years, unit_label)
        if not xs:
            continue
        text_labels = [_format_unit_label(v, unit_label) for v in ys]
        fig.add_trace(go.Scatter(
            x=xs, y=ys, name=label,
            mode="lines+markers+text",
            line=dict(color=color, width=2.5),
            marker=dict(size=9, color=color),
            text=text_labels, textposition="top center",
            textfont=dict(size=11, color=color),
            hovertemplate=f"%{{x}}<br>{label}: %{{text}} {unit_label}<extra></extra>",
        ))
    fig.update_layout(
        title=dict(text=f"<b>{title}</b>", x=0.5, xanchor="center",
                   font=dict(size=18, color="#1F2937")),
        height=520,
        margin=dict(l=40, r=40, t=90, b=40),
        plot_bgcolor="white",
        paper_bgcolor="white",
        legend=dict(orientation="h", yanchor="top", y=1.10, xanchor="center", x=0.5,
                    bgcolor="rgba(0,0,0,0)"),
        xaxis=dict(showgrid=False, showline=True, linecolor="#1F2937", linewidth=1,
                   tickfont=dict(size=13)),
        yaxis=dict(visible=False, showgrid=False, zeroline=False),
    )
    return fig


def build_all_charts(metrics: Dict, years: List[int], unit_label: str = "억원") -> Dict:
    """차트 5개 일괄 생성. 대시보드용 dict 반환.
    Plotly 설치 실패 시 동일 키 구조로 None 반환.
    """
    sorted_years = sorted(years)
    out = {}
    out["revenue"] = _make_combo_chart(
        "매출액 · 성장률", sorted_years,
        metrics["revenue"], metrics["growth_pct"], unit_label,
        bar_name="매출액", line_name="성장률(%)",
    )
    out["ebitda"] = _make_combo_chart(
        "EBITDA · 이익률", sorted_years,
        metrics["ebitda"], metrics["ebitda_margin_pct"], unit_label,
        bar_name="EBITDA", line_name="이익률(%)",
    )
    out["op_income"] = _make_combo_chart(
        "영업이익 · 이익률", sorted_years,
        metrics["op_income"], metrics["op_margin_pct"], unit_label,
        bar_name="영업이익", line_name="이익률(%)",
    )
    out["net_income"] = _make_combo_chart(
        "당기순이익 · 이익률", sorted_years,
        metrics["net_income"], metrics["net_margin_pct"], unit_label,
        bar_name="당기순이익", line_name="이익률(%)",
    )
    out["balance"] = _make_balance_chart(
        "재무상태 (자산·부채·자본 / 현금성자산·총차입금)",
        sorted_years, metrics, unit_label,
    )
    return out


# ====================================================================
# 8) UI
# ====================================================================
# 사이드바 스타일 + 본문 박스(info-pill) + segmented radio 스타일
_UI_CSS = """
<style>
/* 사이드바 헤더 */
section[data-testid="stSidebar"] h2,
section[data-testid="stSidebar"] h3 {
    color: #1E3D6B;
    font-size: 1.0rem;
    margin-top: 0.25rem;
}
section[data-testid="stSidebar"] hr { margin: 8px 0; }

/* 사이드바 라디오 - 가로 토글 형태 */
section[data-testid="stSidebar"] div[role="radiogroup"] {
    gap: 6px;
    flex-wrap: wrap;
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label {
    flex: 1 1 auto;
    min-width: 0;
    padding: 8px 10px;
    border: 1px solid #d6d6d6;
    border-radius: 8px;
    background: #ffffff;
    cursor: pointer;
    text-align: center;
    margin: 0 !important;
    transition: all 0.15s ease;
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label:hover {
    border-color: #E5862E;
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label[data-checked="true"],
section[data-testid="stSidebar"] div[role="radiogroup"] > label:has(input:checked) {
    border: 2px solid #C7383C;
    background: #FEEFEF;
    font-weight: 600;
}
/* 라디오 동그라미 숨김 - 라벨만 표시 */
section[data-testid="stSidebar"] div[role="radiogroup"] > label > div:first-child {
    display: none;
}
section[data-testid="stSidebar"] div[role="radiogroup"] > label > div {
    width: 100%;
}

/* 본문 헤더 둥근 박스 (info-pill) */
.info-pill-row {
    display: flex;
    flex-wrap: wrap;
    gap: 10px;
    margin: 10px 0 14px 0;
}
.info-pill {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 8px 16px;
    border-radius: 999px;
    border: 1px solid #d8d8d8;
    background: #f7f9fc;
    font-size: 0.95rem;
    color: #1E3D6B;
}
.info-pill .pill-key {
    color: #6b7785;
    font-size: 0.82rem;
}
.info-pill .pill-val {
    font-weight: 700;
}
.info-pill.pill-primary {
    background: #1E3D6B;
    color: #ffffff;
    border-color: #1E3D6B;
}
.info-pill.pill-primary .pill-key { color: #c9d4e3; }
.info-pill.pill-accent {
    background: #FEEFEF;
    border-color: #E5862E;
    color: #C7383C;
}
.info-pill.pill-accent .pill-key { color: #9a5a2b; }
</style>
"""
st.markdown(_UI_CSS, unsafe_allow_html=True)

# ----- 좌측 사이드바: 조회 옵션 (토글 3그룹) -----
st.sidebar.markdown("### ⚙️ 조회 옵션")

st.sidebar.markdown("**재무제표 구분**")
fs_label = st.sidebar.radio(
    "재무제표 구분",
    ["연결", "별도"],
    index=0,
    horizontal=True,
    label_visibility="collapsed",
    key="fs_radio",
)
fs_div_target = "CFS" if fs_label == "연결" else "OFS"

st.sidebar.markdown("**조회 기간**")
period_label = st.sidebar.radio(
    "조회 기간",
    ["5년", "10년", "20년", "최대"],
    index=0,
    horizontal=True,
    label_visibility="collapsed",
    key="period_radio",
)
period_map = {"5년": 5, "10년": 10, "20년": 20, "최대": 99}

st.sidebar.markdown("**표시 단위**")
unit_label = st.sidebar.radio(
    "표시 단위",
    ["백만원", "억원", "십억원"],
    index=1,
    horizontal=True,
    label_visibility="collapsed",
    key="unit_radio",
)

st.sidebar.caption("옵션을 바꾸면 결과가 즉시 갱신됩니다.")

current_year = datetime.now().year
end_year = current_year - 1  # 종료연도는 직전 회계연도 자동 고정

# ====================================================================
# 9) 회사 검색 (사명 변경 대응 다층 fallback)
# ====================================================================
@st.cache_data(ttl=86400, show_spinner=False)
def download_corp_code_xml(api_key: str) -> pd.DataFrame:
    """DART 전체 기업 코드 목록을 직접 다운로드.
    OpenDartReader 내부 캐시와 별개로 매일 갱신.
    반환: corp_code / corp_name / corp_eng_name / stock_code / modify_date 컬럼
    """
    try:
        url = f"https://opendart.fss.or.kr/api/corpCode.xml?crtfc_key={api_key}"
        r = requests.get(url, timeout=60)
        if r.status_code != 200:
            return pd.DataFrame()
        z = zipfile.ZipFile(io.BytesIO(r.content))
        xml_content = z.read("CORPCODE.xml").decode("utf-8")
        root = ET.fromstring(xml_content)
        rows = []
        for elem in root.iter("list"):
            rows.append({
                "corp_code": elem.findtext("corp_code", "").strip(),
                "corp_name": elem.findtext("corp_name", "").strip(),
                "corp_eng_name": elem.findtext("corp_eng_name", "").strip(),
                "stock_code": elem.findtext("stock_code", "").strip(),
                "modify_date": elem.findtext("modify_date", "").strip(),
            })
        return pd.DataFrame(rows)
    except Exception as e:
        st.warning(f"corp_code.xml 다운로드 실패: {e}")
        return pd.DataFrame()


def search_in_corp_xml(corp_df: pd.DataFrame, query: str) -> pd.DataFrame:
    """corp_code.xml에서 부분 매칭으로 검색.
    한글/영문 모두 검색. 공백 무시.
    """
    if corp_df is None or corp_df.empty:
        return pd.DataFrame()
    q = query.strip().replace(" ", "")
    if not q:
        return pd.DataFrame()

    name_norm = corp_df["corp_name"].astype(str).str.replace(" ", "")
    eng_norm = corp_df["corp_eng_name"].astype(str).str.replace(" ", "").str.lower()
    q_lower = q.lower()

    mask = name_norm.str.contains(re.escape(q), na=False) | \
           eng_norm.str.contains(re.escape(q_lower), na=False)
    hits = corp_df[mask].copy()
    if hits.empty:
        return pd.DataFrame()

    # 정확 일치 우선 정렬
    def rank(row):
        nm = str(row["corp_name"]).replace(" ", "")
        en = str(row["corp_eng_name"]).replace(" ", "").lower()
        if nm == q:
            return 0
        if q in nm and len(nm) - len(q) <= 3:
            return 1
        if en == q_lower:
            return 2
        if q in nm:
            return 3
        return 4
    hits["_rank"] = hits.apply(rank, axis=1)
    hits = hits.sort_values("_rank")
    return hits.drop(columns=["_rank"]).head(50)


def enrich_company_info(_dart, corp_code: str) -> Optional[Dict]:
    """corp_code로 company() 호출해 corp_cls 등 상세 정보 보강."""
    try:
        info = _dart.company(corp_code)
        if isinstance(info, dict):
            return info
        if isinstance(info, list) and info:
            return info[0]
    except Exception:
        pass
    return None


@st.cache_data(ttl=3600, show_spinner=False)
def search_companies(_dart, _api_key: str, query: str) -> pd.DataFrame:
    """
    다층 검색:
    1) corp_code 직접 입력이면 company() 호출
    2) company_by_name 시도
    3) 결과 없으면 corp_code.xml 직접 검색 (영문명 포함)
    4) 최종 결과의 corp_cls가 비면 company()로 보강
    """
    q = query.strip()
    if not q:
        return pd.DataFrame()

    # 1) corp_code
    if q.isdigit() and len(q) in (6, 8):
        try:
            info = _dart.company(q)
            if info is None:
                return pd.DataFrame()
            if isinstance(info, dict):
                return pd.DataFrame([info])
            if isinstance(info, list):
                return pd.DataFrame(info)
            return pd.DataFrame(info)
        except Exception as e:
            st.warning(f"검색 오류 (corp_code): {e}")
            return pd.DataFrame()

    # 2) company_by_name
    primary_df = pd.DataFrame()
    try:
        result = _dart.company_by_name(q)
        if result is not None:
            primary_df = result if isinstance(result, pd.DataFrame) else pd.DataFrame(result)
    except Exception:
        primary_df = pd.DataFrame()

    # 3) fallback - corp_code.xml 직접 검색 (항상 함께 수행해 누락 방지)
    corp_xml_df = download_corp_code_xml(_api_key)
    xml_hits = search_in_corp_xml(corp_xml_df, q)

    # 두 결과 병합 (primary 우선)
    if primary_df.empty and xml_hits.empty:
        return pd.DataFrame()

    if primary_df.empty:
        # XML hits만 있음 - company() 호출해 corp_cls 보강 (최대 10건)
        enriched = []
        for _, row in xml_hits.head(10).iterrows():
            cc = row["corp_code"]
            info = enrich_company_info(_dart, cc)
            if info:
                enriched.append(info)
            else:
                enriched.append({
                    "corp_code": cc,
                    "corp_name": row["corp_name"],
                    "stock_code": row.get("stock_code", ""),
                    "corp_cls": "",
                })
        return pd.DataFrame(enriched)

    if xml_hits.empty:
        return primary_df

    # 둘 다 있으면 primary를 기본으로 사용 (이미 corp_cls 포함됨)
    primary_codes = set(primary_df["corp_code"].astype(str))
    extras = xml_hits[~xml_hits["corp_code"].astype(str).isin(primary_codes)]
    if extras.empty:
        return primary_df

    # extras에 corp_cls 정보 보강
    enriched_extras = []
    for _, row in extras.head(10).iterrows():
        info = enrich_company_info(_dart, row["corp_code"])
        if info:
            enriched_extras.append(info)
        else:
            enriched_extras.append({
                "corp_code": row["corp_code"],
                "corp_name": row["corp_name"],
                "stock_code": row.get("stock_code", ""),
                "corp_cls": "",
            })
    if enriched_extras:
        combined = pd.concat([primary_df, pd.DataFrame(enriched_extras)], ignore_index=True)
        return combined
    return primary_df


# ====================================================================
# 10) 본문 UI - 회사 검색 + 추출 + 결과 렌더링
# ====================================================================

# ----- 본문 상단: 기업 검색 영역 -----
st.markdown("<div class='hpe-section'>🔎 기업 검색</div>", unsafe_allow_html=True)

search_col1, search_col2, search_col3 = st.columns([5, 1.2, 1.2])
with search_col1:
    company_input = st.text_input(
        "회사명 또는 고유번호",
        placeholder="예: 삼성전자 / 005930 / 이브릿지 / 01178885",
        label_visibility="collapsed",
        key="company_input",
    )
with search_col2:
    search_btn = st.button("🔍 검색", use_container_width=True, type="primary")
with search_col3:
    refresh_cache_btn = st.button(
        "🔄 캐시",
        help="사명이 업데이트 안될 때 눌러 corp_code.xml 재다운로드",
        use_container_width=True,
    )

# 캐시 새로고침 (함수 정의 이후에서 처리)
if refresh_cache_btn:
    download_corp_code_xml.clear()
    search_companies.clear()
    st.success("캐시를 비웠습니다. 다시 검색해주세요.")

if search_btn:
    if not company_input.strip():
        st.warning("회사명 또는 코드를 입력하세요.")
        st.stop()
    with st.spinner(f"'{company_input}' 검색 중..."):
        companies = search_companies(dart, api_key, company_input)
    if companies.empty:
        st.error(
            f"'{company_input}'에 해당하는 회사를 찾을 수 없습니다.\n\n"
            "💡 사명이 최근 변경된 경우 우측 '🔄 캐시' 버튼을 눌러보세요."
        )
        st.stop()
    st.session_state["companies"] = companies

if "companies" in st.session_state and not st.session_state["companies"].empty:
    companies = st.session_state["companies"]
    st.markdown(
        f"<div class='hpe-section'>검색 결과 ({len(companies)}건)</div>",
        unsafe_allow_html=True,
    )

    display_cols = [c for c in ["corp_name", "corp_code", "stock_code", "ceo_nm",
                                "corp_cls", "est_dt", "adres", "induty_code"]
                    if c in companies.columns]
    st.dataframe(
        companies[display_cols] if display_cols else companies,
        use_container_width=True, hide_index=True,
    )

    options = []
    for _, row in companies.iterrows():
        label = f"{row.get('corp_name', '?')} (고유: {row.get('corp_code', '?')}"
        if row.get("stock_code") and str(row.get("stock_code")).strip():
            label += f", 종목: {row.get('stock_code')}"
        label += f", 구분: {row.get('corp_cls', '?')})"
        options.append(label)

    selected_idx = st.selectbox(
        "조회할 회사 선택",
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
        f"법인구분 `{corp_cls}` (Y=유가, K=코스닥, N=코넥스, E=외감)"
    )

    if corp_cls == "E":
        st.warning(
            "ℹ️ 외감 비상장기업입니다. 감사보고서의 HTML 본문을 직접 파싱합니다. "
            "단위는 감사보고서 헤더 기준 자동 감지, 표시 단위는 좌측 사이드바 설정을 따릅니다."
        )

    extract_btn = st.button("2️⃣ 데이터 추출", type="primary", use_container_width=True)

    if extract_btn:
        period_n = period_map[period_label]
        start_year = 2015 if period_n == 99 else max(2015, end_year - period_n + 1)
        years = list(range(start_year, end_year + 1))

        progress = st.progress(0.0)
        status = st.empty()

        def update_progress(idx, total, label):
            if total > 0:
                progress.progress(min((idx + 1) / total, 1.0))
            if label:
                status.text(f"📡 {label} ({idx+1}/{total})")

        yearly_data, yearly_meta = collect_multi_year(
            dart, corp_code, corp_cls, years, fs_div_target,
            progress_callback=update_progress,
        )
        progress.empty()
        status.empty()

        # -------- 연결/별도 실제 사용 연도별 추적 --------
        per_year_fs = {}  # {year: "CFS"|"OFS"|"-"}
        fallback_years = []  # 연결 요청이었으나 별도를 쓴 연도
        for y in years:
            meta = yearly_meta.get(y, {}) or {}
            rpt = (meta.get("report_nm") or "")
            src = (meta.get("source") or "")
            is_cfs = ("연결" in rpt) or ("CFS" in src.upper())
            is_ofs = ("별도" in rpt) or ("OFS" in src.upper())
            if is_cfs and not is_ofs:
                per_year_fs[y] = "CFS"
            elif is_ofs and not is_cfs:
                per_year_fs[y] = "OFS"
                if fs_div_target == "CFS":
                    fallback_years.append(y)
            else:
                per_year_fs[y] = fs_div_target if (meta.get("source") and meta.get("source") != "NONE") else "-"

        valid_fs = [v for v in per_year_fs.values() if v in ("CFS", "OFS")]
        if valid_fs:
            main_fs = max(set(valid_fs), key=valid_fs.count)
        else:
            main_fs = fs_div_target
        main_fs_label = "연결" if main_fs == "CFS" else "별도"

        # -------- 헤더 둥근 박스 4개 --------
        period_disp = f"{start_year}~{end_year}"
        pill_html = (
            "<div class='info-pill-row'>"
            f"<div class='info-pill pill-primary'><span class='pill-key'>회사</span> <span class='pill-val'>{corp_name}</span></div>"
            f"<div class='info-pill'><span class='pill-key'>구분</span> <span class='pill-val'>{main_fs_label}재무제표</span></div>"
            f"<div class='info-pill pill-accent'><span class='pill-key'>단위</span> <span class='pill-val'>{unit_label}</span></div>"
            f"<div class='info-pill'><span class='pill-key'>기간</span> <span class='pill-val'>{period_disp}</span></div>"
            "</div>"
        )
        st.markdown(pill_html, unsafe_allow_html=True)

        # 연결 fallback 안내
        if fs_div_target == "CFS" and fallback_years:
            st.warning(
                f"⚠️ 연결재무제표를 요청했으나 다음 연도는 별도재무제표로 대체되었습니다: "
                f"{', '.join(map(str, fallback_years))} · 해당 연도에는 연결보고서가 공시되지 않아 별도를 표시."
            )

        # -------- 데이터 소스 수집 --------
        source_info = []
        for y in years:
            meta = yearly_meta.get(y, {})
            src = meta.get("source", "NONE")
            source_info.append({
                "연도": y,
                "구분": per_year_fs.get(y, "-"),
                "데이터 소스": src,
                "단위 스케일": meta.get("unit_scale", "-"),
                "보고서": meta.get("report_nm", "-"),
                "접수번호": meta.get("rcept_no", "-"),
                "비고": meta.get("error", "-"),
            })

        # -------- 요약 재무제표 --------
        st.markdown("<div class='hpe-section'>요약 재무제표</div>", unsafe_allow_html=True)
        template_df = build_template_table(yearly_data, yearly_meta, years, unit_label=unit_label)
        st.dataframe(template_df, use_container_width=True, hide_index=True)
        st.caption(
            f"· 표시 단위: {unit_label} · 증감률은 퍼센트.\n"
            "· 현금성자산·총차입금은 하단 구성표 합계와 동일 (valuation 정의).\n"
            "· EBITDA = 영업이익 + (유형자산감가상각비 + 무형자산상각비 + 사용권자산상각비). 주석 표기 상각비 누락 가능."
        )

        # -------- 차트 5개 --------
        st.markdown("<div class='hpe-section'>📊 트렌드 차트</div>", unsafe_allow_html=True)
        try:
            metrics = compute_yearly_metrics(yearly_data, yearly_meta, years)
            charts = build_all_charts(metrics, years, unit_label=unit_label)
            chart_order = ["revenue", "ebitda", "op_income", "net_income", "balance"]
            for key in chart_order:
                fig = charts.get(key)
                if fig is not None:
                    st.plotly_chart(fig, use_container_width=True)
        except ImportError:
            st.warning(
                "⚠️ plotly 패키지가 설치되지 않아 차트를 표시할 수 없습니다. "
                "requirements.txt에 plotly 추가 후 재배포하세요."
            )
        except Exception as e:
            st.warning(f"차트 렌더링 오류: {e}")

        # -------- Valuation 구성표 --------
        cash_comp_df = build_cash_composition_table(yearly_data, yearly_meta, years, unit_label=unit_label)
        debt_comp_df = build_debt_composition_table(yearly_data, yearly_meta, years, unit_label=unit_label)

        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("<div class='hpe-section'>현금성자산 구성</div>", unsafe_allow_html=True)
            st.dataframe(cash_comp_df, use_container_width=True, hide_index=True)
            st.caption(
                "· 요약표 '현금성자산' = 이 구성표 합계.\n"
                "· 기타유동·비유동금융자산은 외감사 통합 계정: 정기예금·MMF·단기금융상품 외에 대여금·보증금·파생상품 포함 가능 → 주석 확인 후 가감 필요.\n"
                "· 사용제한·담보 항목 존재 가능 → valuation 시 차감항 조정 필요."
            )
        with col_b:
            st.markdown("<div class='hpe-section'>차입금 구성</div>", unsafe_allow_html=True)
            st.dataframe(debt_comp_df, use_container_width=True, hide_index=True)
            st.caption(
                "· 요약표 '총차입금' = 이 구성표 합계 (리스부채, CB/BW/EB 포함).\n"
                "· IFRS 16 미적용 기업은 리스부채 0 → valuation 시 별도 조정 필요."
            )

        # 결측 안내
        empty_years = [y for y in years
                       if not yearly_data[y] or all(v is None for v in yearly_data[y].values())]
        if empty_years:
            st.warning(
                f"데이터 미수집 연도: {', '.join(map(str, empty_years))} · 하단 '데이터 소스 추적' 섹션 '비고' 컬럼 확인."
            )

        # -------- 엑셀 다운로드 --------
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            template_df.to_excel(writer, sheet_name=f"요약_{unit_label}", index=False)
            cash_comp_df.to_excel(writer, sheet_name="현금성자산_구성", index=False)
            debt_comp_df.to_excel(writer, sheet_name="차입금_구성", index=False)
            raw_rows = []
            for y in years:
                d = yearly_data.get(y, {})
                meta = yearly_meta.get(y, {})
                row = {"연도": y, "소스": meta.get("source"), "단위스케일": meta.get("unit_scale", 1)}
                for key in ACCOUNT_KEYWORDS.keys():
                    row[key] = d.get(key)
                raw_rows.append(row)
            pd.DataFrame(raw_rows).to_excel(writer, sheet_name="원본_원단위", index=False)
            pd.DataFrame(source_info).to_excel(writer, sheet_name="데이터소스", index=False)
            pd.DataFrame([selected_corp]).to_excel(writer, sheet_name="회사정보", index=False)

        st.download_button(
            label="📥 엑셀 다운로드",
            data=output.getvalue(),
            file_name=f"{corp_name}_{start_year}_{end_year}_{main_fs}_{unit_label}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        # -------- 하단: 검증 · 디버그 --------
        st.markdown(
            "<div class='hpe-debug-header'>⚛️ 검증 · 파싱 디버그 · 소스 추적</div>",
            unsafe_allow_html=True,
        )

        with st.expander("데이터 소스 추적 (검증용)", expanded=False):
            st.dataframe(pd.DataFrame(source_info), use_container_width=True, hide_index=True)

        with st.expander("외감 HTML 파싱 디버그"):
            for y in years:
                meta = yearly_meta.get(y, {})
                if meta.get("source", "").startswith("HTML"):
                    st.markdown(f"**{y}년 — {meta.get('report_nm', '?')}** (접수 `{meta.get('rcept_no')}`)")
                    debug = meta.get("debug", [])
                    if debug:
                        st.dataframe(pd.DataFrame(debug), use_container_width=True, hide_index=True)
                    urls = meta.get("statement_urls", {})
                    if urls:
                        for stmt, u in urls.items():
                            if u:
                                st.caption(f"{stmt}: {u}")

st.sidebar.divider()
with st.sidebar.expander("ℹ️ 데이터 처리 방식"):
    st.markdown(
        "**상장사(Y/K/N)**: XBRL API 우선 (3년 묶음 호출)\n\n"
        "**외감(E)**: 감사보고서 `sub_docs` HTML viewer 직접 파싱\n"
        "- 재무상태표 / 손익계산서 / 현금흐름표 본문 표 1\n"
        "- 단위는 헤더에서 자동 감지\n\n"
        "**연결/별도 선택**: 보고서명에 '연결' 포함 여부로 우선순위"
    )

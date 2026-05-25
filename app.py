"""
DART 재무정보 추출 에이전트 v5
- PE 5개년 재무요약 템플릿
- 상장사: finstate_all (XBRL) 우선
- 외감 비상장사: 감사보고서 sub_docs HTML viewer 직접 파싱 (확정 동작)
- 추후 PDF 검증 단계 추가 예정
"""

import io
import re
from datetime import datetime
from typing import Optional, Dict, List, Tuple

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
st.caption("v5 — 외감 비상장사 감사보고서 HTML 직접 파싱 (sub_docs viewer)")

# ====================================================================
# 1) 상수
# ====================================================================
REPRT_CODE_ANNUAL = "11011"  # 사업보고서

# 손익/재무상태/CF 계정 키워드
ACCOUNT_KEYWORDS = {
    "매출액": ["매출액", "수익(매출액)", "영업수익", "매출", "수익"],
    "영업이익": ["영업이익", "영업이익(손실)", "영업손실"],
    "당기순이익": ["당기순이익", "당기순이익(손실)", "당기순손익", "분기순이익", "반기순이익"],
    "자산총계": ["자산총계", "자산 총계"],
    "현금성자산": ["현금및현금성자산", "현금 및 현금성자산"],
    "부채총계": ["부채총계", "부채 총계"],
    "자본총계": ["자본총계", "자본 총계"],
    # 차입금 합산
    "_단기차입금": ["단기차입금"],
    "_유동성장기부채": ["유동성장기부채", "유동성장기차입금", "유동성사채"],
    "_장기차입금": ["장기차입금"],
    "_사채": ["사채"],
    # EBITDA 가산
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
    "_단기차입금": ["BS"],
    "_유동성장기부채": ["BS"],
    "_장기차입금": ["BS"],
    "_사채": ["BS"],
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
    "_단기차입금": "BS",
    "_유동성장기부채": "BS",
    "_장기차입금": "BS",
    "_사채": "BS",
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
def to_number(x) -> Optional[int]:
    if x is None or pd.isna(x):
        return None
    s = str(x).strip().replace(",", "").replace(" ", "").replace("\xa0", "")
    if s in ("", "-", "nan", "None"):
        return None
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


def find_account_in_html(df: pd.DataFrame, keywords: list, year_offset: int = 0) -> Optional[int]:
    """첫 컬럼(과목)에서 매칭 행 검색 → 해당 연도 값 반환."""
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

    # 2차: 부분 매칭 (가장 짧은 계정명 우선)
    candidates = []
    for idx in range(len(df)):
        cell = df.iloc[idx][first_col]
        if pd.isna(cell):
            continue
        norm = normalize_account_name(str(cell))
        for kw in keywords:
            nkw = normalize_account_name(kw)
            if nkw and nkw in norm:
                candidates.append((len(norm), idx, str(cell)))
                break
    candidates.sort()
    for _, idx, _ in candidates:
        v = extract_value_from_row(df, idx, target_cols)
        if v is not None:
            return v
    return None


def extract_external_audit_data(_dart, corp_code: str, year: int,
                                prefer_consolidated: bool = True) -> Dict:
    """
    외감 비상장사: 감사보고서 sub_docs HTML viewer에서 BS/IS/CF 직접 파싱.
    반환: {'data': {...}, 'unit_scale': int, 'rcept_no': str, ...}
    """
    reports = find_external_audit_reports(_dart, corp_code, year)
    if not reports:
        return {"data": None, "error": f"{year}년 외부감사 감사보고서 미발견"}

    # 연결/별도 우선순위에 따라 선택
    if prefer_consolidated:
        chosen = next((r for r in reports if r["is_consolidated"]), reports[0])
    else:
        chosen = next((r for r in reports if not r["is_consolidated"]), reports[0])

    rcept_no = chosen["rcept_no"]
    sub_docs = get_sub_docs(_dart, rcept_no)
    if sub_docs is None:
        return {"data": None, "error": f"sub_docs 조회 실패 (rcept_no={rcept_no})",
                "rcept_no": rcept_no, "report_nm": chosen["report_nm"]}

    # 각 statement URL
    urls = {
        "BS": find_statement_url(sub_docs, "재무상태표"),
        "IS": find_statement_url(sub_docs, "손익계산서"),
        "CF": find_statement_url(sub_docs, "현금흐름표"),
    }
    # 포괄손익계산서 fallback
    if not urls["IS"]:
        urls["IS"] = find_statement_url(sub_docs, "포괄손익계산서")

    # 단위 감지 (BS 헤더 사용)
    unit_scale = 1
    if urls["BS"]:
        unit_scale = detect_unit_from_header(urls["BS"])

    # 각 표 파싱
    dfs = {k: parse_html_statement(u) if u else None for k, u in urls.items()}

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

    return {
        "data": result_data,
        "unit_scale": unit_scale,
        "rcept_no": rcept_no,
        "report_nm": chosen["report_nm"],
        "is_consolidated": chosen["is_consolidated"],
        "debug": debug,
        "statement_urls": urls,
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
                }
            else:
                yearly_meta[y] = {"source": "FAILED", "error": result.get("error")}

    return yearly_data, yearly_meta


# ====================================================================
# 7) 템플릿 표 생성
# ====================================================================
def build_template_table(yearly_data: Dict, yearly_meta: Dict, years: List[int]) -> pd.DataFrame:
    rows = []

    def get_val_eokwon(year, key):
        v = yearly_data.get(year, {}).get(key)
        scale = yearly_meta.get(year, {}).get("unit_scale", 1)
        return to_eokwon(v, scale)

    def safe_sum_eokwon(year, keys):
        scale = yearly_meta.get(year, {}).get("unit_scale", 1)
        d = yearly_data.get(year, {})
        vals = [d.get(k) for k in keys if d.get(k) is not None]
        if not vals:
            return None
        return to_eokwon(sum(vals), scale)

    # 매출액
    revenues = {y: get_val_eokwon(y, "매출액") for y in years}
    rows.append(["매출액"] + [format_eokwon(revenues[y]) for y in years])

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
                growth_row.append(format_pct((curr - prev) / abs(prev) * 100))
    rows.append(growth_row)

    # EBITDA
    op_inc = {y: get_val_eokwon(y, "영업이익") for y in years}
    ebitda = {}
    for y in years:
        op = op_inc[y]
        da = safe_sum_eokwon(y, ["_유형자산감가상각비", "_무형자산상각비", "_사용권자산상각비"])
        if op is None or da is None:
            ebitda[y] = None
        else:
            ebitda[y] = op + da

    rows.append(["EBITDA"] + [format_eokwon(ebitda[y]) for y in years])
    rows.append(["  Margin"] + [
        format_pct(ebitda[y] / revenues[y] * 100)
        if (ebitda[y] is not None and revenues[y] not in (None, 0)) else "N/A"
        for y in years
    ])

    # 영업이익
    rows.append(["영업이익"] + [format_eokwon(op_inc[y]) for y in years])
    rows.append(["  Margin"] + [
        format_pct(op_inc[y] / revenues[y] * 100)
        if (op_inc[y] is not None and revenues[y] not in (None, 0)) else "N/A"
        for y in years
    ])

    # 당기순이익
    net_inc = {y: get_val_eokwon(y, "당기순이익") for y in years}
    rows.append(["당기순이익"] + [format_eokwon(net_inc[y]) for y in years])
    rows.append(["  Margin"] + [
        format_pct(net_inc[y] / revenues[y] * 100)
        if (net_inc[y] is not None and revenues[y] not in (None, 0)) else "N/A"
        for y in years
    ])

    # 자산총계 / 현금성 / 부채 / 차입금 / 자본
    rows.append(["자산총계"] + [format_eokwon(get_val_eokwon(y, "자산총계")) for y in years])
    rows.append(["  현금성자산"] + [format_eokwon(get_val_eokwon(y, "현금성자산")) for y in years])
    rows.append(["부채총계"] + [format_eokwon(get_val_eokwon(y, "부채총계")) for y in years])
    rows.append(["  총차입금"] + [
        format_eokwon(safe_sum_eokwon(y, ["_단기차입금", "_유동성장기부채", "_장기차입금", "_사채"]))
        for y in years
    ])
    rows.append(["자본총계"] + [format_eokwon(get_val_eokwon(y, "자본총계")) for y in years])

    columns = ["(단위: 억원)"] + [str(y) for y in years]
    return pd.DataFrame(rows, columns=columns)


# ====================================================================
# 8) UI
# ====================================================================
st.sidebar.header("🔍 검색 조건")

company_input = st.sidebar.text_input("회사명 또는 코드", placeholder="예: 삼성전자 / 005930 / 은산해운항공")

period_label = st.sidebar.selectbox(
    "조회 기간",
    options=["최근 5년", "최근 10년", "최근 20년", "최대 (2015~)"],
    index=0,
)
period_map = {"최근 5년": 5, "최근 10년": 10, "최근 20년": 20, "최대 (2015~)": 99}

fs_label = st.sidebar.radio("재무제표 구분", ["연결재무제표(CFS)", "별도재무제표(OFS)"], index=0)
fs_div_target = "CFS" if "연결" in fs_label else "OFS"

current_year = datetime.now().year
end_year = st.sidebar.number_input(
    "종료 연도", min_value=2015, max_value=current_year,
    value=current_year - 1, step=1,
)

search_btn = st.sidebar.button("1️⃣ 회사 검색", use_container_width=True)

# ====================================================================
# 9) 회사 검색
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
            "ℹ️ 외감 비상장기업입니다. 감사보고서의 HTML 본문(재무상태표/손익계산서/현금흐름표)을 "
            "직접 파싱합니다. 단위는 헤더 표시를 기준으로 자동 감지하며, 별도/연결 우선순위는 "
            "사이드바 설정을 따릅니다."
        )

    extract_btn = st.button("2️⃣ 데이터 추출", type="primary", use_container_width=True)

    if extract_btn:
        period_n = period_map[period_label]
        start_year = 2015 if period_n == 99 else max(2015, end_year - period_n + 1)
        years = list(range(start_year, end_year + 1))

        st.subheader(f"2️⃣ {corp_name} | {start_year}~{end_year} | {fs_div_target}")

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

        # 데이터 소스 표시
        source_info = []
        for y in years:
            meta = yearly_meta.get(y, {})
            src = meta.get("source", "NONE")
            source_info.append({
                "연도": y,
                "데이터 소스": src,
                "단위 스케일": meta.get("unit_scale", "-"),
                "보고서": meta.get("report_nm", "-"),
                "접수번호": meta.get("rcept_no", "-"),
                "비고": meta.get("error", "-"),
            })
        with st.expander("📋 데이터 소스 추적 (검증용)", expanded=True):
            st.dataframe(pd.DataFrame(source_info), use_container_width=True, hide_index=True)

        # 템플릿 표
        template_df = build_template_table(yearly_data, yearly_meta, years)
        st.dataframe(template_df, use_container_width=True, hide_index=True)

        # 결측 안내
        empty_years = [y for y in years
                       if not yearly_data[y] or all(v is None for v in yearly_data[y].values())]
        if empty_years:
            st.warning(
                f"⚠️ 데이터 미수집 연도: {', '.join(map(str, empty_years))}\n\n"
                "원인은 위 데이터 소스 표의 '비고' 컬럼을 확인하세요."
            )

        st.caption(
            "⚠️ EBITDA = 영업이익 + (유형자산감가상각비 + 무형자산상각비 + 사용권자산상각비). "
            "주석 표기 상각비는 누락 가능. 외감 HTML 파싱은 표 0의 '단위' 표시를 자동 감지."
        )

        # 엑셀 다운로드
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            template_df.to_excel(writer, sheet_name="5개년_요약", index=False)
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
            file_name=f"{corp_name}_{start_year}_{end_year}_{fs_div_target}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        # 디버그
        with st.expander("🔬 외감 HTML 파싱 디버그"):
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

        st.divider()
        st.info(
            "🔜 다음 단계: PDF 본문 검증\n"
            "외감 HTML 파싱 결과는 표 1(본문)만 사용하며, 주석 표기 항목(상각비 등)은 누락 가능."
        )

st.sidebar.divider()
with st.sidebar.expander("ℹ️ 데이터 처리 방식"):
    st.markdown(
        "**상장사(Y/K/N)**: XBRL API 우선 (3년 묶음 호출)\n\n"
        "**외감(E)**: 감사보고서 `sub_docs` HTML viewer 직접 파싱\n"
        "- 재무상태표 / 손익계산서 / 현금흐름표 본문 표 1\n"
        "- 단위는 헤더에서 자동 감지\n"
        "- 컬럼: `제 N(당) 기` / `제 N(당) 기.1` 모두 후보\n\n"
        "**연결/별도 선택**: 보고서명에 '연결' 포함 여부로 우선순위"
    )

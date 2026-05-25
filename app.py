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
    page_title="DART 재무정보 추출 에이전트",
    page_icon="📊",
    layout="wide",
)

st.title("📊 DART 재무정보 추출 에이전트")
st.caption("v8 — 외감사 HTML + 주석 D&A fallback + 직전년 매출/차입금 표시 개선")

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
            # 첫 연도: yearly_meta에 저장된 직전년 매출 활용 (원 단위)
            curr = revenues[y]
            prev_won = yearly_meta.get(y, {}).get("prev_year_revenue_won")
            if curr is None or prev_won is None or prev_won == 0:
                growth_row.append("N/A")
            else:
                # curr는 억원, prev_won은 원 → 억원으로 환산
                prev_eok = prev_won / 1e8
                growth_row.append(format_pct((curr - prev_eok) / abs(prev_eok) * 100))
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
    # 총차입금: 4개 항목 모두 None이고 부채총계가 정상 추출된 경우 → "0" 표시
    # (실제 대차대조표에 차입금 항목 자체가 없는 케이스)
    # BS 추출 자체가 실패한 경우에만 N/A
    borrow_row = ["  총차입금"]
    borrow_keys = ["_단기차입금", "_유동성장기부채", "_장기차입금", "_사채"]
    for y in years:
        d = yearly_data.get(y, {})
        bv = safe_sum_eokwon(y, borrow_keys)
        if bv is None:
            # 4개 모두 None. 부채총계 추출 여부로 구분.
            if d.get("부채총계") is not None:
                borrow_row.append(format_eokwon(0.0))  # 실제 차입금 0
            else:
                borrow_row.append("N/A")
        else:
            borrow_row.append(format_eokwon(bv))
    rows.append(borrow_row)
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

refresh_cache_btn = st.sidebar.button(
    "🔄 회사 목록 캐시 새로고침",
    help="사명이 변경된 회사가 검색되지 않을 때 사용. corp_code.xml을 다시 다운로드합니다.",
    use_container_width=True,
)

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


# 캐시 새로고침 (함수 정의 이후에서 처리)
if refresh_cache_btn:
    download_corp_code_xml.clear()
    search_companies.clear()
    st.sidebar.success("캐시를 비웠습니다. 다시 검색해주세요.")

if search_btn:
    if not company_input.strip():
        st.warning("회사명 또는 코드를 입력하세요.")
        st.stop()
    with st.spinner(f"'{company_input}' 검색 중..."):
        companies = search_companies(dart, api_key, company_input)
    if companies.empty:
        st.error(
            f"'{company_input}'에 해당하는 회사를 찾을 수 없습니다.\n\n"
            "💡 사명이 최근 변경된 경우 좌측 사이드바 하단의 '회사 목록 캐시 새로고침'을 시도해보세요."
        )
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

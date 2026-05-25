"""
DART 재무정보 추출 에이전트 v4
- PE 5개년 재무요약 템플릿
- 상장사: finstate_all (XBRL) 우선
- 외감 비상장사: 감사보고서 첨부 엑셀 자동 다운로드 + 파싱 fallback
- 추후 PDF 검증 단계 추가 예정
"""

import io
import re
import zipfile
import tempfile
import os
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
st.caption("v4 — 외감 비상장사 감사보고서 첨부 엑셀 자동 파싱 지원")

# ====================================================================
# 1) 상수
# ====================================================================
REPRT_CODE_ANNUAL = "11011"  # 사업보고서

# 손익/재무상태 계정 키워드 (XBRL API + 엑셀 파싱 공용)
ACCOUNT_KEYWORDS = {
    "매출액": ["매출액", "수익(매출액)", "영업수익", "매출", "수익"],
    "영업이익": ["영업이익", "영업이익(손실)", "영업손실"],
    "당기순이익": ["당기순이익", "당기순이익(손실)", "당기순손익", "분기순이익", "반기순이익"],
    "자산총계": ["자산총계", "자산 총계"],
    "현금성자산": ["현금및현금성자산", "현금 및 현금성자산"],
    "부채총계": ["부채총계", "부채 총계"],
    "자본총계": ["자본총계", "자본 총계"],
    # 차입금 합산 항목
    "_단기차입금": ["단기차입금"],
    "_유동성장기부채": ["유동성장기부채", "유동성장기차입금", "유동성사채"],
    "_장기차입금": ["장기차입금"],
    "_사채": ["사채"],
    # EBITDA 가산 (CF 또는 주석)
    "_유형자산감가상각비": ["감가상각비", "유형자산감가상각비"],
    "_무형자산상각비": ["무형자산상각비", "무형자산상각"],
    "_사용권자산상각비": ["사용권자산상각비"],
}

# 재무제표 구분 매핑 (XBRL API용)
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
    if s in ("", "-", "nan", "None", "0.0"):
        if s == "0.0":
            return 0
        return None
    try:
        if s.startswith("(") and s.endswith(")"):
            s = "-" + s[1:-1]
        # 소수점 있을 경우 처리
        return int(float(s))
    except (ValueError, TypeError):
        return None


def to_eokwon(v: Optional[int], unit_scale: int = 1) -> Optional[float]:
    """unit_scale: 1=원, 1000=천원, 1000000=백만원"""
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
    """계정명 정규화: 공백/특수문자 제거, 소문자화."""
    if name is None:
        return ""
    s = str(name).strip()
    # 로마숫자/번호 제거 (예: "I. 매출액", "1. 매출액", "Ⅰ. 매출액")
    s = re.sub(r"^[\dIVXⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]+\.\s*", "", s)
    s = re.sub(r"^[\(\[]\s*\d+\s*[\)\]]\s*", "", s)
    s = s.replace(" ", "").replace("　", "")
    return s


def match_account(name: str, keywords: list) -> bool:
    """정규화된 계정명이 키워드 중 하나와 매칭되는지."""
    norm_name = normalize_account_name(name)
    for kw in keywords:
        norm_kw = normalize_account_name(kw)
        if norm_name == norm_kw:
            return True
    # 부분 매칭
    for kw in keywords:
        norm_kw = normalize_account_name(kw)
        if norm_kw and norm_kw in norm_name:
            return True
    return False


# ====================================================================
# 4) XBRL 기반 추출 (상장사 + XBRL 제출한 외감사)
# ====================================================================
def find_in_xbrl(df: pd.DataFrame, keywords: list, sj_filter: list) -> Optional[pd.Series]:
    if df is None or df.empty:
        return None
    work = df.copy()
    if "sj_div" in work.columns:
        work = work[work["sj_div"].isin(sj_filter)]
    if work.empty:
        return None

    # 정확 일치
    for kw in keywords:
        m = work[work["account_nm"] == kw]
        if not m.empty:
            return m.iloc[0]

    # 부분 일치 (가장 짧은 계정명 우선)
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
# 5) 외감 감사보고서 엑셀 fallback 경로
# ====================================================================
@st.cache_data(ttl=3600, show_spinner=False)
def find_audit_report(_dart, corp_code: str, year: int) -> Optional[Dict]:
    """
    해당 연도의 감사보고서/연결감사보고서를 찾음.
    외감 비상장사는 사업보고서가 없으므로 외부감사관련 공시(kind=A,F)에서 찾음.
    """
    try:
        # 다음 해 1월~6월 사이에 제출되는 게 일반적
        start = f"{year}-01-01"
        end = f"{year+1}-12-31"

        # 1차: 정기공시 (A) - 사업보고서/감사보고서 포함
        reports = _dart.list(corp_code, start=start, end=end, kind="A")

        # 2차: 외부감사관련 (F) 별도 시도
        try:
            audit_reports = _dart.list(corp_code, start=start, end=end, kind="F")
            if audit_reports is not None and not audit_reports.empty:
                if reports is None or reports.empty:
                    reports = audit_reports
                else:
                    reports = pd.concat([reports, audit_reports], ignore_index=True)
        except Exception:
            pass

        if reports is None or reports.empty:
            return None

        # 감사보고서/연결감사보고서/사업보고서 우선순위로 정렬
        def priority(report_nm):
            nm = str(report_nm)
            if "사업보고서" in nm:
                return 0
            if "연결감사보고서" in nm:
                return 1
            if "감사보고서" in nm:
                return 2
            return 99

        reports = reports.copy()
        reports["_priority"] = reports["report_nm"].apply(priority)
        reports = reports[reports["_priority"] < 99]
        if reports.empty:
            return None
        reports = reports.sort_values(["_priority", "rcept_dt"], ascending=[True, False])
        top = reports.iloc[0]
        return {
            "rcept_no": top["rcept_no"],
            "report_nm": top["report_nm"],
            "rcept_dt": top.get("rcept_dt"),
        }
    except Exception:
        return None


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_attach_excel_url(_dart, rcept_no: str, prefer_consolidated: bool = True) -> Optional[Tuple[str, str]]:
    """
    감사보고서의 첨부파일 중 재무제표 엑셀 URL을 반환.
    prefer_consolidated: True면 연결 우선
    """
    try:
        files = _dart.attach_files(rcept_no)
        if not files:
            return None
        # files는 dict: {제목: URL}
        if isinstance(files, dict):
            items = list(files.items())
        else:
            return None

        # 엑셀 파일만 필터링
        excel_items = [(t, u) for t, u in items
                       if any(t.lower().endswith(ext) for ext in [".xls", ".xlsx", ".xlsm"])
                       or "excel" in u.lower() or "xls" in u.lower()]

        if not excel_items:
            # URL 패턴으로 엑셀 추정 (DART는 excel.do URL을 씀)
            excel_items = [(t, u) for t, u in items if "excel" in u.lower()]

        if not excel_items:
            return None

        # 연결/별도 선택
        def score(title):
            t = title
            s = 0
            if prefer_consolidated:
                if "연결" in t: s -= 10
                if "별도" in t: s += 5
            else:
                if "별도" in t: s -= 10
                if "연결" in t: s += 5
            if "재무제표" in t: s -= 5
            if "감사보고서" in t: s -= 3
            return s

        excel_items.sort(key=lambda x: score(x[0]))
        return excel_items[0]
    except Exception:
        return None


def download_file(url: str, api_key: str) -> Optional[bytes]:
    """DART 첨부파일 다운로드. URL에 따라 처리 방식 다름."""
    try:
        # API 키 필요한 URL은 자동 처리, 일반 URL은 requests
        resp = requests.get(url, timeout=30, allow_redirects=True)
        if resp.status_code == 200:
            return resp.content
        return None
    except Exception:
        return None


def detect_unit_scale(excel_text: str) -> int:
    """엑셀 내용에서 단위를 감지. 1=원, 1000=천원, 1000000=백만원"""
    text = excel_text[:5000]  # 상단만 검사
    if re.search(r"단위\s*[:：]\s*백만\s*원", text) or "백만원" in text[:2000]:
        return 1_000_000
    if re.search(r"단위\s*[:：]\s*천\s*원", text) or "천원" in text[:2000]:
        return 1_000
    return 1


def parse_excel_for_accounts(excel_bytes: bytes) -> Tuple[Dict[str, Optional[int]], int, List[Dict]]:
    """
    감사보고서 첨부 엑셀에서 계정값을 추출.
    반환: (계정 dict, 단위 scale, 디버그용 매칭 정보 리스트)
    """
    results = {key: None for key in ACCOUNT_KEYWORDS.keys()}
    debug_info = []
    unit_scale = 1

    try:
        # 다중 시트 모두 검사
        xls = pd.ExcelFile(io.BytesIO(excel_bytes))

        all_text_sample = ""
        for sheet_name in xls.sheet_names[:10]:  # 최대 10개 시트
            try:
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None, dtype=str)
            except Exception:
                continue
            if df.empty:
                continue

            # 단위 감지용 텍스트 누적
            text_sample = " ".join(df.head(15).fillna("").astype(str).values.flatten().tolist())
            all_text_sample += text_sample + " "

            # 각 계정 찾기
            for key, keywords in ACCOUNT_KEYWORDS.items():
                if results[key] is not None:
                    continue  # 이미 찾음

                for row_idx in range(len(df)):
                    row_vals = df.iloc[row_idx].fillna("").astype(str).tolist()
                    for col_idx, cell in enumerate(row_vals):
                        cell_str = str(cell).strip()
                        if not cell_str:
                            continue
                        if match_account(cell_str, keywords):
                            # 같은 행에서 오른쪽으로 가며 첫 숫자 찾기
                            for c2 in range(col_idx + 1, min(col_idx + 10, len(row_vals))):
                                v = to_number(row_vals[c2])
                                if v is not None and v != 0:
                                    results[key] = v
                                    debug_info.append({
                                        "항목": key,
                                        "시트": sheet_name,
                                        "원문계정명": cell_str,
                                        "행/열": f"{row_idx+1}/{c2+1}",
                                        "값": v,
                                    })
                                    break
                            if results[key] is not None:
                                break
                    if results[key] is not None:
                        break

        # 단위 감지
        unit_scale = detect_unit_scale(all_text_sample)
    except Exception as e:
        debug_info.append({"오류": str(e)})

    return results, unit_scale, debug_info


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_audit_excel_data(_dart, _api_key: str, corp_code: str, year: int,
                          prefer_consolidated: bool = True) -> Dict:
    """
    감사보고서 엑셀에서 단일 연도 데이터 추출.
    반환: {'data': {...}, 'unit_scale': int, 'rcept_no': str, 'report_nm': str, 'debug': [...]}
    """
    report = find_audit_report(_dart, corp_code, year)
    if not report:
        return {"data": None, "error": f"{year}년 감사보고서/사업보고서를 찾지 못함"}

    excel_info = fetch_attach_excel_url(_dart, report["rcept_no"], prefer_consolidated)
    if not excel_info:
        return {"data": None, "error": "첨부파일에 재무제표 엑셀 없음", "rcept_no": report["rcept_no"]}

    title, url = excel_info
    content = download_file(url, _api_key)
    if not content:
        return {"data": None, "error": f"엑셀 다운로드 실패: {url}", "rcept_no": report["rcept_no"]}

    data, unit_scale, debug = parse_excel_for_accounts(content)
    return {
        "data": data,
        "unit_scale": unit_scale,
        "rcept_no": report["rcept_no"],
        "report_nm": report["report_nm"],
        "excel_title": title,
        "debug": debug,
    }


# ====================================================================
# 6) 통합 데이터 수집 (XBRL 우선 → 실패 시 엑셀)
# ====================================================================
def collect_multi_year_smart(_dart, _api_key: str, corp_code: str, corp_cls: str,
                            years: List[int], fs_div: str,
                            progress_callback=None) -> Tuple[Dict, Dict]:
    """
    상장사: XBRL을 3년 묶음으로 호출 (효율적)
    외감(E): 매년 엑셀 다운로드 시도 (어쩔 수 없이 매년 호출)
    """
    yearly_data = {y: {} for y in years}
    yearly_meta = {y: {} for y in years}  # source(XBRL/EXCEL), unit_scale, rcept_no 등

    prefer_consolidated = (fs_div == "CFS")

    # ============ 상장사 경로: XBRL 3년 묶음 ============
    if corp_cls in ("Y", "K", "N"):  # 유가/코스닥/코넥스
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
                progress_callback(idx, total, f"XBRL {base_year}년 (상장사 경로)")
            df = fetch_xbrl_finstate(_dart, corp_code, base_year, fs_div)
            if df is None:
                continue
            for offset in [0, 1, 2]:
                ty = base_year - offset
                if ty in years and not yearly_data[ty]:
                    yearly_data[ty] = extract_from_xbrl(df, year_offset=offset)
                    yearly_meta[ty] = {"source": "XBRL", "unit_scale": 1, "fs_div": fs_div}

        # XBRL이 모든 연도를 못 채운 경우 (외감으로 강등됐을 가능성)
        empty_years = [y for y in years
                       if all(v is None for v in yearly_data[y].values())]
        if empty_years and progress_callback:
            progress_callback(total, total + len(empty_years),
                              f"XBRL 누락 {len(empty_years)}년 → 엑셀 fallback")

        # 누락 연도는 엑셀로 보강
        for i, y in enumerate(empty_years):
            if progress_callback:
                progress_callback(total + i, total + len(empty_years), f"엑셀 fallback {y}년")
            result = fetch_audit_excel_data(_dart, _api_key, corp_code, y, prefer_consolidated)
            if result.get("data"):
                yearly_data[y] = result["data"]
                yearly_meta[y] = {
                    "source": "EXCEL",
                    "unit_scale": result.get("unit_scale", 1),
                    "rcept_no": result.get("rcept_no"),
                    "report_nm": result.get("report_nm"),
                    "excel_title": result.get("excel_title"),
                    "debug": result.get("debug"),
                }
            else:
                yearly_meta[y] = {"source": "FAILED", "error": result.get("error")}

    # ============ 외감 비상장사 경로: 매년 엑셀 ============
    else:  # E (외감) 등
        # 외감이라도 XBRL을 제출한 경우가 있으니 일단 시도
        if progress_callback:
            progress_callback(0, len(years) + 1, "외감: XBRL 1차 시도")
        df_latest = fetch_xbrl_finstate(_dart, corp_code, max(years), fs_div)
        xbrl_works = df_latest is not None and not df_latest.empty

        if xbrl_works:
            # XBRL 가능 - 3년 묶음으로 처리
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
                    progress_callback(idx, len(fetch_years), f"XBRL {base_year}년 (외감 XBRL 제출)")
                df = fetch_xbrl_finstate(_dart, corp_code, base_year, fs_div)
                if df is None:
                    continue
                for offset in [0, 1, 2]:
                    ty = base_year - offset
                    if ty in years and not yearly_data[ty]:
                        yearly_data[ty] = extract_from_xbrl(df, year_offset=offset)
                        yearly_meta[ty] = {"source": "XBRL", "unit_scale": 1, "fs_div": fs_div}

        # 빈 연도는 엑셀로
        empty_years = [y for y in years
                       if all(v is None for v in yearly_data[y].values())]
        for i, y in enumerate(empty_years):
            if progress_callback:
                progress_callback(i, len(empty_years), f"외감 엑셀 {y}년 (감사보고서 첨부)")
            result = fetch_audit_excel_data(_dart, _api_key, corp_code, y, prefer_consolidated)
            if result.get("data"):
                yearly_data[y] = result["data"]
                yearly_meta[y] = {
                    "source": "EXCEL",
                    "unit_scale": result.get("unit_scale", 1),
                    "rcept_no": result.get("rcept_no"),
                    "report_nm": result.get("report_nm"),
                    "excel_title": result.get("excel_title"),
                    "debug": result.get("debug"),
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

company_input = st.sidebar.text_input("회사명 또는 코드", placeholder="예: 삼성전자 / 005930")

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
            "ℹ️ 외감 비상장기업입니다. 감사보고서 첨부 엑셀에서 데이터를 추출합니다 "
            "(XBRL이 제출되었으면 그것을 우선 사용)."
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
                progress.progress(min(idx / total, 1.0))
            if label:
                status.text(f"📡 {label} ({idx+1}/{total})")

        yearly_data, yearly_meta = collect_multi_year_smart(
            dart, api_key, corp_code, corp_cls, years, fs_div_target,
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
                       if all(v is None for v in yearly_data[y].values())]
        if empty_years:
            st.warning(
                f"⚠️ 다음 연도는 데이터를 가져오지 못했습니다: {', '.join(map(str, empty_years))}\n\n"
                "원인은 위 데이터 소스 표의 '비고' 컬럼을 확인하세요."
            )

        st.caption(
            "⚠️ EBITDA = 영업이익 + (유형자산감가상각비 + 무형자산상각비 + 사용권자산상각비). "
            "주석 표기 상각비는 누락 가능. 외감 엑셀 파싱은 단위를 자동 감지하나 100% 정확 보장 안됨."
        )

        # 엑셀 다운로드
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            template_df.to_excel(writer, sheet_name="5개년_요약", index=False)
            # 원본 데이터
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
        with st.expander("🔬 엑셀 파싱 디버그 (외감 fallback 추적)"):
            for y in years:
                meta = yearly_meta.get(y, {})
                if meta.get("source") == "EXCEL":
                    st.markdown(f"**{y}년 — {meta.get('excel_title', '?')}**")
                    debug = meta.get("debug", [])
                    if debug:
                        st.dataframe(pd.DataFrame(debug), use_container_width=True, hide_index=True)

        st.divider()
        st.info(
            "🔜 다음 단계: PDF 본문 검증\n"
            "엑셀 fallback이 발동한 연도는 특히 PDF 원문 대조가 권장됩니다."
        )

st.sidebar.divider()
with st.sidebar.expander("ℹ️ 외감 데이터 처리 방식"):
    st.markdown(
        "**상장사(Y/K/N)**: XBRL API 우선 → 누락 연도만 엑셀 fallback\n\n"
        "**외감(E)**: XBRL 1차 시도 → 안되면 감사보고서 첨부 엑셀 자동 다운로드 + 파싱\n\n"
        "**단위 자동 감지**: 엑셀 본문에서 '백만원'/'천원' 표기 탐색해 자동 스케일링\n\n"
        "**연결/별도 선택**: 첨부 파일 제목에 '연결' 포함 여부로 우선순위"
    )

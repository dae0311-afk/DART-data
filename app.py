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

# v28: AgGrid — 셀 클릭 시 행 자동 선택을 위해 도입
try:
    from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode
    _AGGRID_AVAILABLE = True
except Exception:
    _AGGRID_AVAILABLE = False
    JsCode = None  # type: ignore


# v33: 구성표/요약표의 "N/A", "0", "(0)", "" 등 빈값 → "-" 표기 통일
def _df_dash_for_empty(df):
    """DataFrame 복사본을 반환하며 첨 컬럼을 제외한 나머지의 '빈값'을 '-'로 교체.
    빈값 정의: None, '', 'N/A', '0', '0.0', '(0)' 등.
    숫자형 컬럼(명시 안됨)이지만 파일은 이미 문자열로 포맷된 상태임(format_unit 적용 완료).
    """
    if df is None or df.empty:
        return df
    import re as _re
    out = df.copy()
    cols = list(out.columns[1:])
    # 0 판별 정규식: "0", "0.0", "0.00", "(0)", "(0.0)" 등
    _zero_pat = _re.compile(r"^\(?0+(?:\.0+)?\)?$")
    def _convert(v):
        if v is None:
            return "-"
        s = str(v).strip()
        if s == "" or s.upper() == "N/A":
            return "-"
        # 메타 설명 행("[···]" 등)은 원래부터 빈 문자열이므로 이미 "-" 돌아감
        if _zero_pat.match(s):
            return "-"
        return v
    for c in cols:
        out[c] = out[c].map(_convert)
    return out


# v34: 통합 재무 표 스타일러 — 정렬 + 0→dash + 합계행 회색 배경
def _df_style_finance(df, *, zero_to_dash=False, highlight_total_row=False, total_keywords=("합계",)):
    """재무 표 전용 Styler 생성.

    정렬 규칙:
      - 첫 컬럼(계정명/단위): 좌측
      - 나머지 컬럼(연도별 숫자): 우측
      - 헤더 — 첫 컬럼: 좌측, 나머지(연도): 중앙

    옵션:
      - zero_to_dash=True: 컬럼[1:] 의 0/N/A/'' → '-' 교체
      - highlight_total_row=True: total_keywords 중 하나가 포함된 행을 옥은 회색 배경
    """
    if df is None or df.empty:
        return df
    import re as _re
    work = df.copy()
    cols = list(work.columns)
    if len(cols) < 2:
        return work
    first_col = cols[0]
    other_cols = cols[1:]

    # zero → dash 교체 (첨 컬럼 제외)
    if zero_to_dash:
        _zero_pat = _re.compile(r"^\(?0+(?:\.0+)?\)?$")
        def _conv(v):
            if v is None:
                return "-"
            s = str(v).strip()
            if s == "" or s.upper() == "N/A":
                return "-"
            if _zero_pat.match(s):
                return "-"
            return v
        for c in other_cols:
            work[c] = work[c].map(_conv)

    try:
        styler = work.style
        # 본문 정렬: 첨 컬럼 좌측, 나머지 우측
        styler = styler.set_properties(subset=[first_col], **{"text-align": "left"})
        if other_cols:
            styler = styler.set_properties(subset=other_cols, **{"text-align": "right"})

        # 헤더 정렬 — 컬럼별 개별 적용
        table_styles = []
        # 첫 컬럼 헤더 → 좌측
        try:
            first_idx = work.columns.get_loc(first_col)
            table_styles.append({
                "selector": f"th.col_heading.col{first_idx}",
                "props": [("text-align", "left")],
            })
        except Exception:
            pass
        # 나머지(연도) 헤더 → 중앙
        for c in other_cols:
            try:
                idx = work.columns.get_loc(c)
                table_styles.append({
                    "selector": f"th.col_heading.col{idx}",
                    "props": [("text-align", "center")],
                })
            except Exception:
                pass
        if table_styles:
            styler = styler.set_table_styles(table_styles, overwrite=False)

        # 합계 행 회색 배경 — apply (행 단위 스타일)
        if highlight_total_row and total_keywords:
            def _row_style(row):
                first_val = str(row.iloc[0]) if len(row) > 0 else ""
                if any(kw in first_val for kw in total_keywords):
                    return ["background-color: #EFEFEF; font-weight: 600;"] * len(row)
                return ["" for _ in row]
            styler = styler.apply(_row_style, axis=1)

        return styler
    except Exception:
        return work


# v36: HTML 테이블 렌더 — st.dataframe(Styler) CSS가 Glide DataGrid에서 무시되어
# text-align이 적용되지 않는 문제를 우회. 직접 HTML을 만들어 st.markdown으로 렌더하면
# CSS 제어가 완전한다. 정렬/0→dash/합계행 하이라이트를 모두 여기서 처리.
def render_finance_html(df, *, zero_to_dash=True, highlight_total_row=False,
                         total_keywords=("합계",), table_id=None):
    """재무 표를 HTML로 렌더.

    정렬:
      - 첫 컬럼(계정명/단위): 좌측
      - 나머지 컬럼(연도별 숫자): 우측
      - 헤더 — 첫 컬럼: 좌측, 나머지(연도): 중앙

    올션:
      - zero_to_dash: 0/N/A/⓮ → '-'
      - highlight_total_row: 첫 컬럼에 total_keywords 포함된 행은 회색 배경 + bold
    """
    if df is None or df.empty:
        return
    import re as _re
    import html as _html
    cols = list(df.columns)
    first_col = cols[0]
    other_cols = cols[1:]
    _zero_pat = _re.compile(r"^\(?0+(?:\.0+)?\)?$")

    def _conv(v):
        if v is None:
            return "-"
        s = str(v).strip()
        if s == "" or s.upper() == "N/A":
            return "-"
        if _zero_pat.match(s):
            return "-"
        return s

    tid = table_id or "finance_tbl"
    css = f"""
    <style>
    .{tid}-wrap {{
        overflow-x: auto;
        -webkit-overflow-scrolling: touch;
        max-width: 100%;
        margin-bottom: 4px;
        border: 1px solid #d6dbe2;
        border-radius: 4px;
    }}
    .{tid} {{
        width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        background: #ffffff;
        font-size: 14px;
        font-family: "Source Sans Pro", -apple-system, system-ui, sans-serif;
    }}
    .{tid} th, .{tid} td {{
        padding: 7px 12px;
        border-bottom: 1px solid #e8ebef;
        line-height: 1.35;
        vertical-align: middle;
        white-space: nowrap;
    }}
    .{tid} thead th {{
        background: #fafbfc;
        font-weight: 600;
        color: #1E3D6B;
        border-bottom: 1px solid #d6dbe2;
        position: sticky;
        top: 0;
        z-index: 2;
    }}
    .{tid} th.first-col,
    .{tid} td.first-col {{
        text-align: left;
        position: sticky;
        left: 0;
        background: #ffffff;
        z-index: 1;
        box-shadow: 1px 0 0 #e8ebef;
    }}
    .{tid} thead th.first-col {{
        background: #fafbfc;
        z-index: 3;
    }}
    .{tid} tr.total-row td.first-col,
    .{tid} tr.total-row td {{ background: #EFEFEF; font-weight: 600; }}
    .{tid} th.num-col   {{ text-align: center; }}
    .{tid} td.num-col   {{ text-align: right; font-variant-numeric: tabular-nums; }}
    .{tid} tbody tr:last-child td {{ border-bottom: none; }}
    </style>
    """

    # <thead>
    head_cells = [f"<th class='first-col'>{_html.escape(str(first_col))}</th>"]
    for c in other_cols:
        head_cells.append(f"<th class='num-col'>{_html.escape(str(c))}</th>")
    thead = "<thead><tr>" + "".join(head_cells) + "</tr></thead>"

    # <tbody>
    body_rows = []
    for _, row in df.iterrows():
        first_val = str(row.iloc[0]) if len(row) > 0 else ""
        is_total = highlight_total_row and any(kw in first_val for kw in total_keywords)
        tr_class = " class='total-row'" if is_total else ""
        cells = [f"<td class='first-col'>{_html.escape(first_val)}</td>"]
        for c in other_cols:
            v = row[c]
            disp = _conv(v) if zero_to_dash else ("" if v is None else str(v))
            cells.append(f"<td class='num-col'>{_html.escape(disp)}</td>")
        body_rows.append(f"<tr{tr_class}>" + "".join(cells) + "</tr>")
    tbody = "<tbody>" + "".join(body_rows) + "</tbody>"

    table_html = (
        css
        + f"<div class='{tid}-wrap'><table id='{tid}' class='{tid}'>"
        + thead + tbody
        + "</table></div>"
    )
    st.markdown(table_html, unsafe_allow_html=True)


# v32: 모든 st.dataframe에 적용 — 숫자 컬럼 우측 정렬 (문자열 콤마 포함)
def _df_right_align_numbers(df):
    """DataFrame의 숫자적 컬럼(콤마/퍼센트 포함)을 우측 정렬해서 Styler 반환.
    st.dataframe(_df_right_align_numbers(df), ...)로 사용.
    """
    import re as _re
    if df is None or df.empty:
        return df
    # 숫자가 들어가는 컬럼 판별: 첫 컬럼(계정명)을 제외한 나머지
    cols_to_align = list(df.columns[1:])
    if not cols_to_align:
        return df
    try:
        styler = df.style.set_properties(
            subset=cols_to_align,
            **{"text-align": "right"}
        ).set_table_styles([
            {"selector": f"th.col_heading.level0",
             "props": [("text-align", "right")]},
        ], overwrite=False)
        # 헤더도 우측 정렬 (본문과 일치)
        for col in cols_to_align:
            styler = styler.set_table_styles(
                {col: [{"selector": "th", "props": [("text-align", "right")]}]},
                overwrite=False
            )
        return styler
    except Exception:
        return df

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
    "당기순이익": [
        "당기순이익", "당기순이익(손실)", "당기순손실(이익)", "당기순손실", "당기순손익",
        "당기의순이익", "당기의순손실", "당기의순이익(손실)",
        "법인세비용차감후순이익", "법인세비용차감후순손실",
        "법인세비용차감후당기순이익", "법인세비용차감후당기순손실",
        "법인세후순이익", "법인세후순손실",
        "분기순이익", "분기순손실", "반기순이익", "반기순손실",
        "연결당기순이익", "연결당기순손실",
    ],
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
    # v37: 산일전기 등 신규 상장사 변종 계정명 대응. 더 구체적인(긴) 키워드를 앞에 둠.
    "_유형자산감가상각비": [
        "유형자산감가상각비", "유형자산의감가상각비", "감가상각비",
        "유형자산상각비", "감가상각",
    ],
    "_무형자산상각비": [
        "무형자산상각비", "무형자산의상각비", "무형자산상각", "무형자산의상각",
    ],
    "_사용권자산상각비": [
        "사용권자산감가상각비", "사용권자산상각비", "사용권자산의상각비",
        "사용권자산상각", "사용권자산의상각",
    ],
    # v37: XBRL/HTML CF에 통합 표시("감가상각비및무형자산상각비")만 있는 경우 대응.
    # _유형/_무형이 모두 None일 때 _유형에 합산값을 채워넣는 폴백으로 사용.
    "_감가상각비_합산": [
        "유형자산감가상각비및무형자산상각비",
        "감가상각비및무형자산상각비",
        "감가상각비와무형자산상각비",
        "감가상각및무형자산상각비",
    ],
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
    "_감가상각비_합산": ["CF"],
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
    "_감가상각비_합산": "CF",
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
    # 한국 감사보고서의 음수 마커: △ ▲ ▽ ▼ (앞뒤/괄호 안 모두 가능)
    # 예: "△1,234", "▲(1,234)", "(△)1,234", "1,234△"
    neg_from_triangle = False
    for tri in ("△", "▲", "▽", "▼"):
        if tri in s:
            neg_from_triangle = True
            s = s.replace(tri, "")
    # 트라이앵글 제거 후 잔여 빈 괄호 ("()1,234" → "1,234") 정리
    s = re.sub(r"\(\s*\)|\[\s*\]", "", s)
    try:
        if s.startswith("(") and s.endswith(")"):
            s = "-" + s[1:-1]
        v = int(float(s))
        if neg_from_triangle and v > 0:
            v = -v
        return v
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


def _find_all_in_xbrl(df: pd.DataFrame, keywords: list, sj_filter: list) -> List[pd.Series]:
    """find_in_xbrl 대신 매칭되는 모든 row 반환 (account_nm 길이 짧은 순).

    K-IFRS XBRL은 같은 계정이 IS/CIS 등 여러 sj_div에 중복 등재되는 경우가 있고,
    그 중 일부 row는 frmtrm/bfefrmtrm_amount 가 비어있을 수 있다. 호출측이 행을
    순회하며 원하는 amount 컬럼이 채워진 첫 행을 고를 수 있도록 함.
    """
    if df is None or df.empty:
        return []
    work = df.copy()
    if "sj_div" in work.columns:
        work = work[work["sj_div"].isin(sj_filter)]
    if work.empty:
        return []
    exact_rows: List[pd.Series] = []
    for kw in keywords:
        m = work[work["account_nm"] == kw]
        for _, r in m.iterrows():
            exact_rows.append(r)
    pattern = "|".join([re.escape(k) for k in keywords])
    partial = work[work["account_nm"].astype(str).str.contains(pattern, na=False, regex=True)]
    partial = partial.copy()
    partial["_len"] = partial["account_nm"].astype(str).str.len()
    partial = partial.sort_values("_len")
    partial_rows = [r for _, r in partial.iterrows()]
    # 정확 매칭 행을 앞에 두고, 그 뒤에 부분 매칭. 중복(같은 인덱스) 제거.
    seen = set()
    out: List[pd.Series] = []
    for r in exact_rows + partial_rows:
        idx = (str(r.get("account_nm", "")), str(r.get("sj_div", "")), str(r.get("thstrm_amount", "")))
        if idx in seen:
            continue
        seen.add(idx)
        out.append(r)
    return out


def extract_from_xbrl(df: pd.DataFrame, year_offset: int = 0) -> Dict[str, Optional[int]]:
    if df is None or df.empty:
        return {key: None for key in ACCOUNT_KEYWORDS.keys()}
    amount_col = {0: "thstrm_amount", 1: "frmtrm_amount", 2: "bfefrmtrm_amount"}[year_offset]
    result = {}
    matched_acc = {}  # key → matched account_nm (중복 감지용)
    for key, keywords in ACCOUNT_KEYWORDS.items():
        sj = SJ_DIV_MAP.get(key, ["IS", "BS", "CIS", "CF"])
        # v38: 첫 매칭 행이 해당 offset의 amount가 비어있을 수 있어 모든 매칭 행을 순회.
        rows = _find_all_in_xbrl(df, keywords, sj)
        chosen_val = None
        chosen_acc = ""
        for r in rows:
            v = to_number(r.get(amount_col))
            if v is not None:
                chosen_val = v
                chosen_acc = str(r.get("account_nm", ""))
                break
        if chosen_val is None and rows:
            # 모든 매칭 행이 amount 비어있음 — 첫 행의 account_nm은 메타로 남김
            chosen_acc = str(rows[0].get("account_nm", ""))
        result[key] = chosen_val
        matched_acc[key] = chosen_acc

    # v37: D&A 합산 행 중복 매칭 보정.
    # 산일전기처럼 XBRL CF에 "감가상각비및무형자산상각비" 단일 행만 있으면
    # _유형/_무형/_합산이 모두 같은 행을 가리켜 3중 카운트됨 → 합산만 남기고 개별 None화.
    combined_acc = matched_acc.get("_감가상각비_합산", "")
    if combined_acc and result.get("_감가상각비_합산") is not None:
        if matched_acc.get("_유형자산감가상각비") == combined_acc:
            result["_유형자산감가상각비"] = None
        if matched_acc.get("_무형자산상각비") == combined_acc:
            result["_무형자산상각비"] = None
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
    # 당기순이익이 "지배기업소유주귀속순이익" / "주당순이익" 등에 잘못 매칭되지 않도록
    "당기순이익": [
        "지배기업", "비지배지분", "지배지분", "소유주",
        "주당", "기본주당", "희석주당", "주당순이익",
    ],
    "당기순손실": [
        "지배기업", "비지배지분", "지배지분", "소유주",
        "주당", "기본주당", "희석주당", "주당순손실",
    ],
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


_NET_INCOME_DENY = (
    "영업", "매출", "포괄", "지배", "비지배", "소유주",
    "주당", "희석", "기본",
    "세전", "법인세비용차감전", "법인세차감전", "차감전",
    "계속영업", "중단영업",
    "원천", "이자수익", "이자비용", "금융수익", "금융비용",
    "처분", "평가", "환산",
)


def _find_net_income_permissive(df: pd.DataFrame, year_offset: int = 0) -> Optional[int]:
    """IS 표 마지막 부분에서 '~순이익/~순손실'로 끝나는 행 직접 검색.

    풀무원다논 등 적자 외감 회사의 IS가 표준 키워드와 다른 라벨
    (예: '법인세비용차감후 당기순손실', '당기의 순손실', '계속·중단영업합 순손실')을
    쓰는 케이스 대응. 매칭 규칙:
      - 정규화 라벨이 '순이익' 또는 '순손실'을 포함
      - _NET_INCOME_DENY 토큰을 포함하지 않음 (영업/포괄/지배지분/주당/세전 등 배제)
      - 후보 중 가장 아래쪽 행(IS 최하단 = 진짜 당기순이익) 선택
    """
    if df is None or df.empty:
        return None
    year_cols = extract_year_columns(df)
    target_cols = year_cols.get(year_offset, [])
    if not target_cols:
        return None
    first_col = df.columns[0]
    matches = []
    for idx in range(len(df)):
        cell = df.iloc[idx][first_col]
        if pd.isna(cell):
            continue
        norm = normalize_account_name(str(cell))
        if not norm:
            continue
        if ("순이익" not in norm) and ("순손실" not in norm) and ("순손익" not in norm):
            continue
        if any(tok in norm for tok in _NET_INCOME_DENY):
            continue
        matches.append(idx)
    # 가장 아래쪽 행부터 시도 — 첫 유효 숫자 반환
    for idx in reversed(matches):
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


def find_da_from_notes(tables, tables_with_units: Optional[List[Tuple]] = None) -> Dict[str, Optional[Dict]]:
    """
    감사보고서의 모든 표에서 D&A 후보를 찾는다.
    v16: tables_with_units 인자가 주어지면 DOM 기반 단위 상속 결과를 우선 사용.
    그렇지 않으면 구 방식(이웃 표 텍스트 주변 검색) 적용.
    동일 키에 여러 후보가 있으면 value_in_won 기준 최대값 선택.
    """
    targets = {
        "감가상각비": ["감가상각비", "유형자산감가상각비"],
        "감가상각비_합산": [
            "감가상각비와무형자산상각비",
            "감가상각비및무형자산상각비",
            "유형자산감가상각비및무형자산상각비",
            "유형자산감가상각비와무형자산상각비",
            "감가상각및무형자산상각비",
            "감가상각비와기타상각비",
            "감가상각비및기타상각비",
            "감가상각비및무형자산상각",
            "감가상각비와무형자산상각",
        ],
        "무형자산상각비": ["무형자산상각비", "무형자산상각"],
        "사용권자산상각비": ["사용권자산상각비", "사용권자산감가상각비"],
    }
    # v37: 매칭 우선순위 — 긴 패턴부터(=더 구체적) 검사. substring 매칭으로 변경.
    # 예: "감가상각비및무형자산상각비합계" 셀은 합산 키와 매칭, "감가상각비" 단독 셀은 감가만 매칭.
    _matchers = []  # (normalized_pattern, key, length)
    for _k, _plist in targets.items():
        for _p in _plist:
            _pn = _p.replace(" ", "").replace("　", "")
            _matchers.append((_pn, _k, len(_pn)))
    _matchers.sort(key=lambda x: -x[2])
    # 누계액/잔액 등 BS 항목 오인 방지
    _NEGATIVE_LABEL_SUFFIXES = ("누계액", "잔액", "차감액", "장부금액")

    # tables_with_units가 제공되면 그것을 우선 사용
    # v18: 페어 (table, unit) 또는 트리플 (table, unit, period) 모두 허용
    period_per_table = {}
    if tables_with_units:
        tables = [pair[0] for pair in tables_with_units]
        unit_per_table = {i: pair[1] for i, pair in enumerate(tables_with_units)}
        for i, pair in enumerate(tables_with_units):
            if len(pair) >= 3:
                period_per_table[i] = pair[2]
            else:
                period_per_table[i] = None
    else:
        # 구 방식: 표별 단위 추정 (직전 3개 표 텍스트에서 '단위' 탐색)
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
            period_per_table[i] = None

    def col_norm(c):
        return str(c).replace(" ", "").replace("　", "")

    def col_is_current(cn: str) -> bool:
        # 당기 컬럼 검사 (전기 제외)
        if "전" in cn:
            return False
        return ("당기" in cn) or ("당)기" in cn) or ("당" in cn and "기" in cn)

    def col_is_disclosure(cn: str) -> bool:
        # v18: 2025+ 신규 포맷 '공시금액' / '금액' / '당기금액' 단일 컬럼
        if cn in ("공시금액", "금액", "당기금액"):
            return True
        if "공시금액" in cn:
            return True
        return False

    def col_is_total(cn: str) -> bool:
        # v18: '합계', '총계', '계', '자산합계' 등 총액 컬럼 (분해표 대응)
        if cn in ("합계", "총계", "계"):
            return True
        if "합계" in cn or "총계" in cn:
            return True
        return False

    candidates = {k: [] for k in targets.keys()}

    # v20: "비용의 성격별 분류" 표 식별자
    # 이 표의 '감가상각비'는 통상 유형 + 사용권자산 합산값이므로 사용권을 다시 더하면 더블카운팅 발생.
    #
    # 강한 신호(마커): 표 안에 성격별 표 헤더 문구가 존재.
    # 약한 신호: 성격별 표 특유의 행 라벨이 여러 개 등장. 단, CF 간접법 조정표는 "종업원급여부채"
    # "판매보증비(환입)" 등 일부 라벨을 공유하므로 negative markers로 명시 배제 필요.
    _NATURE_TABLE_MARKERS = (
        "성격별비용합계", "비용의성격별분류", "영업비용의성격별",
        "성격별영업비용", "비용의성격별",
    )
    # 결정적 양성 키워드: 성격별 분류 표에만 나타나는 라벨
    _NATURE_TABLE_STRONG_HINTS = (
        "재고자산의변동", "재고자산변동",
        "사용된원재료", "사용된원재료및상품", "원재료의사용",
        "원/부재료", "원재료매입",
        "외주가공비",
    )
    # 보조 양성 키워드: 성격별에도 자주 나오지만 CF 조정표 등에도 등장
    _NATURE_TABLE_WEAK_HINTS = (
        "급료와임금", "종업원급여", "종업원급여복리후생비",
        "퇴직급여", "판매보증비", "복리후생비",
    )
    # CF 간접법 조정표 / 이익잉여금처분계산서 등 비-성격별 표 부정 신호
    _NATURE_TABLE_NEGATIVE_MARKERS = (
        "영업활동관련자산", "영업활동으로인한자산", "영업으로창출된현금",
        "영업활동현금흐름", "영업활동으로인한현금흐름",
        "매출채권의감소", "매출채권의증가", "재고자산의감소",
        "미수금의감소", "미지급금의증가", "선수금의증가", "선수금의감소",
        "이익잉여금처분", "미처분이익잉여금",
        "종업원급여부채의증가", "종업원급여부채의감소",
    )

    def is_nature_table(t) -> bool:
        """표 내용 전체를 스캔해 성격별 분류 표인지 판단.

        결정 규칙 (v20 강화):
          1. negative marker 발견 시 즉시 False (CF 간접법 조정표 등 배제)
          2. positive marker 발견 시 True
          3. strong hint 1개 이상 + (strong+weak) 총 2개 이상이면 True
             (종업원급여 + 판매보증비 같은 weak 조합만으로는 False — CF 조정표와 충돌)
        """
        try:
            arr = t.fillna("").astype(str).values
        except Exception:
            return False
        joined = "".join("".join(row) for row in arr).replace(" ", "").replace("　", "")
        # (1) negative markers — CF 조정표/처분계산서 등
        for nm in _NATURE_TABLE_NEGATIVE_MARKERS:
            if nm in joined:
                return False
        # (2) positive markers
        for mk in _NATURE_TABLE_MARKERS:
            if mk in joined:
                return True
        # (3) 결정 키워드 1개 이상 + 총 hint 2개 이상
        strong_hits = sum(1 for h in _NATURE_TABLE_STRONG_HINTS if h in joined)
        weak_hits = sum(1 for h in _NATURE_TABLE_WEAK_HINTS if h in joined)
        return strong_hits >= 1 and (strong_hits + weak_hits) >= 2

    nature_table_flags: Dict[int, bool] = {}

    for i, t in enumerate(tables):
        # v18: 단일 row도 허용 (shape[0] < 1 만 제외)
        if t is None or t.shape[0] < 1 or t.shape[1] < 2:
            continue
        cols = [col_norm(c) for c in t.columns]
        col_text = " ".join(cols)
        # v37: '전기' 표 스킵을 보수적으로 — 표 컬럼에 '당기' 마커가 있으면
        # 상속된 prior 마커는 무시 (표 자체가 당기 값을 들고 있음).
        # 컬럼이 '공시금액'/'금액'만 있는 표(단일 컬럼)일 때만 상속 period가 의미 있음.
        if period_per_table.get(i) == "prior":
            has_current_col_marker = any(col_is_current(c) for c in cols)
            if not has_current_col_marker:
                continue
        # 변동 분석 표 (기초/증가/감소/기말) 제외
        if "기초" in col_text and "기말" in col_text:
            continue

        # value 컬럼 후보: 당기 컬럼 또는 (당기 마커 표인 경우) 공시금액 컬럼
        value_cols = []
        for c in t.columns:
            cn = col_norm(c)
            if col_is_current(cn):
                value_cols.append(c)
            elif col_is_disclosure(cn) and period_per_table.get(i) in ("current", None):
                value_cols.append(c)

        # v18: 당기 마커 표이면서 위 구조에 없었다면, '합계/총계' 총액 컬럼 허용
        # (사용권자산 분해표 같은 카테고리별 세분표 대응)
        if not value_cols and period_per_table.get(i) == "current":
            for c in t.columns:
                cn = col_norm(c)
                if col_is_total(cn):
                    value_cols.append(c)

        if not value_cols:
            continue

        # 라벨 컬럼 후보: value_cols 제외한 모든 컬럼 (카테고리 헤더형 표 대응)
        label_cols = [c for c in t.columns if c not in value_cols]
        if not label_cols:
            continue

        for idx in range(len(t)):
            # 모든 라벨 컬럼에서 매칭 시도 — substring + 긴 패턴 우선
            matched_key = None
            for lc in label_cols:
                cell = t.iloc[idx][lc]
                if pd.isna(cell):
                    continue
                cell_norm = str(cell).replace(" ", "").replace("　", "")
                if not cell_norm:
                    continue
                # BS 잔액/누계 표 셀은 D&A 비용 후보 아님 — 배제
                if any(cell_norm.endswith(sfx) for sfx in _NEGATIVE_LABEL_SUFFIXES):
                    continue
                for pat, key, _ in _matchers:
                    if pat in cell_norm:
                        matched_key = key
                        break
                if matched_key:
                    break
            if not matched_key:
                continue

            # value_cols에서 첫 유효 값 추출 (abs로 차감 표기 대응)
            for c in value_cols:
                v = to_number(t.iloc[idx][c])
                if v is not None:
                    v_abs = abs(v)
                    if i not in nature_table_flags:
                        nature_table_flags[i] = is_nature_table(t)
                    candidates[matched_key].append({
                        "table_idx": i,
                        "row_idx": idx,
                        "value": v_abs,
                        "unit_scale": unit_per_table[i],
                        "value_in_won": v_abs * unit_per_table[i],
                        "column": str(c),
                        "is_nature_table": nature_table_flags[i],
                    })
                    break

    # 각 키별 최대값 (가장 포괄적인 합계) 선택
    result = {}
    for key, items in candidates.items():
        if items:
            result[key] = max(items, key=lambda x: x["value_in_won"])
        else:
            result[key] = None

    # v20: 더블카운팅 방지 (개선된 판별)
    # 성격별 분류 표의 "감가상각비" 행이 유형+사용권자산 합산값인지,
    # 유형 단독값이고 사용권자산은 별도 행으로 나와 있는지 구분.
    #
    # 판별: 성격별 분류 표 내에 "사용권자산상각비" 행이 있는가?
    #   → 있다: 감가상각비 = 유형 단독 (예: 한솔) → 둘 다 유지
    #   → 없다: 감가상각비 = 유형+사용권 합산 (예: 한국석유공업)
    #          → 다른 표에서 찾은 사용권자산은 더블카운팅 → 제거
    #
    # 구현: 다시 해당 표를 스캔해 '사용권자산상각비' 라벨 존재 여부 확인.
    da = result.get("감가상각비")
    rou = result.get("사용권자산상각비")
    if da and da.get("is_nature_table") and rou:
        da_tbl_idx = da["table_idx"]
        rou_in_same_table = False
        try:
            t = tables[da_tbl_idx]
            arr = t.fillna("").astype(str).values
            for r in range(arr.shape[0]):
                for c in range(arr.shape[1]):
                    cell_norm = str(arr[r, c]).replace(" ", "").replace("　", "")
                    if cell_norm in ("사용권자산상각비", "사용권자산감가상각비"):
                        rou_in_same_table = True
                        break
                if rou_in_same_table:
                    break
        except Exception:
            pass

        # 성격별 표에 사용권 행이 없으면 감가상각비는 유형+사용권 합산 → 사용권 제거
        if not rou_in_same_table:
            result["사용권자산상각비"] = None
            result["_da_includes_rou"] = True

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


def fetch_audit_tables_with_units(sub_docs: pd.DataFrame, include_period: bool = False):
    """v16: 감사보고서 통합·주석 페이지에서 (table, unit_scale) 페어 리스트 반환.
    v18: include_period=True이면 (table, unit_scale, period) 트리플 리스트 반환.
          period는 'current' (당기) / 'prior' (전기) / None.

    이유: pd.read_html은 표 셀 내부만 읽어 표 사이 텍스트 노드의 '단위: 원/천원/백만원'
    마커를 놓치는 경우가 많음. BeautifulSoup으로 DOM 순회하며 텍스트 노드에 있는 마커를 상속.

    v18: '당기 (단위 : 원)' / '전기 (단위 : 원)' 키 마커도 추가로 상속.
    DART 2025년부터의 '공시금액' 단일 컴럼 포맷은 표 앞에 '당기'/'전기' 마커를 별도로 둘으므로
    이를 상속해야 어느 표가 당기 데이터인지 구분 가능.
    """
    if sub_docs is None or sub_docs.empty:
        return []
    out = []
    norm = sub_docs["title"].astype(str).str.replace(" ", "").str.replace("　", "")
    urls = []
    matched = sub_docs[norm.str.contains("재무제표", na=False)]
    urls.extend(matched.head(2)["url"].tolist())
    matched = sub_docs[norm == "주석"]
    urls.extend(matched.head(2)["url"].tolist())

    UNIT_RE = re.compile(r"단위\s*[:：]?\s*[\(\[]?\s*(백만\s*원|천\s*원|원)\b")

    def _unit_str_to_int(s: str) -> int:
        s_norm = s.replace(" ", "")
        if "백만원" in s_norm: return 1_000_000
        if "천원" in s_norm: return 1_000
        return 1

    try:
        from bs4 import BeautifulSoup
        import io as _io
    except Exception:
        # bs4 없으면 구한 경로로 폴백 (단위 1로 명시)
        for u in urls:
            try:
                for t in pd.read_html(u):
                    out.append((t, 1))
            except Exception:
                continue
        return out

    for u in urls:
        try:
            r = requests.get(u, timeout=15)
            r.encoding = r.encoding or "utf-8"
            html = r.text
        except Exception:
            continue
        soup = BeautifulSoup(html, "lxml")
        body = soup.find("body") or soup

        # v18: period 추적 (당기/전기). 상속 방식 — 마지막으로 등장한 마커 적용.
        PERIOD_RE = re.compile(r"(?<![가-힣])(당기|전기|당기말|전기말|당분기|전분기)(?![가-힣])")
        def _period_norm(s: str) -> Optional[str]:
            s = s.strip()
            if s in ("당기", "당기말", "당분기"):
                return "current"
            if s in ("전기", "전기말", "전분기"):
                return "prior"
            return None

        current_unit = 1
        current_period: Optional[str] = None
        for el in body.descendants:
            name = getattr(el, "name", None)
            if name is None:
                # 텍스트 노드—단위/period 마커 검색
                try:
                    txt = str(el)
                except Exception:
                    continue
                m = UNIT_RE.search(txt)
                if m:
                    current_unit = _unit_str_to_int(m.group(1))
                pm = PERIOD_RE.search(txt)
                if pm:
                    p = _period_norm(pm.group(1))
                    if p:
                        current_period = p
            elif name == "table":
                # 표 자체 텍스트에도 명시가 있을 수 있음
                t_txt = el.get_text(" ", strip=True)
                m = UNIT_RE.search(t_txt[:300])
                table_unit = _unit_str_to_int(m.group(1)) if m else current_unit
                # period: 표 앞동 300자 이내에 당기/전기 명시 자체 있으면 강제, 아니면 상속 값
                pm = PERIOD_RE.search(t_txt[:300])
                table_period = _period_norm(pm.group(1)) if pm else current_period
                # 표 파싱
                try:
                    sub_tables = pd.read_html(_io.StringIO(str(el)))
                    for t in sub_tables:
                        if include_period:
                            out.append((t, table_unit, table_period))
                        else:
                            out.append((t, table_unit))
                except Exception:
                    pass
    return out


# ====================================================================
# v17: 상장사 사업보고서 주석에서 D&A 보강
# XBRL CF는 "비현금항목의 조정"으로 D&A를 통합 표시하므로 개별 행이 없음.
# 사업보고서(kind='A')의 주석 페이지를 직접 파싱해 보강.
# ====================================================================
@st.cache_data(ttl=3600, show_spinner=False)
def find_annual_report_rcept(_dart, corp_code: str, year: int) -> Optional[str]:
    """해당 회계연도의 정기 사업보고서(kind='A') rcept_no.

    report_nm에 '사업보고서 (YYYY.MM)' 패턴이 있어야 하며, YYYY가 year와 일치.
    분기/반기 보고서는 제외.
    """
    try:
        start = f"{year+1}-01-01"
        end = f"{year+2}-06-30"
        reports = _dart.list(corp_code, start=start, end=end, kind="A")
        if reports is None or reports.empty:
            return None
        reports = reports.copy()

        def matches(nm):
            nm = str(nm)
            if "사업보고서" not in nm:
                return False
            m = re.search(r"\((\d{4})\.\d{2}\)", nm)
            if m:
                return int(m.group(1)) == year
            return False

        reports["_m"] = reports["report_nm"].apply(matches)
        matched = reports[reports["_m"]]
        if matched.empty:
            return None
        matched = matched.sort_values("rcept_dt", ascending=False)
        return str(matched.iloc[0]["rcept_no"])
    except Exception:
        return None


_EMPTY_DA = {"감가상각비": None, "무형자산상각비": None, "사용권자산상각비": None,
             "감가상각비_합산": None}


def _extract_da_from_sub_docs(sub_docs, prefer_consolidated: bool) -> Dict[str, Optional[Dict]]:
    """주어진 sub_docs(보고서 첨부문서 목록)에서 주석 페이지 → D&A 추출.

    v37: 주석에서 못 찾으면 통합 재무제표 페이지(현금흐름표 본문)도 탐색.
    """
    if sub_docs is None or sub_docs.empty:
        return dict(_EMPTY_DA)

    notes_mask = sub_docs["title"].astype(str).str.contains("주석", na=False)
    notes = sub_docs[notes_mask].copy()

    def priority(t):
        t = str(t)
        if prefer_consolidated and "연결재무제표 주석" in t:
            return 0
        if (not prefer_consolidated) and "재무제표 주석" in t and "연결" not in t:
            return 0
        if "연결재무제표 주석" in t:
            return 1
        if "재무제표 주석" in t:
            return 2
        return 9

    candidates = pd.DataFrame()
    if not notes.empty:
        notes["_p"] = notes["title"].apply(priority)
        notes = notes.sort_values("_p")
        candidates = notes.iloc[0:2][["title", "url"]]

    # v37: 주석 외에도 통합 재무제표 페이지 추가 — CF 본문에 D&A 개별 행이 있을 수 있음
    fs_mask = sub_docs["title"].astype(str).str.contains("재무제표", na=False) & ~notes_mask
    fs_pages = sub_docs[fs_mask].copy()
    if not fs_pages.empty:
        fs_pages = fs_pages.head(2)[["title", "url"]]
        candidates = pd.concat([candidates, fs_pages], ignore_index=True) if not candidates.empty else fs_pages

    if candidates.empty:
        return dict(_EMPTY_DA)

    try:
        triples = fetch_audit_tables_with_units(candidates, include_period=True)
        if not triples:
            return dict(_EMPTY_DA)
        tables_only = [p[0] for p in triples]
        return find_da_from_notes(tables_only, tables_with_units=triples)
    except Exception:
        return dict(_EMPTY_DA)


def _has_any_da(da_dict: Dict) -> bool:
    return any(
        da_dict.get(k) is not None and da_dict[k].get("value_in_won") is not None
        for k in ("감가상각비", "무형자산상각비", "사용권자산상각비", "감가상각비_합산")
    )


def fetch_da_from_business_report(_dart, corp_code: str, year: int,
                                  prefer_consolidated: bool = True) -> Dict[str, Optional[Dict]]:
    """상장사 D&A 보강 — 사업보고서(kind='A') 우선, 실패 시 감사보고서(kind='F') 폴백.

    v37: 두 소스 중 D&A가 검출되는 첫 번째 결과 반환.
      - 신규 상장사(예: 산일전기)는 상장 이전 회계연도(2023)에 사업보고서가 없고
        감사보고서만 존재 → 감사보고서 폴백으로 D&A 추출 가능.
      - 사업보고서 주석에서 D&A 추출 실패 시에도 감사보고서 시도.
    """
    # 1) 사업보고서 (kind='A')
    rcept = find_annual_report_rcept(_dart, corp_code, year)
    if rcept:
        sub_docs = get_sub_docs(_dart, rcept)
        result = _extract_da_from_sub_docs(sub_docs, prefer_consolidated)
        if _has_any_da(result):
            return result
    else:
        result = dict(_EMPTY_DA)

    # 2) 폴백: 감사보고서/연결감사보고서 (kind='F')
    audit_reports = find_external_audit_reports(_dart, corp_code, year)
    for rep in audit_reports:
        sub_docs = get_sub_docs(_dart, rep["rcept_no"])
        audit_result = _extract_da_from_sub_docs(sub_docs, prefer_consolidated)
        if _has_any_da(audit_result):
            return audit_result

    return result


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

    # 재무제표별 단위 개별 감지 (v16 fix: CF가 BS와 다른 단위일 경우 대응)
    # 감사보고서는 통상 BS/IS는 원 단위, CF는 천원 단위인 경우가 있음
    unit_scales = {
        "BS": detect_unit_from_header(urls["BS"]) if urls["BS"] else 1,
        "IS": detect_unit_from_header(urls["IS"]) if urls["IS"] else 1,
        "CF": detect_unit_from_header(urls["CF"]) if urls["CF"] else 1,
    }
    # 메이타에는 대표값으로 BS 단위를 높다 (하위호환)
    unit_scale = unit_scales["BS"] or 1

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
            # 단위 감지 (전체 페이지 표에서) - 통합 페이지는 보통 동일 단위
            try:
                all_tables = pd.read_html(combined_url)
                combined_unit = detect_unit_from_tables(all_tables)
                unit_scale = combined_unit
                for k in ("BS", "IS", "CF"):
                    unit_scales[k] = combined_unit
            except Exception:
                pass

    # 데이터 추출 (각 값을 대표 단위 unit_scale 기준으로 정규화)
    result_data = {}
    debug = []
    for key, keywords in ACCOUNT_KEYWORDS.items():
        stmt = STATEMENT_OF[key]
        df = dfs.get(stmt)
        val = find_account_in_html(df, keywords, year_offset=0)
        # v38: 당기순이익 keyword 미스 시 (1) IS 표에서 '순이익/순손실' 행 permissive 검색,
        # (2) 손익계산서 페이지에 당기순이익 행이 없으면 포괄손익계산서 페이지 추가 시도.
        # 풀무원다논 등 적자 외감 회사 라벨/페이지 변형 대응.
        if val is None and key == "당기순이익":
            if df is not None:
                val = _find_net_income_permissive(df, year_offset=0)
            if val is None:
                cis_url = find_statement_url(sub_docs, "포괄손익계산서")
                if cis_url and cis_url != urls.get("IS"):
                    cis_df = parse_html_statement(cis_url)
                    if cis_df is not None:
                        raw = find_account_in_html(cis_df, keywords, year_offset=0)
                        if raw is None:
                            raw = _find_net_income_permissive(cis_df, year_offset=0)
                        if raw is not None:
                            # 후속 stmt="IS" 단위 보정과 충돌 방지: CIS 단위 → IS 단위 환산해
                            # raw를 IS 표 기준 값처럼 저장. 외부 if 분기가 IS→대표 단위로 마저 환산.
                            cis_unit = detect_unit_from_header(cis_url) or 1
                            is_unit = unit_scales.get("IS") or 1
                            val = int(raw * cis_unit / is_unit) if is_unit else raw
        # v16 fix: 단위 불일치 보정 — 각 재무제표의 고유 단위로 원 환산 후
        # 대표 단위(unit_scale)로 다시 나눠 저장 → ev() 계산 일관성 확보
        if val is not None and stmt in unit_scales:
            stmt_unit = unit_scales[stmt] or 1
            if stmt_unit != unit_scale and unit_scale:
                # 원활산: val * stmt_unit → 추출값을 대표 단위 기준으로 재조정
                val_in_won = val * stmt_unit
                val = int(val_in_won / unit_scale)
        result_data[key] = val
        debug.append({
            "항목": key,
            "재무제표": stmt,
            "값(대표단위기준)": val,
            "표단위": unit_scales.get(stmt, 1),
            "대표단위": unit_scale,
            "발견여부": "✓" if val is not None else "—",
        })

    # ============================================================
    # D&A fallback: CF에서 감가상각비/무형자산상각비/사용권자산상각비가 누락되면
    # 주석 페이지를 포함한 모든 표에서 자동 탐색.
    # CF에 통합 표시(예: "현금의 유출이 없는 비용등의 가산")만 있는 경우 대응.
    # 주의: 주석 표의 값은 통상 천원 단위 → unit_scale 별도 적용.
    # ============================================================
    da_keys = ["_유형자산감가상각비", "_무형자산상각비", "_사용권자산상각비"]
    # v38: 값이 0인 경우도 누락으로 간주 — HTML CF 0값 행 매칭 케이스 (EBITDA=op_income 방지)
    da_missing = [k for k in da_keys if not result_data.get(k)]
    da_fallback_used = False
    da_fallback_info = {}
    if da_missing:
        # v16: DOM 기반 단위 상속 적용된 (table, unit_scale) 페어 수집
        # bs4 설치/HTML 파싱 실패 시 구 함수로 폴백
        try:
            audit_pairs = fetch_audit_tables_with_units(sub_docs)
        except Exception:
            audit_pairs = []
        if audit_pairs:
            all_audit_tables = [p[0] for p in audit_pairs]
            notes_da = find_da_from_notes(all_audit_tables, tables_with_units=audit_pairs)
        else:
            all_audit_tables = fetch_all_tables_from_audit(sub_docs)
            notes_da = find_da_from_notes(all_audit_tables) if all_audit_tables else {}
        if all_audit_tables:
            # 매핑: 주석 키 → result_data 키
            mapping = {
                "_유형자산감가상각비": "감가상각비",
                "_무형자산상각비": "무형자산상각비",
                "_사용권자산상각비": "사용권자산상각비",
            }
            for r_key, n_key in mapping.items():
                # v38: 0 값도 누락 취급 (HTML CF 0값 잘못 매칭 보정)
                if not result_data.get(r_key):
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
            # v38: 0 값도 누락 취급
            if all(not result_data.get(k) for k in ["_유형자산감가상각비", "_무형자산상각비"]):
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

        # v17: 상장사 D&A 보강 — XBRL CF에는 개별 감가/상각 행이 없으므로
        # 사업보고서 주석에서 별도 추출. 이미 HTML(외감 폴백) 경로에서 채워진
        # 연도는 건드리지 않음.
        da_missing_years = []
        for y in years:
            d = yearly_data.get(y) or {}
            # v37: _감가상각비_합산(XBRL 합산 행)이 있으면 EBITDA 계산 가능 → 보강 불필요
            # v38: D&A 값이 있더라도 합이 0이면 (XBRL '감가상각' 약한 키워드가 0값 행에
            # 매칭된 케이스 등) 의미 있는 D&A가 아니므로 주석 보강을 다시 시도.
            da_keys_listed = (
                "_유형자산감가상각비", "_무형자산상각비",
                "_사용권자산상각비", "_감가상각비_합산",
            )
            da_vals = [d.get(k) for k in da_keys_listed if d.get(k) is not None]
            if (not da_vals) or sum(da_vals) == 0:
                # 최소 한개 매출 또는 영업이익이 있어야 보강 의미 있음 (완전 실패 제외)
                if d.get("매출액") is not None or d.get("영업이익") is not None:
                    da_missing_years.append(y)

        for i, y in enumerate(da_missing_years):
            if progress_callback:
                progress_callback(i, len(da_missing_years), f"D&A 주석 보강 {y}년")
            da = fetch_da_from_business_report(_dart, corp_code, y, prefer_consolidated)
            updated = False
            if da.get("감가상각비") and da["감가상각비"].get("value_in_won") is not None:
                yearly_data[y]["_유형자산감가상각비"] = da["감가상각비"]["value_in_won"]
                updated = True
            if da.get("무형자산상각비") and da["무형자산상각비"].get("value_in_won") is not None:
                yearly_data[y]["_무형자산상각비"] = da["무형자산상각비"]["value_in_won"]
                updated = True
            if da.get("사용권자산상각비") and da["사용권자산상각비"].get("value_in_won") is not None:
                yearly_data[y]["_사용권자산상각비"] = da["사용권자산상각비"]["value_in_won"]
                updated = True
            # v37: 개별 감가/무형이 모두 누락이면 '감가상각비_합산' (감가+무형 합산 행) 폴백.
            # 산일전기 등 주석에 합산 행만 있는 케이스 → EBITDA 계산 실패 방지.
            # v38: 0 값도 누락으로 간주 (XBRL 0값 행 잘못 매칭 케이스 보정).
            if (not yearly_data[y].get("_유형자산감가상각비")
                and not yearly_data[y].get("_무형자산상각비")):
                combined = da.get("감가상각비_합산")
                if combined and combined.get("value_in_won") is not None:
                    yearly_data[y]["_유형자산감가상각비"] = combined["value_in_won"]
                    updated = True
            if updated:
                meta = yearly_meta.get(y, {})
                meta["da_source"] = "사업보고서 주석"
                yearly_meta[y] = meta

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
    # v38: 당기순이익 누락 (None 또는 0) — HTML 다중 폴백.
    # 상장사/외감 공통. 매출·영업이익은 잡혔는데 NI만 None인 케이스:
    #  - XBRL이 net income을 못 잡았거나 prior-year amount가 비어있음
    #  - 적자 표기 라벨(당기순손실 단독) keyword 미스
    # 전략:
    #  ① 해당 연도 자체 보고서 (extract_external_audit_data)
    #  ② 그래도 누락이면 최신 보고서의 prior/prior-prior 컬럼
    # ============================================================
    ni_missing = [
        y for y in years
        if yearly_data.get(y)
        and not yearly_data[y].get("당기순이익")  # None 또는 0
        and (yearly_data[y].get("매출액") is not None
             or yearly_data[y].get("영업이익") is not None)
    ]
    for y in ni_missing:
        if progress_callback:
            progress_callback(0, 1, f"당기순이익 HTML 보강 {y}년")
        try:
            ni_result = extract_external_audit_data(_dart, corp_code, y, prefer_consolidated)
        except Exception:
            ni_result = {}
        ni_val = (ni_result.get("data") or {}).get("당기순이익")
        if ni_val is not None and ni_val != 0:
            html_us = ni_result.get("unit_scale", 1) or 1
            # 이미 저장된 unit_scale 기준에 맞춰 환산 (저장 단위는 corp_type별로 다름)
            existing_us = (yearly_meta.get(y, {}) or {}).get("unit_scale", 1) or 1
            yearly_data[y]["당기순이익"] = int(ni_val * html_us / existing_us)
            meta = yearly_meta.get(y, {}) or {}
            meta["ni_html_fallback"] = True
            yearly_meta[y] = meta

    # ② 여전히 누락된 연도 — 최신 보고서의 multi-year IS 컬럼에서 추출.
    still_missing = [
        y for y in years
        if yearly_data.get(y) and not yearly_data[y].get("당기순이익")
        and (yearly_data[y].get("매출액") is not None
             or yearly_data[y].get("영업이익") is not None)
    ]
    if still_missing:
        latest_year = max(years)
        try:
            reports = find_external_audit_reports(_dart, corp_code, latest_year)
        except Exception:
            reports = []
        if reports:
            chosen = next((r for r in reports if r["is_consolidated"]), reports[0]) \
                if prefer_consolidated else \
                next((r for r in reports if not r["is_consolidated"]), reports[0])
            try:
                sub_docs = get_sub_docs(_dart, chosen["rcept_no"])
            except Exception:
                sub_docs = None
            if sub_docs is not None:
                is_url = find_statement_url(sub_docs, "손익계산서") \
                    or find_statement_url(sub_docs, "포괄손익계산서")
                is_df = parse_html_statement(is_url) if is_url else None
                if is_df is None:
                    combined_url = find_statement_url(sub_docs, "재무제표")
                    if combined_url:
                        classified = parse_combined_statements(combined_url)
                        is_df = classified.get("IS")
                        if is_url is None:
                            is_url = combined_url
                if is_df is not None and is_url is not None:
                    unit_scale_html = detect_unit_from_header(is_url) or 1
                    for y in still_missing:
                        offset = latest_year - y
                        if offset < 0 or offset > 2:
                            continue
                        v = find_account_in_html(is_df, ACCOUNT_KEYWORDS["당기순이익"], year_offset=offset)
                        if v is None:
                            v = _find_net_income_permissive(is_df, year_offset=offset)
                        if v is not None and v != 0:
                            existing_us = (yearly_meta.get(y, {}) or {}).get("unit_scale", 1) or 1
                            yearly_data[y]["당기순이익"] = int(v * unit_scale_html / existing_us)
                            meta = yearly_meta.get(y, {}) or {}
                            meta["ni_html_fallback"] = f"latest-report-offset-{offset}"
                            yearly_meta[y] = meta

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
        # v37: XBRL/HTML CF에 합산 행("감가상각비및무형자산상각비")만 있고 개별 항목이
        # 없는 경우(산일전기 등 신규 상장사 패턴) 합산값으로 폴백.
        # v38: combined == 0 (잘못 매칭된 0값 행)은 무시 → da를 None으로 유지
        if da is None:
            combined = _val_won(yearly_data, yearly_meta, y, "_감가상각비_합산")
            if combined:
                rou = _val_won(yearly_data, yearly_meta, y, "_사용권자산상각비")
                da = combined + (rou or 0)
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

    # ----- (2) v34: 주석 분해 cash-like 항목을 "포함" 헤더 없이 일반 항목과 동일 인덴트로 추가
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

    # v34: [기타금융자산 주석 분해 — 포함] 헤더 삭제, 하위 인덴트 제거
    for name in all_cash_names:
        row = [name]  # 인덴트 제거 (이전엔 f"　{name}")
        for y in sorted_years:
            won_v = cash_like_by_year[y].get(name)
            row.append(fu(won_v))
            if won_v is not None:
                totals_won[y] += won_v
                any_total[y] = True
        rows.append(row)

    # 합계 (v34: 인덴트 제거 — styler에서 회색 배경 적용 예정)
    total_row = ["합계"]
    for y in sorted_years:
        total_row.append(fu(totals_won[y]) if any_total[y] else "N/A")
    rows.append(total_row)

    # v34: [···비포함 · 참고] 섹션 전체 삭제 (합계 아래 아무것도 안 나오도록)

    columns = [f"(단위: {unit_label})"] + [str(y) for y in sorted_years]
    return pd.DataFrame(rows, columns=columns)


def build_debt_composition_table(yearly_data: Dict, yearly_meta: Dict, years: List[int],
                                 unit_label: str = "억원") -> pd.DataFrame:
    """차입금 구성표 (valuation Gross Debt 가산항 후보)."""
    return _build_composition_table(
        _DEBT_COMPOSITION_ITEMS, yearly_data, yearly_meta, years,
        total_label="합계",  # v34: 인덴트/설명 제거 — styler에서 회색 배경 적용
        unit_label=unit_label,
    )


_DA_COMPOSITION_ITEMS = [
    ("유형자산 감가상각비", "_유형자산감가상각비"),
    ("무형자산 상각비", "_무형자산상각비"),
    ("사용권자산 감가상각비", "_사용권자산상각비"),
]


def build_da_composition_table(yearly_data: Dict, yearly_meta: Dict, years: List[int],
                                unit_label: str = "억원") -> pd.DataFrame:
    """D&A 구성표 (EBITDA 가산항).

    구성 (v35):
    - 유형 감가상각비 / 무형 상각비 / 사용권자산 감가상각비 각 항목 행
    - 합계 (D&A 총합 = EBITDA 가산항)
    - v35: '추출 소스' 행 제거 — 사용자 요구사항 "합계 아래 삭제" 적용
      (소스 정보는 페이지 하단 '데이터 처리 방식' 안내에서 별도 확인 가능)
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

    for label, key in _DA_COMPOSITION_ITEMS:
        vals = {y: get_won(y, key) for y in sorted_years}
        rows.append([label] + [fu(vals[y]) for y in sorted_years])
        for y in sorted_years:
            if vals[y] is not None:
                totals_won[y] += vals[y]
                any_total[y] = True

    # 합계 행 (v35: 마지막 행으로 유지, 아래 추출 소스 행 제거)
    total_row = ["합계"]
    for y in sorted_years:
        total_row.append(fu(totals_won[y]) if any_total[y] else "N/A")
    rows.append(total_row)

    columns = [f"(단위: {unit_label})"] + [str(y) for y in sorted_years]
    return pd.DataFrame(rows, columns=columns)


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
    """막대(금액) + 라인(%) 콤보 차트 — v31:
      - x축 라인을 y=0 위치에 고정 (shape로 그림)
      - 연도 레이블은 annotation으로 차트 맨 아래 paper 좌표에 배치
      - 따라서 바 바닥 = y=0 = x축 완벽 일치 구조적 보장
      - 텍스트 14px, 모든 글자 검정
    """
    go = _import_plotly()
    if go is None:
        return None

    xs = [str(y) for y in sorted_years]
    ys_bar = [to_unit(bar_values_won.get(y), unit_label) for y in sorted_years]
    ys_line = [line_pct.get(y) for y in sorted_years]
    bar_text = [_format_unit_label(v, unit_label) if v is not None else "" for v in ys_bar]
    line_text = [_format_pct_label(v) for v in ys_line]

    BLACK = "#000000"
    FONT_SIZE = 14

    fig = go.Figure()

    # 바 — 기본 yaxis (좌측, 보이지 않음)
    fig.add_trace(go.Bar(
        x=xs, y=ys_bar, name=bar_name,
        marker_color=bar_color,
        text=bar_text, textposition="outside",
        textfont=dict(size=FONT_SIZE, color=BLACK),
        cliponaxis=False,
        yaxis="y",
        hovertemplate=f"%{{x}}<br>{bar_name}: %{{text}} {unit_label}<extra></extra>",
    ))

    # 라인 — 별도 yaxis2 (overlaying y, 보이지 않음)
    fig.add_trace(go.Scatter(
        x=xs, y=ys_line, name=line_name,
        mode="lines+markers+text",
        line=dict(color=line_color, width=2.5),
        marker=dict(size=9, color=line_color),
        text=line_text, textposition="top center",
        textfont=dict(size=FONT_SIZE, color=BLACK),
        cliponaxis=False,
        yaxis="y2",
        hovertemplate=f"%{{x}}<br>{line_name}: %{{text}}<extra></extra>",
    ))

    # 바 y축 range — v32: 바가 전체 차트의 2/3(=67%) 이하만 차지하도록 제한
    # range=[0, max*1.55]이면 바 max는 1/1.55 = 64.5% 위치 → 2/3 미만 보장
    _bar_vals = [v for v in ys_bar if v is not None]
    if _bar_vals:
        _b_max = max(_bar_vals)
        _b_min = min(_bar_vals + [0])
        if _b_min < 0:
            # 음수 포함 시: 전체 range를 양음 합쳐 잡고 양수쪽이 2/3 미만이 되도록
            _span = (_b_max - _b_min)
            bar_yrange = [_b_min - _span * 0.20, _b_max * 1.55]
        else:
            bar_yrange = [0, _b_max * 1.55]  # 바 최고점이 전체의 64.5% → 2/3 미만
    else:
        bar_yrange = [0, 1]

    # 라인 y축 range — v32: 라인이 항상 바차트+레이블 위(상단 25-30%)에만 그려지도록
    # 트릭: 라인 range의 하단을 매우 낮게 잡아 라인 도메인이 상단 좁은 구간에 위치
    _line_vals = [v for v in ys_line if v is not None]
    if _line_vals:
        _l_min = min(_line_vals)
        _l_max = max(_line_vals)
        _l_span = max(_l_max - _l_min, abs(_l_max) * 0.5, 5)
        # v37: 라벨이 잘리지 않도록 상단 여유 확장 (0.3 → 0.9). 라인 최고점은
        # (4.0 / 4.9) = 81.6% 위치 → 그 위 18.4%가 라벨 공간 (top center 텍스트)
        line_yrange = [_l_min - _l_span * 4.0, _l_max + _l_span * 0.9]
    else:
        line_yrange = [0, 1]

    # v31: x축을 아예 숨기고, y=0 위치에 축 라인을 shape로 그림.
    # 연도 레이블은 paper 좌표 기준 annotations로 맨 아래 고정 배치.
    # → 바 바닥과 x축 선이 물리적으로 일치 (다른 부동 공간 아예 없음)
    _x_axis_shape = dict(
        type="line",
        xref="paper", yref="y",
        x0=0, x1=1, y0=0, y1=0,
        line=dict(color=BLACK, width=1),
    )

    # v34: 차트 제목을 회색 둘근박스 배지로 표시 (annotation 활용)
    _title_badge = dict(
        x=0.5, y=1.12,
        xref="paper", yref="paper",
        text=f"<b>{title}</b>",
        showarrow=False,
        font=dict(size=16, color="#333333"),
        bgcolor="#F0F0F0",
        bordercolor="#F0F0F0",  # 테두리 없는 효과 (bgcolor와 동일)
        borderwidth=0,
        borderpad=8,  # 내부 패딩
        xanchor="center",
        yanchor="middle",
    )
    # 연도 레이블 — paper 좌표기준 하단에 균등 배치
    _year_annotations = []
    if xs:
        n = len(xs)
        # x 데이터 좌표과 paper 좌표의 관계: bargap=0.65에서 각 바 중심이
        # x 카테고리 접축의 중심이므로 data좌표 사용 (xref=x)
        for xv in xs:
            _year_annotations.append(dict(
                x=xv, y=0,
                xref="x", yref="paper",
                text=str(xv),
                showarrow=False,
                font=dict(size=FONT_SIZE, color=BLACK),
                yshift=-8,
                xanchor="center",
                yanchor="top",
            ))

    fig.update_layout(
        # v34: title을 비우고 annotation 배지로 제목 대체
        title=dict(text="", x=0.5, xanchor="center"),
        height=520,  # v34: 제목 배지 공간 확보
        margin=dict(l=40, r=40, t=110, b=60),  # v34: 상단 margin 더 넓게 — 배지 공간
        plot_bgcolor="white",
        paper_bgcolor="white",
        showlegend=True,
        legend=dict(orientation="h", yanchor="top", y=1.03, xanchor="center", x=0.5,
                    bgcolor="rgba(0,0,0,0)", font=dict(color=BLACK, size=FONT_SIZE)),
        bargap=0.325,  # v33: 이전 0.65 → 절반으로 줄임 (x축 항목 간격 축소, 바 폭 더 크게)
        font=dict(color=BLACK, size=FONT_SIZE),
        dragmode=False,
        xaxis=dict(
            visible=False,           # 기본 x축 숨김
            showgrid=False,
            fixedrange=True,
        ),
        yaxis=dict(
            visible=False, showgrid=False, zeroline=False,
            range=bar_yrange,
            fixedrange=True,
        ),
        yaxis2=dict(
            visible=False, showgrid=False, zeroline=False,
            overlaying="y", side="right",
            range=line_yrange,
            fixedrange=True,
        ),
        shapes=[_x_axis_shape],
        annotations=_year_annotations + [_title_badge],  # v34: 회색 제목 배지 추가
    )
    return fig


def _make_balance_chart(title: str, sorted_years: List[int], metrics: Dict, unit_label: str):
    """재무상태 5라인 차트. v29: 텍스트 14px 통일."""
    go = _import_plotly()
    if go is None:
        return None
    BLACK = "#000000"
    FONT_SIZE = 14
    series_def = [
        ("자산총계", "total_assets",      "#1E3D6B"),
        ("현금성자산", "cash_equiv",      "#7DB8E8"),
        ("부채총계", "total_liabilities", "#C7383C"),
        ("총차입금", "total_borrow",      "#F2A6A8"),
        ("자본총계", "total_equity",      "#2E9F5C"),
    ]
    fig = go.Figure()
    all_y_vals = []
    for label, key, color in series_def:
        xs, ys = _years_with_unit_values(metrics[key], sorted_years, unit_label)
        if not xs:
            continue
        text_labels = [_format_unit_label(v, unit_label) for v in ys]
        all_y_vals.extend([v for v in ys if v is not None])
        fig.add_trace(go.Scatter(
            x=xs, y=ys, name=label,
            mode="lines+markers+text",
            line=dict(color=color, width=2.5),
            marker=dict(size=9, color=color),
            text=text_labels, textposition="top center",
            textfont=dict(size=FONT_SIZE, color=BLACK),
            cliponaxis=False,
            hovertemplate=f"%{{x}}<br>{label}: %{{text}} {unit_label}<extra></extra>",
        ))
    # v37: 라벨이 plot 영역 밖으로 잘리지 않도록 y축 range에 상단 25% 여유 확보
    if all_y_vals:
        _y_max = max(all_y_vals)
        _y_min = min(all_y_vals + [0])
        _y_span = _y_max - _y_min
        balance_yrange = [_y_min - _y_span * 0.05, _y_max + _y_span * 0.25]
    else:
        balance_yrange = None
    # v34: 회색 둘근박스 제목 배지 (콤보 차트와 동일 스타일)
    _title_badge = dict(
        x=0.5, y=1.10,
        xref="paper", yref="paper",
        text=f"<b>{title}</b>",
        showarrow=False,
        font=dict(size=16, color="#333333"),
        bgcolor="#F0F0F0",
        bordercolor="#F0F0F0",
        borderwidth=0,
        borderpad=8,
        xanchor="center",
        yanchor="middle",
    )
    fig.update_layout(
        title=dict(text="", x=0.5, xanchor="center"),  # v34: 배지로 대체
        height=560,  # v34: 배지 공간 확보
        margin=dict(l=40, r=40, t=110, b=40),
        plot_bgcolor="white",
        paper_bgcolor="white",
        legend=dict(orientation="h", yanchor="top", y=1.02, xanchor="center", x=0.5,
                    bgcolor="rgba(0,0,0,0)", font=dict(color=BLACK, size=FONT_SIZE)),
        xaxis=dict(showgrid=False, showline=True, linecolor=BLACK, linewidth=1,
                   tickfont=dict(size=FONT_SIZE, color=BLACK),
                   fixedrange=True),  # v30: 줌 차단
        yaxis=dict(visible=False, showgrid=False, zeroline=False,
                   fixedrange=True,
                   range=balance_yrange),
        font=dict(color=BLACK, size=FONT_SIZE),
        dragmode=False,
        annotations=[_title_badge],  # v34: 제목 배지
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

/* 사이드바 옵션 라벨 (재무제표 구분 / 조회 기간 / 표시 단위) */
section[data-testid="stSidebar"] .opt-label {
    font-size: 0.92rem;
    color: #1f2937;
    margin: 14px 0 6px 0;
    font-weight: 500;
}
section[data-testid="stSidebar"] .opt-caption {
    font-size: 0.82rem;
    color: #9aa3af;
    margin-top: 14px;
}

/* 사이드바 segmented_control — 버튼 사이 간격 제거 (빈틈없이 붙임) */
section[data-testid="stSidebar"] div[data-testid="stSegmentedControl"] > div,
section[data-testid="stSidebar"] [data-baseweb="button-group"] {
    gap: 0 !important;
}
section[data-testid="stSidebar"] div[data-testid="stSegmentedControl"] button {
    margin: 0 !important;
    border-radius: 0 !important;
}
section[data-testid="stSidebar"] div[data-testid="stSegmentedControl"] button + button {
    border-left-width: 0 !important;
}
section[data-testid="stSidebar"] div[data-testid="stSegmentedControl"] button:first-of-type {
    border-top-left-radius: 8px !important;
    border-bottom-left-radius: 8px !important;
}
section[data-testid="stSidebar"] div[data-testid="stSegmentedControl"] button:last-of-type {
    border-top-right-radius: 8px !important;
    border-bottom-right-radius: 8px !important;
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

/* v27: 검색 결과 표 — st.dataframe(체크박스 선택) 기본 스타일 유지, 별도 CSS 없음 */

/* v38: '데이터 추출' 버튼 — 선택 전(disabled)은 옅은 붉은색, 선택 후는 기본 primary */
div.st-key-extract_btn_disabled button {
    background-color: #FEEFEF !important;
    border-color: #F5C2C7 !important;
    color: #842029 !important;
    opacity: 1 !important;
}
div.st-key-extract_btn_disabled button:hover,
div.st-key-extract_btn_disabled button:focus,
div.st-key-extract_btn_disabled button:active {
    background-color: #FEEFEF !important;
    border-color: #F5C2C7 !important;
    color: #842029 !important;
}
</style>
"""
st.markdown(_UI_CSS, unsafe_allow_html=True)

# ----- 좌측 사이드바: 조회 옵션 (segmented_control) -----
with st.sidebar:
    st.header("⚙️ 조회 옵션")
    fs_label = st.segmented_control(
        "재무제표 구분", options=["연결", "별도"], default="연결", key="fs_seg",
    )
    fs_div_target = "CFS" if fs_label == "연결" else "OFS"

    period_val = st.segmented_control(
        "조회 기간",
        options=[5, 10, 20, "최대"],
        format_func=lambda x: (f"{x}년" if isinstance(x, int) else x),
        default=5,
        key="period_seg",
    )
    period_label = f"{period_val}년" if isinstance(period_val, int) else period_val
    period_map = {"5년": 5, "10년": 10, "20년": 20, "최대": 99}

    unit_label = st.segmented_control(
        "표시 단위", options=["백만원", "억원", "십억원"], default="억원", key="unit_seg",
    )

    st.caption("옵션을 바꾸면 결과가 즉시 갱신됩니다.")

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
# v21: 엔터 키 입력 시 자동 검색 + 검색 결과 1건이면 자동 데이터 추출
st.markdown("<div class='hpe-section'>🔎 기업 검색</div>", unsafe_allow_html=True)

# 검색 영역 — st.form으로 감싸 엔터 입력 시 자동 submit
with st.form(key="search_form", clear_on_submit=False):
    fcol1, fcol2 = st.columns([6, 1.2])
    with fcol1:
        company_input = st.text_input(
            "회사명 또는 고유번호",
            placeholder="회사명을 입력하세요",
            label_visibility="collapsed",
            key="company_input",
        )
    with fcol2:
        search_btn = st.form_submit_button(
            "🔍 검색", use_container_width=True, type="primary"
        )

# 캐시 새로고침은 폼 밖에 (폼 submit과 분리)
refresh_cache_btn = st.button(
    "🔄 캐시 새로고침",
    help="사명이 업데이트 안될 때 눌러 corp_code.xml 재다운로드",
)

# 캐시 새로고침 (함수 정의 이후에서 처리)
if refresh_cache_btn:
    download_corp_code_xml.clear()
    search_companies.clear()
    st.success("캐시를 비웠습니다. 다시 검색해주세요.")

# 입력값 변경 자동 감지 (엔터 외에도 동작) — 단, 매번 자동 검색은 비효율이라 form submit만 사용
if search_btn:
    if not company_input.strip():
        st.warning("회사명 또는 코드를 입력하세요.")
        st.stop()
    with st.spinner(f"'{company_input}' 검색 중..."):
        companies = search_companies(dart, api_key, company_input)
    if companies.empty:
        st.error(
            f"'{company_input}'에 해당하는 회사를 찾을 수 없습니다.\n\n"
            "💡 사명이 최근 변경된 경우 위쪽 '🔄 캐시 새로고침' 버튼을 눌러보세요."
        )
        st.stop()
    st.session_state["companies"] = companies
    # v27: 이전 검색 잔존 상태 초기화 — 새 검색 시 선택 리셋
    st.session_state.pop("auto_extract", None)
    st.session_state.pop("company_table", None)
    st.session_state.pop("company_radio", None)
    st.session_state.pop("company_dataframe", None)
    st.session_state.pop("company_aggrid", None)
    st.session_state.pop("company_select_idx", None)
    # v32: AgGrid key를 완전히 새로 부여해 이전 선택 완전 리프레시
    st.session_state["search_counter"] = st.session_state.get("search_counter", 0) + 1

if "companies" in st.session_state and not st.session_state["companies"].empty:
    companies = st.session_state["companies"]
    st.markdown(
        f"<div class='hpe-section'>검색 결과 ({len(companies)}건) "
        f"<span style='font-size:0.85rem;font-weight:400;color:#6b7785;'>: 대상기업 클릭/터치</span></div>",
        unsafe_allow_html=True,
    )

    # v29: AgGrid — 체크박스 숨김 + 셀 테두리 제거 + 행 선택만 표시
    # - 컬럼: 회사명, 대표이사, 주소
    # - 셀 클릭으로 행 선택 (체크박스 UI 없이)
    # - 셀 단위 포커스 테두리 숨김
    # - 이전 st.dataframe 포맷(폰트/패딩/헤더 스타일)에 근접
    _display_col_map = {
        "corp_name": "회사명",
        "ceo_nm": "대표이사",
        "adres": "주소",
    }
    display_cols = [c for c in _display_col_map.keys() if c in companies.columns]
    if display_cols:
        df_show = companies[display_cols].rename(columns=_display_col_map).reset_index(drop=True)
    else:
        df_show = companies.reset_index(drop=True)

    # 원본 idx 매핑용 (선택된 행 → companies.iloc[idx])
    df_show_with_idx = df_show.copy()
    df_show_with_idx["_row_idx"] = range(len(df_show_with_idx))

    selected_idx = None

    if _AGGRID_AVAILABLE:
        # v31: AgGrid 스타일을 st.dataframe과 거의 동일하게 맞춤
        # - 폰트 14px, 행 높이 30px, 셀 패딩 좁게
        # - 줄바꿈 없이 포함 수 있는 최소 열 폭 (autoSize)
        # - 세로 선/했 구분선 st.dataframe 스타일로 단순화
        st.markdown("""
            <style>
            /* v31: AgGrid → st.dataframe 스타일로 맞춤 */
            .ag-cell-focus, .ag-cell-no-focus, .ag-cell {
                border: none !important;
                outline: none !important;
            }
            .ag-cell:focus {
                border: none !important;
                outline: none !important;
            }
            .ag-theme-streamlit, .ag-theme-streamlit .ag-cell,
            .ag-theme-streamlit .ag-header-cell-text {
                font-family: "Source Sans Pro", "Noto Sans KR", -apple-system, sans-serif !important;
                font-size: 14px !important;
            }
            .ag-theme-streamlit .ag-cell {
                padding-left: 8px !important;
                padding-right: 8px !important;
                line-height: 30px !important;
            }
            .ag-theme-streamlit .ag-header-cell {
                padding-left: 8px !important;
                padding-right: 8px !important;
            }
            .ag-theme-streamlit .ag-header-cell-text {
                font-weight: 600 !important;
                color: rgba(49, 51, 63, 0.95) !important;
            }
            /* 행 선택 하이라이트 — st.dataframe 선택 색과 유사 */
            .ag-theme-streamlit .ag-row-selected {
                background-color: #FFE6E6 !important;
            }
            /* 세로 구분선 제거 (st.dataframe은 더 단순) */
            .ag-theme-streamlit .ag-cell {
                border-right: none !important;
            }
            </style>
        """, unsafe_allow_html=True)

        # v34: 케파일 셀/헤더 둘 다 14px 동일 폰트
        _cell_style_14 = {"font-size": "14px", "line-height": "30px",
                          "padding-left": "8px", "padding-right": "8px",
                          "border-right": "none"}
        _header_style_14 = {"font-size": "14px", "font-weight": "600",
                            "padding-left": "8px", "padding-right": "8px"}

        gb = GridOptionsBuilder.from_dataframe(df_show_with_idx)
        # 체크박스 없이 셀 클릭만으로 선택
        gb.configure_selection(
            selection_mode="single",
            use_checkbox=False,
            header_checkbox=False,
            rowMultiSelectWithClick=False,
            suppressRowDeselection=False,
        )
        # 컬럼 설정 — cellStyle/headerStyle 모두 부여 (v34: 헤더 폰트 통일)
        for _col in ["회사명", "대표이사", "주소"]:
            gb.configure_column(
                _col,
                wrapText=False,
                autoHeight=False,
                resizable=True,
                minWidth=80,
                cellStyle=_cell_style_14,
                headerStyle=_header_style_14,
            )
        gb.configure_column("_row_idx", hide=True)
        # 주소 컬럼: minWidth=200 설정 하지만, JS에서 콘텐츠폭 < container폭일 때 잔여공간 채움
        gb.configure_column("주소", wrapText=False, autoHeight=False, minWidth=120,
                            cellStyle=_cell_style_14, headerStyle=_header_style_14)

        # v33: 자동피팅 강화 — 이브릿지(결과 소량) vs 삼성전자(결과 다량, 가상스크롤) 둘 다 작동하도록
        # 1. setTimeout 100ms 지연 → 가상 스크롤 셀이 렌더링 완료 후 측정
        # 2. autoSize 후 컬럼 합계 폭 < grid 폭이면 마지막 컬럼(주소)을 남은 공간만큼 늘림
        #    → 표 내부가 테두리보다 좁아지는 현상 해소
        # 3. onGridReady + onFirstDataRendered + onGridSizeChanged 세 훅 이벤트 모두 트리거
        # v34: 자동피팅 강화 — (1) autoSize → (2) 컬럼 합계 < container 폭일 때 잔여공간을 마지막 컬럼에 채움
        # 이전 이슈: "한솔오리온텍" 의 주소가 짧음 → minWidth=200이어도 grid 전체 폭보다 작아 테두리에 빈공간 생김
        # 이 JS가 주소컬럼 폭을 "container폭 - 다른컬럼합" 으로 강제 설정
        # v36: 자동피팅 — 사용자 요구에 맞습所 "콘텐츠 폭으로 자동 맞춤 + 잔여공간은 주소칸이 흡수"
        # v35 sizeColumnsToFit은 컬럼을 균등 분배(380/379/379)해서 사용자 요구와 부합 불일치.
        # v34 방식 복원 + .ag-root-wrapper를 직접 찾는 방식으로 수정 (params.api.getGui()가
        # 좀은 element 반환하는 경우 대비).
        _auto_size_js = JsCode("""
        function(params) {
            var doFit = function() {
                try {
                    // 1단계: autoSize (콘텐츠 폭에 맞습所 컬럼 폭 자동 계산)
                    var allColumnIds = [];
                    var cols = params.api.getColumns ? params.api.getColumns() : (params.columnApi ? params.columnApi.getAllColumns() : []);
                    if (!cols || cols.length === 0) return;
                    cols.forEach(function(column) {
                        var cid = column.getColId ? column.getColId() : column.colId;
                        if (cid && cid !== '_row_idx') allColumnIds.push(cid);
                    });
                    if (allColumnIds.length === 0) return;

                    if (params.api.autoSizeColumns) {
                        try { params.api.autoSizeColumns(allColumnIds, false); } catch(_) {}
                    } else if (params.columnApi && params.columnApi.autoSizeColumns) {
                        try { params.columnApi.autoSizeColumns(allColumnIds, false); } catch(_) {}
                    }

                    // 2단계: container 폭 측정 — .ag-root-wrapper 직접 찾기
                    var rootWrap = document.querySelector('.ag-root-wrapper');
                    var gridWidth = rootWrap ? rootWrap.clientWidth : 0;
                    if (gridWidth <= 0) {
                        console.log('[v36 autoFit] gridWidth=0, abort');
                        return;
                    }

                    // 3단계: 실제 컬럼 폭 합계
                    var widthMap = {};
                    var totalCol = 0;
                    cols.forEach(function(column) {
                        var cid = column.getColId ? column.getColId() : column.colId;
                        if (cid && cid !== '_row_idx') {
                            var w = column.getActualWidth ? column.getActualWidth() : 0;
                            widthMap[cid] = w;
                            totalCol += w;
                        }
                    });

                    var available = gridWidth - 2;  // 테두리 1px×2
                    console.log('[v36 autoFit] grid=' + gridWidth + ', totalCol=' + totalCol + ', available=' + available);

                    // 4단계: 컬럼 합 < container 폭 → 마지막 컬럼(주소) 잔여공간 흡수
                    if (totalCol < available) {
                        var lastColId = allColumnIds[allColumnIds.length - 1];
                        var lastW = widthMap[lastColId] || 0;
                        var newLastW = lastW + (available - totalCol);
                        console.log('[v36 autoFit] expanding ' + lastColId + ' from ' + lastW + ' to ' + newLastW);
                        var done = false;
                        if (params.api.setColumnWidth) {
                            try { params.api.setColumnWidth(lastColId, newLastW); done = true; } catch(e) { console.log('setColumnWidth err', e); }
                        }
                        if (!done && params.columnApi && params.columnApi.setColumnWidth) {
                            try { params.columnApi.setColumnWidth(lastColId, newLastW); done = true; } catch(_) {}
                        }
                        if (!done && params.api.applyColumnState) {
                            try {
                                params.api.applyColumnState({
                                    state: [{ colId: lastColId, width: newLastW }],
                                    applyOrder: false,
                                });
                                done = true;
                                console.log('[v36 autoFit] applyColumnState used');
                            } catch(e) { console.log('applyColumnState err', e); }
                        }
                        if (!done) console.log('[v36 autoFit] all setWidth methods failed');
                    }
                } catch(e) {
                    console.log('[v36 autoFit] error', e);
                }
            };
            // 세 시점: 가상스크롤/소량/대량 모두 커버
            setTimeout(doFit, 80);
            setTimeout(doFit, 350);
            setTimeout(doFit, 900);
        }
        """)
        gb.configure_grid_options(
            suppressCellFocus=True,
            rowSelection="single",
            suppressRowClickSelection=False,
            rowHeight=30,
            headerHeight=32,
            onFirstDataRendered=_auto_size_js,
            onGridReady=_auto_size_js,
            onGridSizeChanged=_auto_size_js,  # v33: 추가
        )
        grid_options = gb.build()

        # v32: search_counter로 key 동적 부여 → 새 검색 시 완전 리프레시
        _search_count = st.session_state.get("search_counter", 0)
        _aggrid_key = f"company_aggrid_{_search_count}"

        grid_response = AgGrid(
            df_show_with_idx,
            gridOptions=grid_options,
            update_mode=GridUpdateMode.SELECTION_CHANGED,
            data_return_mode=DataReturnMode.AS_INPUT,
            fit_columns_on_grid_load=False,  # v32: 자동 피팅 적용 위해 해제
            allow_unsafe_jscode=True,
            height=min(50 + 30 * max(len(df_show_with_idx), 1), 400),
            theme="streamlit",
            key=_aggrid_key,
        )

        # 검색 결과 1건이면 자동 선택
        if len(companies) == 1:
            selected_idx = 0
        else:
            sel_rows = grid_response.get("selected_rows")
            try:
                if sel_rows is None:
                    selected_idx = None
                elif isinstance(sel_rows, pd.DataFrame):
                    if not sel_rows.empty and "_row_idx" in sel_rows.columns:
                        selected_idx = int(sel_rows.iloc[0]["_row_idx"])
                elif isinstance(sel_rows, list) and len(sel_rows) > 0:
                    selected_idx = int(sel_rows[0].get("_row_idx", 0))
            except Exception:
                selected_idx = None
    else:
        # AgGrid 미설치 시 폴백 — 기존 st.dataframe 체크박스 방식
        event = st.dataframe(
            df_show,
            use_container_width=True,
            hide_index=True,
            on_select="rerun",
            selection_mode="single-row",
            key="company_dataframe",
        )
        if len(companies) == 1:
            selected_idx = 0
        else:
            try:
                sel_rows = event.selection.rows  # type: ignore[attr-defined]
            except Exception:
                sel_rows = []
            selected_idx = sel_rows[0] if sel_rows else None

    if selected_idx is None:
        st.caption("⬆️ 표에서 회사를 클릭해 선택하세요. (셀 또는 체크박스 클릭)")
        st.button("데이터 추출", type="secondary", use_container_width=True, disabled=True, key="extract_btn_disabled")
        st.stop()

    selected_corp = companies.iloc[selected_idx]
    corp_code = selected_corp.get("corp_code")
    corp_cls = selected_corp.get("corp_cls", "")
    corp_name = selected_corp.get("corp_name", "")

    # v30: 외감 비상장기업 안내 메시지 제거 (사용자 요구)

    extract_btn = st.button("데이터 추출", type="primary", use_container_width=True, key="extract_btn_active")

    # v23: 자동 추출 로직 제거 — 항상 버튼 클릭으로만 실행 (사용자 요구)
    st.session_state.pop("auto_extract", None)

    # v36 이슈2: 옵션 변경 자동 감지 — 동일 회사에서 fs/period 변경 시 자동 재추출
    _ext = st.session_state.get("extracted")
    auto_refresh_needed = False
    if (not extract_btn) and _ext and _ext.get("corp_code") == corp_code:
        _snap = _ext.get("snapshot", {}) or {}
        if (_snap.get("fs_div_target") != fs_div_target
                or _snap.get("period_label") != period_label
                or _snap.get("end_year") != end_year):
            auto_refresh_needed = True

    # 회사가 바뀌면 기존 추출 결과 폐기
    if _ext and _ext.get("corp_code") != corp_code:
        st.session_state.pop("extracted", None)
        _ext = None

    if extract_btn or auto_refresh_needed:
        period_n = period_map[period_label]
        start_year = 2015 if period_n == 99 else max(2015, end_year - period_n + 1)
        years = list(range(start_year, end_year + 1))

        progress = st.progress(0.0)
        status = st.empty()
        if auto_refresh_needed:
            status.text("⚙️ 옵션 변경 감지 — 데이터 재추출 중...")

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

        # session_state 저장
        st.session_state["extracted"] = {
            "corp_code": corp_code,
            "corp_name": corp_name,
            "corp_cls": corp_cls,
            "years": years,
            "start_year": start_year,
            "end_year": end_year,
            "yearly_data": yearly_data,
            "yearly_meta": yearly_meta,
            "selected_corp": selected_corp.to_dict() if hasattr(selected_corp, "to_dict") else dict(selected_corp),
            "snapshot": {
                "fs_div_target": fs_div_target,
                "period_label": period_label,
                "end_year": end_year,
            },
        }

    # v36 이슈2: 추출 결과가 session_state에 있으면 항상 렌더 (옵션 변경 즉시 반영)
    _ext = st.session_state.get("extracted")
    if _ext and _ext.get("corp_code") == corp_code:
        yearly_data = _ext["yearly_data"]
        yearly_meta = _ext["yearly_meta"]
        years = _ext["years"]
        start_year = _ext["start_year"]
        end_year = _ext["end_year"]
        selected_corp = _ext.get("selected_corp", selected_corp)

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

        # v36: 연결/별도가 연도별 다를 때 칩 표시 변경
        # 고유 fs 값들 수집
        cfs_years = sorted([y for y, v in per_year_fs.items() if v == "CFS"])
        ofs_years = sorted([y for y, v in per_year_fs.items() if v == "OFS"])
        if cfs_years and ofs_years:
            # 혼합 — 연도 지정 표시 (연도 많으면 "연결(N개) + 별도(M개)")
            if len(cfs_years) + len(ofs_years) <= 8:
                fs_chip_val = (
                    f"연결 {','.join(map(str, cfs_years))} · 별도 {','.join(map(str, ofs_years))}"
                )
            else:
                fs_chip_val = f"연결 {len(cfs_years)}개 · 별도 {len(ofs_years)}개"
        else:
            fs_chip_val = f"{main_fs_label}재무제표"

        # -------- 헤더 둥근 박스 4개 --------
        period_disp = f"{start_year}~{end_year}"
        pill_html = (
            "<div class='info-pill-row'>"
            f"<div class='info-pill pill-primary'><span class='pill-key'>회사</span> <span class='pill-val'>{corp_name}</span></div>"
            f"<div class='info-pill'><span class='pill-key'>구분</span> <span class='pill-val'>{fs_chip_val}</span></div>"
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
        # v31: 원래 st.dataframe 형태로 원복 (사용자 요구)
        # v32: 수직 스크롤 없이 전체 표시 + 숫자 우측 정렬
        st.markdown("<div class='hpe-section'>요약 재무제표</div>", unsafe_allow_html=True)
        template_df = build_template_table(yearly_data, yearly_meta, years, unit_label=unit_label)
        # v36: HTML 테이블 렌더 — st.dataframe은 Styler text-align이 무시됨
        render_finance_html(template_df, zero_to_dash=True, table_id="finance_summary")
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
            # v30: 차트 인터랙션 비활성화 — 드래그/줌/클릭 모두 차단
            _chart_config = {
                "staticPlot": True,
                "displayModeBar": False,
                "scrollZoom": False,
                "doubleClick": False,
                "showAxisDragHandles": False,
                "showAxisRangeEntryBoxes": False,
            }
            for key in chart_order:
                fig = charts.get(key)
                if fig is not None:
                    st.plotly_chart(fig, use_container_width=True, config=_chart_config)
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
        da_comp_df = build_da_composition_table(yearly_data, yearly_meta, years, unit_label=unit_label)

        # v34: 좌우 2컬럼 배치 제거 — 현금성자산 → 차입금 → D&A 수직 배치
        st.markdown("<div class='hpe-section'>현금성자산 구성</div>", unsafe_allow_html=True)
        # v36: HTML 테이블로 렌더 (우측정렬 보장)
        render_finance_html(cash_comp_df, zero_to_dash=True, highlight_total_row=True,
                            table_id="finance_cash")
        st.caption(
            "· 요약표 '현금성자산' = 이 구성표 합계.\n"
            "· 기타유동·비유동금융자산은 외감사 통합 계정: 정기예금·MMF·단기금융상품 외에 대여금·보증금·파생상품 포함 가능 → 주석 확인 후 가감 필요.\n"
            "· 사용제한·담보 항목 존재 가능 → valuation 시 차감항 조정 필요."
        )

        st.markdown("<div class='hpe-section'>차입금 구성</div>", unsafe_allow_html=True)
        # v36: HTML 테이블로 렌더
        render_finance_html(debt_comp_df, zero_to_dash=True, highlight_total_row=True,
                            table_id="finance_debt")
        st.caption(
            "· 요약표 '총차입금' = 이 구성표 합계 (리스부채, CB/BW/EB 포함).\n"
            "· IFRS 16 미적용 기업은 리스부채 0 → valuation 시 별도 조정 필요."
        )

        # -------- D&A 구성표 --------
        st.markdown("<div class='hpe-section'>D&A 구성 (EBITDA 가산항)</div>", unsafe_allow_html=True)
        # v36: HTML 테이블로 렌더
        render_finance_html(da_comp_df, zero_to_dash=True, highlight_total_row=True,
                            table_id="finance_da")
        st.caption(
                "· EBITDA = 영업이익 + D&A 합계 (유형 감가상각비 + 무형 상각비 + 사용권자산 감가상각비).\n"
                "· '추출 소스' 행: 상장사 XBRL CF는 개별 D&A가 없어 사업보고서 주석에서 보강(`XBRL + 사업보고서 주석`). 외감 비상장사는 감사보고서 CF/주석에서 직접 추출.\n"
                "· N/A 연도는 해당 항목이 주석에서 매칭되지 않은 경우 (표제목/라벨 이질 또는 사용권자산 미공시).\n"
                "· 사용권자산 감가상각비를 유형자산에 합산 표기하는 기업이 있어 행별 년도별 0 여부와 무관하게 합계는 올바른 값 가능."
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
            da_comp_df.to_excel(writer, sheet_name="D&A_구성", index=False)
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

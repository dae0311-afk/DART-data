# DART 재무정보 추출 에이전트

DART OpenAPI 기반으로 회사의 사업보고서/감사보고서에서 재무 데이터를 추출하는 Streamlit 앱.

## 현재 버전 (v2 — API 기반)

- 회사명/종목코드/고유번호로 검색 (동명회사 구분 지원)
- 사업보고서/반기/분기보고서 선택
- 연결재무제표(CFS)·별도재무제표(OFS) **동시 조회**로 혼동 방지
- 매출액·영업이익 자동 추출 (키워드 기반 정확/부분 매칭)
- 단위 변환 (원/천원/백만원/억원)
- 전년 대비 증감액·증감률 자동 계산
- 공시 목록 메타정보(rcept_no) 표시 — PDF 검증 단계 준비
- 추출 결과 + 원본 전체 재무제표 + 공시목록 엑셀 다운로드

## 다음 버전 (v3 — PDF 검증, 개발 예정)

- 접수번호 기반 감사보고서/연결감사보고서 PDF 자동 다운로드
- PDF 본문 표 추출 후 API 값과 자동 대조
- 불일치 시 경고 + PDF 원문 페이지 미리보기

## 사용법

1. https://opendart.fss.or.kr 에서 인증키 발급 (40자리)
2. 좌측 사이드바에 키 입력
3. 회사 검색 → 동명회사 중 정확한 회사 선택 → 재무 추출

## Streamlit Cloud 배포

1. 이 저장소를 본인 GitHub에 Public으로 푸시
2. share.streamlit.io 에서 New app
3. Repository / Branch / Main file: `app.py` 지정
4. Advanced settings → Secrets에 추가:
   ```
   DART_API_KEY = "발급받은_40자리_키"
   ```
5. Deploy

## 로컬 실행

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 한계 명시

- 비상장 외감기업(corp_cls=E)은 OpenAPI 데이터 부재/부분 제공 가능
- 금융업은 `finstate_all` 제외
- 계정명 매칭은 키워드 기반이므로 특수업종은 PDF 본문 대조 필수

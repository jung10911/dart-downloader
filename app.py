import streamlit as st
import OpenDartReader
import pandas as pd
import io
import time

# 1. 페이지 기본 설정
st.set_page_config(page_title="IR 전문 대량 데이터 수집기", layout="wide")
st.title("📊 Open DART 대량 기업 분석 & 개황 수집기")
st.markdown("수백 개의 종목코드를 한 번에 입력하여 기업 정보와 재무 데이터를 엑셀로 추출합니다.")
st.markdown("---")

# 2. 사이드바: API 키 설정
st.sidebar.header("🔧 서비스 설정")
api_key_input = st.sidebar.text_input(
    "DART API 키 입력", 
    value="0d3337714983152206d438906ff525000677118e", 
    type="password"
)

# 3. 사용자 입력 섹션
st.subheader("1. 대상 기업 입력")
st.info("종목코드를 줄바꿈(엔터) 또는 쉼표(,)로 구분하여 입력하세요. 엑셀에서 복사해서 바로 붙여넣어도 됩니다. (예: 005930, 000660)")
corp_codes_input = st.text_area("종목코드 다중 입력 (300개 이상 가능)", height=150, placeholder="005930\n000660\n035420")

st.subheader("2. 수집 조건 설정")
col1, col2, col3 = st.columns(3)
with col1:
    data_type = st.radio("수집할 데이터 선택", ["기업개황만", "재무제표만", "기업개황 + 재무제표 모두"])
with col2:
    year = st.selectbox("재무제표 조회 연도", options=[2023, 2022, 2021, 2020], index=0)
with col3:
    report_type = st.selectbox(
        "보고서 종류", 
        options=["11011", "11012", "11013", "11014"],
        format_func=lambda x: {"11011":"사업보고서", "11012":"반기보고서", "11013":"1분기보고서", "11014":"3분기보고서"}[x]
    )

st.markdown("---")

# 4. 데이터 수집 실행
if st.button("🚀 대량 데이터 수집 시작"):
    if not api_key_input:
        st.error("API 키를 입력해주세요.")
    elif not corp_codes_input.strip():
        st.error("종목코드를 하나 이상 입력해주세요.")
    else:
        # 입력된 종목코드 정리 (쉼표, 줄바꿈 제거 및 리스트화)
        raw_codes = corp_codes_input.replace(',', '\n').split('\n')
        corp_codes = [code.strip() for code in raw_codes if code.strip()]
        corp_codes = list(set(corp_codes)) # 중복 제거
        
        st.success(f"총 {len(corp_codes)}개의 종목코드가 인식되었습니다. 수집을 시작합니다!")
        
        # 진행 상태 표시 준비
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        dart = OpenDartReader(api_key_input)
        
        company_results = []
        finance_results = []
        error_codes = []

        # 300개 이상의 기업을 순회하며 데이터 수집
        for i, code in enumerate(corp_codes):
            status_text.text(f"수집 중... ({i+1}/{len(corp_codes)}) : {code}")
            
            # [1] 기업개황 수집
            if "기업개황" in data_type:
                try:
                    comp_info = dart.company(code) # 기업개황 API 호출 (딕셔너리 형태 반환)
                    if comp_info:
                        company_results.append(comp_info)
                except Exception:
                    error_codes.append(f"{code} (개황 실패)")

            # [2] 재무제표 수집
            if "재무제표" in data_type:
                try:
                    fin_info = dart.fin_stat(code, year, report_type)
                    if fin_info is not None and not fin_info.empty:
                        # 숫자형 변환
                        if 'thstrm_amount' in fin_info.columns:
                            fin_info['thstrm_amount'] = pd.to_numeric(fin_info['thstrm_amount'].astype(str).str.replace(',', ''), errors='coerce')
                        finance_results.append(fin_info)
                except Exception:
                    if f"{code} (개황 실패)" not in error_codes:
                        error_codes.append(f"{code} (재무 실패)")
            
            # API 호출 속도 조절 (과부하 방지)
            time.sleep(0.1) 
            
            # 진행률 바 업데이트
            progress_bar.progress((i + 1) / len(corp_codes))

        status_text.text("데이터 수집 완료! 엑셀 파일을 생성합니다.")

        # 5. 결과 정리 및 엑셀 다운로드
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            
            # 기업개황 데이터 엑셀 시트 생성
            if company_results:
                df_company = pd.DataFrame(company_results)
                df_company.to_excel(writer, index=False, sheet_name='기업개황')
                st.subheader("🏢 기업개황 수집 결과")
                st.dataframe(df_company.head(), use_container_width=True)
            
            # 재무제표 데이터 엑셀 시트 생성
            if finance_results:
                df_finance = pd.concat(finance_results, ignore_index=True)
                df_finance.to_excel(writer, index=False, sheet_name='재무제표')
                st.subheader("💰 재무제표 수집 결과")
                st.dataframe(df_finance.head(), use_container_width=True)

        # 에러 발생한 코드 안내
        if error_codes:
            st.warning(f"일부 종목은 데이터가 없거나 수집에 실패했습니다: {', '.join(error_codes)}")

        # 최종 엑셀 파일 다운로드 버튼
        st.download_button(
            label="📥 전체 데이터 엑셀 파일 다운로드",
            data=output.getvalue(),
            file_name=f"DART_대량수집_결과.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

import streamlit as st
import OpenDartReader
import pandas as pd
import io

# 1. 페이지 설정 및 제목
st.set_page_config(page_title="IR 전문 데이터 수집기", layout="wide")
st.title("📊 Open DART 재무 데이터 추출 서비스")
st.markdown("---")

# 2. 사이드바: 설정 및 API 키 관리
st.sidebar.header("🔧 서비스 설정")

# 보안을 위해 API 키를 입력받거나 Streamlit secrets에서 가져옴
api_key_input = st.sidebar.text_input(
    "DART API 키 입력", 
    value="0d3337714983152206d438906ff525000677118e", # 기본값으로 설정
    type="password",
    help="오픈다트에서 발급받은 인증키를 입력하세요."
)

# 3. 사용자 검색 조건 입력
col1, col2, col3 = st.columns(3)

with col1:
    corp_code = st.text_input("종목코드 (6자리)", value="005930", help="예: 삼성전자는 005930")

with col2:
    year = st.selectbox("조회 연도", options=[2023, 2022, 2021, 2020], index=0)

with col3:
    report_type = st.selectbox(
        "보고서 종류", 
        options=["11011", "11012", "11013", "11014"],
        format_func=lambda x: {"11011":"사업보고서", "11012":"반기보고서", "11013":"1분기보고서", "11014":"3분기보고서"}[x]
    )

# 4. 데이터 수집 및 실행
if st.button("데이터 불러오기 및 분석 🚀"):
    if not api_key_input:
        st.error("API 키를 입력해주세요.")
    else:
        try:
            with st.spinner('DART 서버에서 데이터를 가져오는 중입니다...'):
                # API 연결
                dart = OpenDartReader(api_key_input)
                
                # 재무제표 수집
                df = dart.fin_stat(corp_code, year, report_type)
                
                if df is not None and not df.empty:
                    # 데이터 전처리: 금액 컬럼 숫자화
                    df['thstrm_amount'] = pd.to_numeric(df['thstrm_amount'].str.replace(',', ''), errors='coerce')
                    df['pastrm_amount'] = pd.to_numeric(df['pastrm_amount'].str.replace(',', ''), errors='coerce')
                    
                    # 핵심 지표 표시 (탭 구분)
                    tab1, tab2 = st.tabs(["📋 데이터 미리보기", "📉 주요 지표 요약"])
                    
                    with tab1:
                        st.subheader(f"조회 결과: {df['corp_name'].iloc[0]} ({year}년)")
                        st.dataframe(df[['fs_nm', 'account_nm', 'thstrm_nm', 'thstrm_amount']], use_container_width=True)
                    
                    with tab2:
                        # 주요 항목 필터링 (매출, 영업이익, 당기순이익 등)
                        summary_items = ['매출액', '영업이익', '당기순이익']
                        summary_df = df[df['account_nm'].isin(summary_items)]
                        if not summary_df.empty:
                            st.table(summary_df[['account_nm', 'thstrm_nm', 'thstrm_amount']])
                        else:
                            st.info("주요 요약 항목을 필터링할 수 없습니다. 전체 데이터를 확인하세요.")

                    # 5. 엑셀 다운로드 기능 (중요)
                    # 메모리에 엑셀 파일을 생성
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        df.to_excel(writer, index=False, sheet_name='Sheet1')
                    
                    excel_data = output.getvalue()
                    
                    st.download_button(
                        label="📥 엑셀 파일로 다운로드 (딸깍)",
                        data=excel_data,
                        file_name=f"{corp_code}_{year}_DART_Data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.success("데이터 준비 완료! 위 버튼을 눌러 다운로드하세요.")
                else:
                    st.warning("해당 조건의 데이터가 존재하지 않습니다.")
        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")

st.markdown("---")
st.caption("제공: IR 컨설턴트를 위한 자동화 도구")

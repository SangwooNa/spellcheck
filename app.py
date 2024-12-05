import streamlit as st
import openpyxl
import pandas as pd
import requests

# 맞춤법 검사 함수
def check_spelling(text):
    """네이버 맞춤법 검사 API 호출"""
    url = "https://m.search.naver.com/p/csearch/dcontent/spellchecker.nhn"
    params = {"_callback": "mycallback", "q": text}
    response = requests.get(url, params=params)
    if response.status_code == 200:
        corrected_text = response.text.split('"checked":')[1].split('"')[1]
        return corrected_text
    return text  # API 호출 실패 시 원본 반환

def process_excel(file):
    """엑셀 파일 읽기"""
    try:
        wb = openpyxl.load_workbook(file)
        ws = wb.active
        if ws is None:
            raise ValueError("엑셀 파일에 활성화된 워크시트가 없습니다.")
        data = []
        for row in ws.iter_rows(values_only=True):
            data.append([cell if cell is not None else "" for cell in row])  # 빈 셀 처리
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"엑셀 파일 처리 중 오류 발생: {e}")
        return pd.DataFrame()  # 빈 데이터프레임 반환

def perform_spell_check(df):
    """데이터프레임의 텍스트 셀에 대해 맞춤법 검사"""
    corrected_df = df.copy()
    for col in df.columns:
        corrected_df[col] = df[col].apply(lambda x: check_spelling(x) if isinstance(x, str) else x)
    return corrected_df

# Streamlit 앱 UI
st.title("엑셀 맞춤법 검사기")
st.write("업로드한 엑셀 데이터를 웹 상에 표시하고, 맞춤법 검사를 수행합니다.")

# 엑셀 파일 업로드
uploaded_file = st.file_uploader("엑셀 파일 업로드", type=["xlsx"])

if uploaded_file:
    # 엑셀 데이터 불러오기
    df = process_excel(uploaded_file)

    if not df.empty:
        # 원본 데이터 테이블 표시
        st.write("### 업로드된 데이터")
        st.dataframe(df, use_container_width=True)

        # 맞춤법 검사 버튼
        if st.button("맞춤법 검사"):
            st.write("맞춤법 검사를 수행 중입니다...")
            corrected_df = perform_spell_check(df)

            # 맞춤법 검사 결과 표시
            st.write("### 맞춤법 검사 결과")
            st.dataframe(corrected_df, use_container_width=True)

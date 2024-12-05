import streamlit as st
import openpyxl
import pandas as pd
import requests
from io import BytesIO

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
    """엑셀 파일 읽기 및 맞춤법 검사"""
    wb = openpyxl.load_workbook(file)
    ws = wb.active

    original_data = []
    corrected_data = []

    for row in ws.iter_rows(values_only=True):
        original_row = []
        corrected_row = []
        for cell in row:
            if cell and isinstance(cell, str):
                corrected = check_spelling(cell)
                original_row.append(cell)
                corrected_row.append(corrected)
            else:
                original_row.append(cell)
                corrected_row.append(cell)
        original_data.append(original_row)
        corrected_data.append(corrected_row)

    return original_data, corrected_data

def highlight_differences(original, corrected):
    """수정이 필요한 부분에 강조 표시"""
    styled_data = []
    for orig_row, corr_row in zip(original, corrected):
        styled_row = []
        for orig, corr in zip(orig_row, corr_row):
            if orig != corr:
                # 수정이 필요한 경우 강조
                styled_row.append(f"<span style='color: red;'>{corr}</span>")
            else:
                styled_row.append(orig)
        styled_data.append(styled_row)
    return styled_data

# Streamlit 앱 UI
st.title("엑셀 맞춤법 검사기")
st.write("업로드한 엑셀 파일을 테이블로 표현하고 수정된 부분을 강조합니다.")

# 파일 업로드
uploaded_file = st.file_uploader("엑셀 파일 업로드", type=["xlsx"])

if uploaded_file:
    # 엑셀 파일 처리
    original_data, corrected_data = process_excel(uploaded_file)

    # 데이터프레임 생성
    original_df = pd.DataFrame(original_data)
    corrected_df = pd.DataFrame(corrected_data)

    # 수정된 데이터를 HTML로 표시
    styled_corrected_data = highlight_differences(original_data, corrected_data)

    # 원본 데이터 테이블
    st.write("### 원본 데이터")
    st.dataframe(original_df, use_container_width=True)

    # 수정된 데이터 테이블 (강조된 부분 포함)
    st.write("### 수정된 데이터")
    styled_corrected_df = pd.DataFrame(styled_corrected_data)
    st.write(styled_corrected_df.to_html(escape=False, index=False), unsafe_allow_html=True)

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
    """엑셀 파일 읽기 및 맞춤법 검사"""
    wb = openpyxl.load_workbook(file)
    ws = wb.active

    data = []
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
        data.append({"Original": original_row, "Corrected": corrected_row})
    return data

# Streamlit 앱 UI
st.title("엑셀 맞춤법 검사기")
st.write("업로드된 데이터와 수정된 데이터를 웹에서 확인하세요.")

# 파일 업로드
uploaded_file = st.file_uploader("엑셀 파일 업로드", type=["xlsx"])

if uploaded_file:
    st.write("파일을 처리 중입니다...")
    data = process_excel(uploaded_file)

    # 결과 테이블 생성
    st.write("맞춤법 검사 결과:")
    for i, row in enumerate(data):
        st.markdown(f"### Row {i + 1}")
        original = row["Original"]
        corrected = row["Corrected"]

        for j, (orig, corr) in enumerate(zip(original, corrected)):
            if orig != corr:
                # 수정이 필요한 부분 강조
                st.markdown(
                    f"- **Cell {j + 1}:** {orig} → <span style='color:red;'>{corr}</span>",
                    unsafe_allow_html=True,
                )
                # 팝업 링크
                if st.button(f"Cell {j + 1} 수정 보기 (Row {i + 1})", key=f"{i}-{j}"):
                    st.write(f"**원본 텍스트:** {orig}")
                    st.write(f"**수정된 텍스트:** {corr}")
            else:
                st.markdown(f"- **Cell {j + 1}:** {orig}")


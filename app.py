import streamlit as st
import openpyxl
import requests
from io import BytesIO

# 맞춤법 검사 함수
def check_spelling(text):
    """네이버 맞춤법 검사 API 호출"""
    url = "https://m.search.naver.com/p/csearch/dcontent/spellchecker.nhn"
    params = {"_callback": "mycallback", "q": text}
    response = requests.get(url, params=params)
    if response.status_code == 200:
        # 응답에서 수정된 텍스트 추출
        corrected_text = response.text.split('"checked":')[1].split('"')[1]
        return corrected_text
    return text  # API 호출 실패 시 원본 반환

def process_excel(file):
    """엑셀 파일 처리 및 맞춤법 검사"""
    wb = openpyxl.load_workbook(file)
    ws = wb.active

    # 맞춤법 검사 및 수정
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                cell.value = check_spelling(cell.value)

    # 수정된 엑셀 파일 반환
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# Streamlit UI
st.title("엑셀 맞춤법 검사기")
st.write("업로드한 엑셀 파일의 텍스트 데이터를 맞춤법 검사하고 수정된 파일을 다운로드하세요.")

# 파일 업로드
uploaded_file = st.file_uploader("엑셀 파일 업로드", type=["xlsx"])

if uploaded_file:
    # 엑셀 파일 처리 및 결과 표시
    st.write("파일을 처리 중입니다...")
    corrected_file = process_excel(uploaded_file)

    # 수정된 파일 다운로드 버튼
    st.download_button(
        label="수정된 파일 다운로드",
        data=corrected_file,
        file_name="corrected.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

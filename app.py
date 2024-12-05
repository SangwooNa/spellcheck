import streamlit as st
import openpyxl
from hanspell import spell_checker
from io import BytesIO

# 맞춤법 검사 함수
def check_spelling(text):
    result = spell_checker.check(text)
    return result.checked if result.errors else text

def process_excel(file):
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

st.title("엑셀 맞춤법 검사기")
uploaded_file = st.file_uploader("엑셀 파일 업로드", type=["xlsx"])

if uploaded_file:
    st.write("파일을 처리 중입니다...")
    corrected_file = process_excel(uploaded_file)

    st.download_button(
        label="수정된 파일 다운로드",
        data=corrected_file,
        file_name="corrected.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

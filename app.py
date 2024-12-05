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
    """엑셀 파일 읽기 및 데이터프레임 생성"""
    try:
        wb = openpyxl.load_workbook(file)
        ws = wb.active
        if ws is None:
            raise ValueError("엑셀 파일에 활성화된 워크시트가 없습니다.")
        data = []
        for row in ws.iter_rows(values_only=True):
            # 빈 셀은 빈 문자열("")로 처리
            data.append([cell if cell is not None else "" for cell in row])
        # 데이터프레임 생성 및 모든 값을 문자열로 변환
        return pd.DataFrame(data).astype(str)
    except Exception as e:
        st.error(f"엑셀 파일 처리 중 오류 발생: {e}")
        return pd.DataFrame()  # 빈 데이터프레임 반환

def perform_spell_check(df):
    """데이터프레임의 텍스트 셀에 대해 맞춤법 검사"""
    corrected_df = df.copy()
    error_count = 0
    progress_bar = st.progress(0)  # 진행률 바 추가
    total_cells = df.size
    processed_cells = 0

    # 현재 진행 중인 문자열 표시를 위한 공간
    current_text_placeholder = st.empty()

    for col in df.columns:
        for idx, value in df[col].items():
            if isinstance(value, str):  # 텍스트 셀만 검사
                current_text_placeholder.text(f"현재 검사 중인 텍스트: '{value}'")  # 현재 텍스트 표시
                corrected_value = check_spelling(value)
                corrected_df.at[idx, col] = corrected_value
                if corrected_value != value:  # 수정이 필요한 경우
                    error_count += 1
            processed_cells += 1
            progress_bar.progress(processed_cells / total_cells)  # 진행률 업데이트

    # 모든 작업 완료 후 메시지 제거
    current_text_placeholder.text("")
    return corrected_df, error_count

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
        if st.button("맞춤법 검사 시작"):
            st.write("맞춤법 검사를 수행 중입니다...")
            corrected_df, error_count = perform_spell_check(df)

            # 맞춤법 검사 결과 표시
            st.write("### 맞춤법 검사 결과")
            st.dataframe(corrected_df, use_container_width=True)
            st.success(f"맞춤법 검사 완료! 총 {error_count}개의 오류를 발견했습니다.")

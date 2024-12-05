import pandas as pd
import streamlit as st

def preprocess_excel(file):
    """엑셀 데이터 전처리"""
    try:
        # 엑셀 파일 읽기
        df = pd.read_excel(file, header=None)

        # 헤더 찾기: "번 호", "성 명", "학 년" 포함 행
        header_row = df.apply(lambda row: row.astype(str).str.contains("번 호|성 명|학 년", na=False).any(), axis=1).idxmax()
        df.columns = df.iloc[header_row]  # 헤더로 설정
        df = df[header_row + 1:].reset_index(drop=True)  # 데이터 시작 부분부터 다시 구성

        # 데이터가 가장 많이 채워진 열 찾기 (메인 데이터 열)
        main_data_col = df.notnull().sum().idxmax()

        # 빈 셀 채우기 (성명, 학년, 번호 등)
        df = df.fillna(method="ffill")

        # 불필요한 열 제거 (메인 데이터 열 기준으로 관련 데이터 유지)
        df = df.loc[:, :main_data_col]  # 메인 데이터 열까지만 유지

        return df

    except Exception as e:
        st.error(f"데이터 전처리 중 오류 발생: {e}")
        return pd.DataFrame()

# Streamlit UI
st.title("엑셀 데이터 전처리 및 정리")

# 엑셀 파일 업로드
uploaded_file = st.file_uploader("엑셀 파일 업로드", type=["xlsx"])

if uploaded_file:
    st.write("업로드된 파일을 처리 중입니다...")
    processed_data = preprocess_excel(uploaded_file)

    if not processed_data.empty:
        st.write("### 전처리된 데이터")
        st.dataframe(processed_data, use_container_width=True)
        st.success("데이터 전처리가 완료되었습니다!")
    else:
        st.warning("데이터 전처리 결과가 비어 있습니다.")

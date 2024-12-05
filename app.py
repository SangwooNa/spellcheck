import streamlit as st
import pandas as pd
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
        corrected_text = response.text.split('"checked":')[1].split('"')[1]
        return corrected_text
    return text  # API 호출 실패 시 원본 반환

def preprocess_excel(file):
    """엑셀 데이터 전처리 및 진행률 표시"""
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active
        if ws is None:
            raise ValueError("엑셀 파일에 활성화된 워크시트가 없습니다.")

        # 초기화
        data = []
        total_rows = len(list(ws.rows))
        progress = 0
        progress_bar = st.progress(0)
        progress_text = st.empty()

        # 엑셀 데이터를 리스트로 변환
        for i, row in enumerate(ws.iter_rows(values_only=True), 1):
            data.append([cell if cell is not None else "" for cell in row])

            # 진행률 업데이트
            progress = int((i / total_rows) * 100)
            progress_bar.progress(progress / 100)
            progress_text.text(f"전처리 진행률: {progress}%")

        # 데이터프레임 생성
        df = pd.DataFrame(data)

        # 헤더 추출: "번 호", "성 명", "학 년"이 포함된 행
        header_row = df.apply(
            lambda row: row.astype(str).str.contains("번 호|성 명|학 년", na=False).any(), axis=1
        ).idxmax()
        df.columns = df.iloc[header_row]  # 헤더로 설정
        df = df[header_row + 1:].reset_index(drop=True)  # 데이터만 남기기

        # 빈 셀 처리 및 앞의 데이터로 채우기
        df.fillna(method="ffill", inplace=True)

        # 중복된 열 이름 해결
        df.columns = pd.io.parsers.ParserBase({'names': df.columns})._maybe_dedup_names(df.columns)

        # 진행 완료 메시지
        progress_text.text("전처리 완료!")
        progress_bar.empty()

        return df

    except Exception as e:
        st.error(f"데이터 전처리 중 오류 발생: {e}")
        return pd.DataFrame()

def highlight_spell_errors(df):
    """맞춤법 검사 및 강조 스타일 추가"""
    corrected_df = df.copy()
    error_map = {}

    for col in df.columns:
        for idx, value in df[col].items():
            if isinstance(value, str):  # 텍스트 셀만 검사
                corrected_value = check_spelling(value)
                if corrected_value != value:  # 수정 필요
                    error_map[(idx, col)] = corrected_value
                    corrected_df.at[idx, col] = f"<span style='background-color: yellow;'>{value}</span>"

    return corrected_df, error_map

def render_interactive_table(df, error_map):
    """인터랙티브 HTML 테이블 생성"""
    def create_cell_html(row, col, value):
        corrected_value = error_map.get((row, col), None)
        if corrected_value:
            return f"<span onclick='showCorrection(\"{value}\", \"{corrected_value}\")'>{value}</span>"
        return value

    html_table = "<table border='1' style='border-collapse: collapse; width: 100%;'>"
    # 테이블 헤더 생성
    html_table += "<thead><tr>"
    for col in df.columns:
        html_table += f"<th>{col}</th>"
    html_table += "</tr></thead>"

    # 테이블 데이터 생성
    html_table += "<tbody>"
    for idx, row in df.iterrows():
        html_table += "<tr>"
        for col in df.columns:
            value = row[col]
            html_table += f"<td>{create_cell_html(idx, col, value)}</td>"
        html_table += "</tr>"
    html_table += "</tbody></table>"

    # JavaScript 팝업 추가
    html_table += """
    <script>
    function showCorrection(original, corrected) {
        alert("원본: " + original + "\\n수정안: " + corrected);
    }
    </script>
    """

    return html_table

# Streamlit 앱 UI
st.title("엑셀 맞춤법 검사기")
st.write("업로드한 엑셀 데이터를 브라우저에서 처리하고, 맞춤법 검사를 수행합니다.")

# 엑셀 파일 업로드
uploaded_file = st.file_uploader("엑셀 파일 업로드", type=["xlsx"])

if uploaded_file:
    # 엑셀 데이터 전처리
    st.write("업로드된 파일을 처리 중입니다...")
    processed_data = preprocess_excel(uploaded_file)

    if not processed_data.empty:
        st.write("### 전처리된 데이터")
        st.dataframe(processed_data, use_container_width=True)  # 전처리 결과 테이블 표시

        # 맞춤법 검사 수행
        if st.button("맞춤법 검사 시작"):
            st.write("맞춤법 검사를 수행 중입니다...")
            corrected_data, error_map = highlight_spell_errors(processed_data)

            # 인터랙티브 테이블 생성
            html_table = render_interactive_table(processed_data, error_map)
            st.markdown(html_table, unsafe_allow_html=True)
    else:
        st.warning("전처리된 데이터가 비어 있습니다.")

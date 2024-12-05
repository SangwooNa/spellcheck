import os
import openpyxl
import requests

# 네이버 맞춤법 검사 API 호출
def spell_check(text):
    url = "https://m.search.naver.com/p/csearch/ocontent/util/SpellerProxy"
    params = {"q": text, "where": "nexearch"}
    response = requests.get(url, params=params)
    if response.status_code == 200:
        data = response.text.replace("window.__jindo2_callback._speller(", "").rstrip(");")
        return data
    return None

# 엑셀 파일 로드
file_name = os.getenv("FILE_NAME")
wb = openpyxl.load_workbook(file_name)
ws = wb.active

# 맞춤법 검사 수행
for row in ws.iter_rows(values_only=True):
    for cell in row:
        if cell:
            result = spell_check(cell)
            print(f"Original: {cell}, Corrected: {result}")

# 결과 저장
wb.save("corrected.xlsx")

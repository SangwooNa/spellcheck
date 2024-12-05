from flask import Flask, request, send_file
from hanspell import spell_checker
import openpyxl
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
RESULT_FOLDER = "results"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

def check_spelling(text):
    """맞춤법 검사 함수"""
    result = spell_checker.check(text)
    if result.errors:
        return result.checked
    return text

@app.route('/upload', methods=['POST'])
def upload_file():
    """엑셀 파일 업로드 및 맞춤법 검사"""
    if 'file' not in request.files:
        return {"error": "No file provided"}, 400

    file = request.files['file']
    if file.filename == '':
        return {"error": "No file selected"}, 400

    input_path = os.path.join(UPLOAD_FOLDER, file.filename)
    output_path = os.path.join(RESULT_FOLDER, f"corrected_{file.filename}")
    file.save(input_path)

    # 엑셀 파일 처리
    wb = openpyxl.load_workbook(input_path)
    ws = wb.active

    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                original = cell.value
                corrected = check_spelling(original)
                cell.value = corrected  # 수정된 텍스트로 업데이트

    wb.save(output_path)
    return {"message": "File processed successfully", "output_file": f"corrected_{file.filename}"}

@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    """수정된 엑셀 파일 다운로드"""
    file_path = os.path.join(RESULT_FOLDER, filename)
    if not os.path.exists(file_path):
        return {"error": "File not found"}, 404
    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)

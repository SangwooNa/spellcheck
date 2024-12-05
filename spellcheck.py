import openpyxl
from hanspell import spell_checker

def check_spelling(text):
    result = spell_checker.check(text)
    return result.checked if result.errors else text

def process_excel(file_name):
    wb = openpyxl.load_workbook(file_name)
    ws = wb.active
    output_html = "<table><tr><th>Original</th><th>Corrected</th></tr>"

    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                corrected = check_spelling(cell.value)
                output_html += f"<tr><td>{cell.value}</td><td>{corrected}</td></tr>"

    output_html += "</table>"
    with open("output.html", "w", encoding="utf-8") as f:
        f.write(output_html)

if __name__ == "__main__":
    process_excel("input.xlsx")

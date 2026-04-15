import fitz
import re
from openpyxl import Workbook
import os

def extract_text_from_pdf(pdf_path):
    text_all = ""
    doc = fitz.open(pdf_path)
    for page in doc:
        text_all += page.get_text()
    return text_all

def extract_amount(text):
    match = re.search(r'(合計金額|請求金額|合計).*?([¥￥]?\d{1,3}(?:,\d{3})*)', text)
    if match:
        return match.group(2)
    return "不明"

def write_to_excel(results, output_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["ファイル名", "金額"])

    for row in results:
        ws.append(row)

    wb.save(output_path)

def main():
    folder = "pdfs"
    results = []

    for file in os.listdir(folder):
        if file.endswith(".pdf"):
            path = os.path.join(folder, file)
            text = extract_text_from_pdf(path)
            amount = extract_amount(text)
            results.append((file, amount))

    write_to_excel(results, "result.xlsx")

if __name__ == "__main__":
    main()

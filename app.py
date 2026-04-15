from flask import Flask, render_template, request, send_file
import fitz
import pdfplumber
from openpyxl import Workbook
import re
import os
import unicodedata

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# =========================
# 正規化
# =========================
def normalize_text(text):
    return unicodedata.normalize("NFKC", text)


# =========================
# テキスト抽出
# =========================
def extract_text(pdf_path):
    text = ""
    doc = fitz.open(pdf_path)
    for page in doc:
        text += page.get_text()
    return normalize_text(text)


# =========================
# 行整形
# =========================
def clean_lines(text):
    return [l.strip() for l in text.split("\n") if l.strip()]


# =========================
# 請求日
# =========================
def extract_date(text):
    lines = clean_lines(text)

    for line in lines:
        if "請求日" in line:
            m = re.search(r'\d{1,2}/\d{1,2}/\d{4}', line)
            if m:
                return m.group()

    m = re.search(r'\d{1,2}/\d{1,2}/\d{4}', text)
    return m.group() if m else "不明"


# =========================
# 請求金額
# =========================
def extract_amount(text):
    lines = clean_lines(text)

    for i, line in enumerate(lines):
        if "ご請求金額" in line or "合計" in line:
            m = re.search(r'\d{1,3}(?:,\d{3})*', line)
            if m:
                return m.group()

            for j in range(i+1, min(i+6, len(lines))):
                m = re.search(r'\d{1,3}(?:,\d{3})*', lines[j])
                if m:
                    return m.group()

    nums = re.findall(r'\d{1,3}(?:,\d{3})*', text)
    if nums:
        nums = [int(n.replace(",", "")) for n in nums]
        return f"{max(nums):,}"

    return "不明"


# =========================
# 税率抽出（NEW🔥）
# =========================
def extract_tax_rate(text):
    # 10% or 8%を探す
    m = re.search(r'(10|8)\s*%', text)
    if m:
        return int(m.group(1)) / 100

    # fallback
    return 0.10


# =========================
# 明細抽出
# =========================
def extract_items(pdf_path):
    items = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables()

            for table in tables:
                for row in table:
                    if not row:
                        continue

                    row = [str(c).strip() if c else "" for c in row]

                    if "品目" in "".join(row):
                        continue

                    nums = [c for c in row if re.search(r'\d', c)]

                    if len(nums) >= 3:
                        price = nums[-1]
                        unit_price = nums[-2]
                        qty = nums[-3] if len(nums) >= 3 else ""

                        name = row[0]

                        items.append([
                            name,
                            qty,
                            unit_price,
                            price
                        ])

    return items


# =========================
# 計算ロジック（NEW🔥）
# =========================
def calculate_summary(items, tax_rate):
    subtotal = 0

    for item in items:
        price = item[3].replace(",", "")
        if price.isdigit():
            subtotal += int(price)

    # 端数処理（切り捨て）
    tax = int(subtotal * tax_rate)
    total = subtotal + tax

    return {
        "小計": f"{subtotal:,}",
        "消費税": f"{tax:,} ({int(tax_rate*100)}%)",
        "合計": f"{total:,}"
    }


# =========================
# メイン処理
# =========================
def process_pdf(pdf_path):
    text = extract_text(pdf_path)

    date = extract_date(text)
    amount = extract_amount(text)
    tax_rate = extract_tax_rate(text)

    items = extract_items(pdf_path)
    summary = calculate_summary(items, tax_rate)

    return date, amount, summary, items


# =========================
# Flask
# =========================
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        files = request.files.getlist("pdfs")

        wb = Workbook()
        ws = wb.active

        ws.append([
            "ファイル名",
            "請求日",
            "請求金額",
            "品目",
            "数量",
            "単価",
            "金額",
            "小計",
            "消費税",
            "合計"
        ])

        for file in files:
            if file.filename.endswith(".pdf"):
                path = os.path.join(UPLOAD_FOLDER, file.filename)
                file.save(path)

                date, amount, summary, items = process_pdf(path)

                if items:
                    for item in items:
                        ws.append([
                            file.filename,
                            date,
                            amount,
                            item[0],
                            item[1],
                            item[2],
                            item[3],
                            summary["小計"],
                            summary["消費税"],
                            summary["合計"]
                        ])
                else:
                    ws.append([
                        file.filename,
                        date,
                        amount,
                        "",
                        "",
                        "",
                        "",
                        summary["小計"],
                        summary["消費税"],
                        summary["合計"]
                    ])

        output_path = "result.xlsx"
        wb.save(output_path)

        return send_file(output_path, as_attachment=True)

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
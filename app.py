from flask import Flask, request, send_file, jsonify, render_template
import pdfplumber
from openpyxl import Workbook
from io import BytesIO

app = Flask(__name__)


def process_pdf_with_styles(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Extracted Table"

        headers_seen = set()

        for page in pdf.pages:
            tables = page.extract_tables()
            if not tables:
                continue

            for table in tables:
                for row in table:
                    first_cell = row[0].strip() if row[0] else ""
                    if "Түсу күні/ Дата\nпоступления" in first_cell:
                        if first_cell in headers_seen:
                            continue
                        headers_seen.add(first_cell)
                    sheet.append(row)

        excel_file = BytesIO()
        workbook.save(excel_file)
        excel_file.seek(0)
        return excel_file


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "Файл не выбран"}), 400

    pdf_file = request.files['file']

    try:
        excel_file = process_pdf_with_styles(pdf_file)
        return send_file(
            excel_file,
            as_attachment=True,
            download_name="output.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True)

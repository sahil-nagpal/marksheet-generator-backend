from io import BytesIO
from flask import Flask, request, send_file, jsonify
import pandas as pd
from flask_cors import CORS
app = Flask(__name__)

CORS(app)


@app.route("/")
def hello():
    return "Excel generator API. POST JSON to /generate"

@app.route("/generate", methods=["POST"])
def generate_xlsx():
    try:
        data = request.get_json()

        columns = data.get("columns", [])
        rows = data.get("rows", [])
        sheet_name = data.get("sheet_name", "Sheet1")
        filename = data.get("filename", "marksheet.xlsx")
        school_name = data.get("schoolName", "My School Name")
        class_name = data.get("className", "Class 10A")
        exam_name = data.get("exam_name", "Mid Term 2025")

        if not columns or not rows:
            return jsonify({"error": "Columns and rows are required"}), 400

        output = BytesIO()
        writer = pd.ExcelWriter(output, engine="xlsxwriter")
        workbook = writer.book
        worksheet = workbook.add_worksheet(sheet_name)
        writer.sheets[sheet_name] = worksheet

        # 🎨 Define formats
        header_fmt = workbook.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter',
            'border': 1, 'bg_color': '#8E44AD', 'font_color': 'white'
        })
        sub_header_fmt = workbook.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter',
            'border': 1, 'bg_color': '#5DADE2', 'font_color': 'white'
        })
        total_header_fmt = workbook.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter',
            'border': 1, 'bg_color': '#27AE60', 'font_color': 'white'
        })
        rank_header_fmt = workbook.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter',
            'border': 1, 'bg_color': '#F4D03F', 'font_color': 'black'
        })
        title_fmt = workbook.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter',
            'font_size': 14
        })
        sub_title_fmt = workbook.add_format({
            'bold': True, 'align': 'center', 'valign': 'vcenter',
            'font_size': 12
        })
        cell_fmt = workbook.add_format({'align': 'center', 'border': 1})

        # 🧮 Compute total columns for merging headers
        total_subject_cols = len(columns) - 3  # minus Student, Total, Rank
        total_cols = total_subject_cols + 3    # add back Student, Total, Rank

        # === Row 0: School name ===
        worksheet.merge_range(0, 0, 0, min(total_cols - 1, 14), school_name, title_fmt)

        # === Row 1: Class and Exam name ===
        split_point = 10 if total_cols > 10 else total_cols // 2
        worksheet.merge_range(1, 0, 1, split_point - 1, f"Class: {class_name}", sub_title_fmt)
        worksheet.merge_range(1, split_point, 1, total_cols - 1, f"Exam: {exam_name}", sub_title_fmt)

        # === Start header rows from row 2 ===
        current_row = 2

        # Write Student header
        col = 0
        worksheet.merge_range(current_row, col, current_row + 1, col, "STUDENT NAME", header_fmt)
        col += 1

        # 🧩 Build subject-subpart map
        subparts_map = {}
        for subject, subpart in columns[1:-2]:  # Skip Student, Total, Rank
            if subject not in subparts_map:
                subparts_map[subject] = []
            if subpart:
                subparts_map[subject].append(subpart)

        # Write merged subjects and subparts
        for subject, subparts in subparts_map.items():
            start_col = col
            end_col = col + len(subparts) - 1

            if len(subparts) > 1:
                worksheet.merge_range(current_row, start_col, current_row, end_col, subject, header_fmt)
            else:
                worksheet.write(current_row, start_col, subject, header_fmt)

            for sp in subparts:
                worksheet.write(current_row + 1, col, sp, sub_header_fmt)
                col += 1

        # TOTAL and RANK headers
        worksheet.merge_range(current_row, col, current_row + 1, col, "TOTAL", total_header_fmt)
        col += 1
        worksheet.merge_range(current_row, col, current_row + 1, col, "RANK", rank_header_fmt)

        # Write student data (starts from row + 2)
        for r, row in enumerate(rows, start=current_row + 2):
            for c, val in enumerate(row):
                worksheet.write(r, c, val, cell_fmt)

        # Adjust column widths
        worksheet.set_column(0, 0, 20)  # Student name
        worksheet.set_column(1, col, 12)

        writer.close()
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    app.run(debug=True, port=5000)

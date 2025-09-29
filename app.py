from flask import Flask, render_template, request, redirect, url_for, send_file, jsonify, send_from_directory
import os
import pandas as pd
from bs4 import BeautifulSoup
import uuid
import json

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = os.path.join(os.getcwd(), "uploads")
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)

output_file = r"C:\Users\Vishal\OneDrive\Desktop\extract\output.xlsx"

def get_selection_file(filename):
    return os.path.join(app.config["UPLOAD_FOLDER"], f"{filename}_selections.json")


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        files = request.files.getlist("html_files")
        file_ids = []
        for file in files:
            filename = str(uuid.uuid4()) + ".html"
            filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
            file.save(filepath)
            file_ids.append(filename)
        return redirect(url_for("preview", filename=file_ids[0]))
    return render_template("index.html")


@app.route("/uploads/<filename>")
def uploaded_file(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)


@app.route("/preview")
def preview():
    filename = request.args.get("filename")
    if not filename:
        return "File not specified", 400

    file_url = url_for("uploaded_file", filename=filename)

    # Read existing Excel columns
    try:
        df = pd.read_excel(output_file, engine="openpyxl")
        columns = df.columns.tolist()
    except Exception:
        columns = []

    return render_template("preview.html", file_url=file_url, filename=filename, columns=columns)


@app.route("/save_selection", methods=["POST"])
def save_selection():
    label = request.json.get("label")
    selector = request.json.get("selector")
    filename = request.json.get("filename")
    if not filename:
        return jsonify({"status": "error", "message": "Filename missing"}), 400

    selections_file = get_selection_file(filename)
    if os.path.exists(selections_file):
        try:
            with open(selections_file, "r", encoding="utf-8") as f:
                selections = json.load(f)
        except Exception:
            selections = {}
    else:
        selections = {}

    selections[label] = selector
    with open(selections_file, "w", encoding="utf-8") as f:
        json.dump(selections, f, indent=2)

    return jsonify({"status": "success", "label": label, "selector": selector})

@app.route("/extract", methods=["POST"])
def extract():
    files = request.form.getlist("files")
    pasted_data = request.form.get("pasted_data", "").strip()

    if not pasted_data:
        return "No pasted data provided!", 400

    try:
        pasted_columns = json.loads(pasted_data)  # dict: {label: [values]}
    except Exception:
        pasted_columns = {}

    all_data = []

    # Determine number of rows to insert based on the longest pasted column
    max_rows = max(len(v) for v in pasted_columns.values()) if pasted_columns else 0

    for file_path in files:
        selections_file = get_selection_file(file_path)
        if not os.path.exists(selections_file):
            return f"No selections saved for {file_path}!", 400

        with open(selections_file, "r", encoding="utf-8") as f:
            selections = json.load(f)

        file_path_full = os.path.join(app.config["UPLOAD_FOLDER"], file_path)
        with open(file_path_full, "r", encoding="utf-8") as f:
            soup = BeautifulSoup(f, "html.parser")

        # Add one row per pasted value set
        for i in range(max_rows):
            row = {}
            for label, selector in selections.items():
                element = soup.select_one(selector)
                row[label] = element.get_text(strip=True) if element else ""

            # Add pasted column data for this row
            for label, values in pasted_columns.items():
                row[label] = values[i] if i < len(values) else ""

            all_data.append(row)

    if os.path.exists(output_file):
        try:
            df_existing = pd.read_excel(output_file, engine="openpyxl")
            df_new = pd.DataFrame(all_data)
            df = pd.concat([df_existing, df_new], ignore_index=True)
        except Exception as e:
            return f"Error reading existing Excel file: {e}", 500
    else:
        df = pd.DataFrame(all_data)

    df.to_excel(output_file, index=False, engine="openpyxl")
    return send_file(output_file, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)

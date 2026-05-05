from __future__ import annotations

from pathlib import Path

from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from excel_processor import WorkbookProcessingError, process_uploaded_workbook


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "tmp" / "uploads"
OUTPUT_DIR = BASE_DIR / "tmp" / "outputs"
ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024
app.secret_key = "mb-calculator-local-secret"


def is_allowed_file(filename: str) -> bool:
    return Path(filename).suffix.lower() in ALLOWED_EXTENSIONS


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/process")
def process_file():
    uploaded_file = request.files.get("file")
    if uploaded_file is None or uploaded_file.filename == "":
        flash("Choose an Excel workbook first.")
        return redirect(url_for("index"))

    if not is_allowed_file(uploaded_file.filename):
        flash("Upload a .xlsx or .xlsm file.")
        return redirect(url_for("index"))

    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    safe_name = secure_filename(uploaded_file.filename)
    upload_path = UPLOAD_DIR / safe_name
    uploaded_file.save(upload_path)

    try:
        result = process_uploaded_workbook(upload_path, OUTPUT_DIR)
    except WorkbookProcessingError as exc:
        flash(str(exc))
        return redirect(url_for("index"))
    finally:
        upload_path.unlink(missing_ok=True)

    download_name = f"{Path(safe_name).stem}_output.xlsx"
    return send_file(
        result.output_path,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    app.run(debug=True)

"""
app.py - Flask backend for the Transfer Matrix Reshaper.

Design goal: nothing this app receives or produces is ever written to
permanent storage. Every request that touches a file does its work inside
a tempfile.TemporaryDirectory(), which is deleted the moment the request
finishes (success or failure) - so uploaded workbooks and generated
matrices never pile up on disk between requests.

Two endpoints:
  POST /api/inspect  - read a workbook's sheets/blocks so the frontend can
                        show real sheet names, block titles, STA ranges and
                        outbound codes to choose from. Reads straight out
                        of the uploaded file's memory stream; never touches
                        disk at all.
  POST /api/process  - actually build the modified matrix (or, in "all"
                        mode, every block on every sheet, zipped) and
                        stream the result back as a download.
"""
import io
import os
import tempfile
import zipfile

from flask import Flask, request, jsonify, send_file, send_from_directory
from werkzeug.utils import secure_filename
import openpyxl

from build_modified_matrix import (
    find_matrix_blocks,
    _read_source,
    generate_modified_matrix,
    generate_all,
)

app = Flask(__name__, static_folder="static", static_url_path="")

# Keep uploads bounded - this is a small internal tool, not a file host.
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024  # 25 MB


def _fmt_time(v):
    """Best-effort HH:MM string for whatever openpyxl handed back."""
    if v is None:
        return None
    if hasattr(v, "strftime"):
        try:
            return v.strftime("%H:%M")
        except (ValueError, TypeError):
            return None
    return None


def _inspect(file_stream):
    """Read sheet/block structure straight from an in-memory file object.
    Never writes anything to disk."""
    wb = openpyxl.load_workbook(file_stream, data_only=True)
    sheets = []
    for name in wb.sheetnames:
        ws = wb[name]
        title_rows = find_matrix_blocks(ws)
        if not title_rows:
            continue
        blocks = []
        for tr in title_rows:
            title, incoming, outbound = _read_source(ws, title_row=tr)
            stas = [f["sta"] for f in incoming if f["sta"] is not None]
            codes = sorted({o["code"] for o in outbound if o["code"]})
            blocks.append({
                "title_row": tr,
                "title": title or f"Block at row {tr}",
                "incoming_count": len(incoming),
                "outbound_codes": codes,
                "sta_min": _fmt_time(min(stas)) if stas else None,
                "sta_max": _fmt_time(max(stas)) if stas else None,
            })
        sheets.append({"name": name, "blocks": blocks})
    return sheets


@app.route("/")
def index():
    return send_from_directory(app.static_folder, "index.html")


@app.route("/api/inspect", methods=["POST"])
def api_inspect():
    file = request.files.get("file")
    if file is None or file.filename == "":
        return jsonify({"error": "No file was uploaded."}), 400
    if not file.filename.lower().endswith((".xlsx", ".xlsm")):
        return jsonify({"error": "Please upload an .xlsx workbook."}), 400

    try:
        sheets = _inspect(file.stream)
    except Exception as exc:  # noqa: BLE001 - surface a readable message
        return jsonify({"error": f"Could not read that workbook: {exc}"}), 400

    if not sheets:
        return jsonify({
            "error": "No 'TRANSFER MATRIX' block was found on any sheet in this file."
        }), 400

    return jsonify({"sheets": sheets})


@app.route("/api/process", methods=["POST"])
def api_process():
    file = request.files.get("file")
    if file is None or file.filename == "":
        return jsonify({"error": "No file was uploaded."}), 400
    if not file.filename.lower().endswith((".xlsx", ".xlsm")):
        return jsonify({"error": "Please upload an .xlsx workbook."}), 400

    mode = request.form.get("mode", "single")
    sheet = request.form.get("sheet") or None
    sort_by = request.form.get("sort_by", "departure")
    min_sta = request.form.get("min_sta") or None
    max_sta = request.form.get("max_sta") or None
    outbound_raw = request.form.get("outbound") or ""
    outbound_codes = [c.strip() for c in outbound_raw.split(",") if c.strip()] or None

    try:
        title_row = int(request.form.get("title_row", 1))
    except ValueError:
        title_row = 1

    if sort_by not in ("departure", "count"):
        return jsonify({"error": "sort_by must be 'departure' or 'count'."}), 400

    # Every uploaded/generated file for this request lives in here, and it
    # is removed the instant we leave this block - nothing persists.
    with tempfile.TemporaryDirectory(prefix="matrix_") as tmpdir:
        safe_name = secure_filename(file.filename) or "source.xlsx"
        src_path = os.path.join(tmpdir, safe_name)
        file.save(src_path)

        try:
            if mode == "all":
                out_dir = os.path.join(tmpdir, "out")
                os.makedirs(out_dir, exist_ok=True)
                paths = generate_all(
                    src_path, out_dir,
                    min_sta=min_sta, max_sta=max_sta,
                    outbound_codes=outbound_codes, sort_by=sort_by,
                )
                if not paths:
                    return jsonify({
                        "error": "No 'TRANSFER MATRIX' block was found on any sheet in this file."
                    }), 400

                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for p in paths:
                        zf.write(p, arcname=os.path.basename(p))
                zip_buffer.seek(0)
                return send_file(
                    zip_buffer,
                    as_attachment=True,
                    download_name="modified_matrices.zip",
                    mimetype="application/zip",
                )

            out_path = os.path.join(tmpdir, "modified_matrix.xlsx")
            generate_modified_matrix(
                src_path, sheet, out_path,
                min_sta=min_sta, max_sta=max_sta,
                outbound_codes=outbound_codes, sort_by=sort_by,
                title_row=title_row,
            )
            with open(out_path, "rb") as fh:
                buffer = io.BytesIO(fh.read())
            buffer.seek(0)
            download_name = "modified_matrix.xlsx"
            return send_file(
                buffer,
                as_attachment=True,
                download_name=download_name,
                mimetype=(
                    "application/vnd.openxmlformats-officedocument"
                    ".spreadsheetml.sheet"
                ),
            )
        except KeyError:
            return jsonify({"error": f"Sheet '{sheet}' was not found in the workbook."}), 400
        except Exception as exc:  # noqa: BLE001 - surface a readable message
            return jsonify({"error": f"Could not process that workbook: {exc}"}), 400


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)

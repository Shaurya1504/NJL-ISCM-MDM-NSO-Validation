"""
Flask backend for PR Lookup web app.

Routes:
  GET  /           → serve frontend
  POST /process    → accept PR xlsx upload → run report → stream xlsx back
                     Output sheet: PR cols (10) + SMC cols (15) + Duplicate Flag
                     SMC cols 13-15: Item Stage | Item Model Group | Ledger Dimension
  GET  /health     → uptime check for Render.com

Accepted upload formats (two prescribed types):
  Format A — NS0061 style (10-col PR sheet):
    Purchase Requsition Number, Item Id, CW Qty, Store, Priority,
    Site, Warehouse, Location, Set Id, Merchandise Expected Date
  Format B — PR Validation input style (2-col):
    Standalone Pcs, Set Code
"""

import os, io, traceback, uuid, time
from flask import Flask, request, send_file, jsonify

app = Flask(__name__, static_folder=None)

# ── config ────────────────────────────────────────────────────────────────────
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024   # 20 MB upload cap

AZURE_SMC_URL = os.environ.get(
    "AZURE_SMC_URL",
    "https://njlprodimages.blob.core.windows.net/protopisfolder/Set_Master_Combined.parquet"
)

# ── prescribed formats ────────────────────────────────────────────────────────
FORMAT_A_COLS = {
    'Purchase Requsition Number', 'Item Id', 'CW Qty', 'Store',
    'Priority', 'Site', 'Warehouse', 'Location', 'Set Id',
    'Merchandise Expected Date',
}
FORMAT_B_COLS = {'Standalone Pcs', 'Set Code'}

UPLOAD_FOLDER = "/tmp/pr_lookup"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def detect_format(header):
    """
    Returns 'A', 'B', or None.
    Format A requires all 10 NS0061-style columns.
    Format B requires exactly the 2 PR-Validation-input columns.
    """
    col_set = set(header)
    if FORMAT_A_COLS.issubset(col_set):
        return 'A'
    if FORMAT_B_COLS.issubset(col_set):
        return 'B'
    return None


# ── routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_file(os.path.join(os.path.dirname(os.path.abspath(__file__)), "index.html"))


@app.route("/health")
def health():
    return jsonify({"status": "ok", "time": time.time()})


@app.route("/process", methods=["POST"])
def process():
    # ── validate upload ───────────────────────────────────────────────────────
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded."}), 400

    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "Empty filename."}), 400
    if not f.filename.lower().endswith(".xlsx"):
        return jsonify({"error": "INVALID_FORMAT"}), 422

    # ── save upload to temp ───────────────────────────────────────────────────
    uid      = uuid.uuid4().hex[:8]
    pr_path  = os.path.join(UPLOAD_FOLDER, f"pr_{uid}.xlsx")
    out_path = os.path.join(UPLOAD_FOLDER, f"pr_{uid}_Lookup.xlsx")

    try:
        f.save(pr_path)

        # ── detect format ─────────────────────────────────────────────────────
        from python_calamine import CalamineWorkbook
        wb    = CalamineWorkbook.from_path(pr_path)
        sheet = wb.get_sheet_by_index(0)
        rows  = sheet.to_python(skip_empty_area=False)
        if not rows:
            return jsonify({"error": "INVALID_FORMAT"}), 422

        header = [str(h).strip() for h in rows[0]]
        fmt    = detect_format(header)
        if fmt is None:
            return jsonify({"error": "INVALID_FORMAT"}), 422

        # ── run report ────────────────────────────────────────────────────────
        from report import run_lookup
        run_lookup(pr_path, AZURE_SMC_URL, out_path)

        base    = os.path.splitext(f.filename)[0]
        dl_name = f"{base}_Lookup.xlsx"

        return send_file(
            out_path,
            as_attachment=True,
            download_name=dl_name,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except ValueError as e:
        return jsonify({"error": str(e)}), 422
    except Exception:
        return jsonify({"error": "Processing failed.\n\n" + traceback.format_exc()}), 500
    finally:
        if os.path.exists(pr_path):
            os.remove(pr_path)


# ── entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
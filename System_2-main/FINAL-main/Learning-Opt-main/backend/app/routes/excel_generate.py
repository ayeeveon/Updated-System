# backend/app/routes/excel_generate.py
import os
import re
from flask import Blueprint, request, jsonify, send_file, current_app

from app.services.excel_filler import ExcelTemplateFiller

# Optional: Import shared history tracker
try:
    from run import recent_downloads
except ImportError:     
    recent_downloads = []

excel_bp = Blueprint("excel_bp",  __name__, url_prefix="/api")

DEFAULT_TEMPLATE_PATH = os.path.normpath(
    os.path.join(os.path.dirname(__file__), "..", "static", "excel", "template.xlsx")
)
DEFAULT_MAPPING = {}

@excel_bp.route("/delete_excel", methods=["DELETE"])
def delete_excel_file():
    filename = request.args.get("filename")
    if not filename:
        return jsonify({"error": "Filename is required"}), 400

    # Secure path to /static/generated/
    generated_dir = os.path.join("static", "generated")
    file_path = os.path.join(generated_dir, filename)

    if not os.path.exists(file_path):
        return jsonify({"error": "File not found"}), 404

    try:
        os.remove(file_path)

        # Optionally remove from recent_downloads
        global recent_downloads
        recent_downloads = [
            item for item in recent_downloads if item.get("filename") != filename
        ]

        return jsonify({"message": "File deleted successfully"}), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@excel_bp.route("/generate", methods=["POST", "OPTIONS"])
def generate_excel():
    if request.method == "OPTIONS":
        return ("", 204)

    def err(msg, status=400):
        print(f"[ERROR] {msg}")
        return jsonify({"error": msg}), status

    if "file" not in request.files:
        return err("No file part 'file'")

    f = request.files["file"]
    original_name = f.filename or "uploaded"
    if not original_name.lower().endswith((".xlsx", ".xlsm")):
        return err("Please upload an .xlsx or .xlsm file")

    mapping_json = request.form.get("mapping")

    # Sanitize the filename (remove spaces and symbols)
    base_name = os.path.splitext(original_name)[0]
    safe_base_name = re.sub(r'[^a-zA-Z0-9_-]', '_', base_name)
    
    out_name = f"{safe_base_name}_.xlsx"

    # Load template and generate output
    template_path = getattr(current_app, "EXCEL_TEMPLATE_PATH", DEFAULT_TEMPLATE_PATH)
    print(f"[INFO] Using template: {template_path}  (exists={os.path.exists(template_path)})")

    try:
        filler = ExcelTemplateFiller(template_path, default_mapping=DEFAULT_MAPPING)
        out_io, _ = filler.generate_from_filestorage(f, mapping_json)

        # Save to /static/generated/
        output_dir = os.path.join("static", "generated")
        os.makedirs(output_dir, exist_ok=True)
        output_path = os.path.join(output_dir, out_name)
        with open(output_path, "wb") as out_file:
            out_file.write(out_io.getbuffer())

        # Track in download history
        recent_downloads.insert(0, {
            "type": "tesda",
            "filename": out_name,
            
            "url": f"/static/generated/{out_name}"
        })

    except Exception as e:
        return err(str(e), status=500)

    return send_file(
        output_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=out_name,
    )
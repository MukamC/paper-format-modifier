"""
server.py — DocxAnnotator HTTP service
=======================================
Exposes a single endpoint:

    POST /annotate
      multipart/form-data:
        file   : the original .docx
        errors : JSON string (array of error objects from the frontend checker)
      Returns: annotated .docx as an attachment

    GET /health
      Returns: {"status": "ok"}

Start:
    python server.py

The server listens on http://127.0.0.1:5001 by default.
"""

import io
import json
import sys

from flask import Flask, jsonify, request, send_file
from flask_cors import CORS

from annotator import DocxAnnotator

app = Flask(__name__)
CORS(app)  # Allow requests from file:// and any localhost origin

_annotator = DocxAnnotator()


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "version": "1.0"})


@app.route("/annotate", methods=["POST"])
def annotate():
    # --- validate file --------------------------------------------------
    if "file" not in request.files:
        return jsonify({"error": "请求中未包含 file 字段"}), 400

    upload = request.files["file"]
    if not upload.filename.lower().endswith(".docx"):
        return jsonify({"error": "仅支持 .docx 格式"}), 400

    # --- validate errors ------------------------------------------------
    try:
        errors_json = request.form.get("errors", "[]")
        errors_list = json.loads(errors_json)
        if not isinstance(errors_list, list):
            raise ValueError("errors 必须是 JSON 数组")
    except (json.JSONDecodeError, ValueError) as exc:
        return jsonify({"error": f"errors 格式有误：{exc}"}), 400

    # --- annotate -------------------------------------------------------
    try:
        input_bytes   = upload.read()
        result_bytes  = _annotator.annotate_document(input_bytes, errors_list)
    except Exception as exc:
        return jsonify({"error": f"标注失败：{exc}"}), 500

    base_name       = upload.filename.rsplit(".", 1)[0]
    output_filename = f"标注版_{base_name}.docx"

    return send_file(
        io.BytesIO(result_bytes),
        mimetype=(
            "application/vnd.openxmlformats-officedocument"
            ".wordprocessingml.document"
        ),
        as_attachment=True,
        download_name=output_filename,
    )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 5001
    print(f"DocxAnnotator 服务已启动 → http://127.0.0.1:{port}")
    print("按 Ctrl+C 停止服务")
    app.run(host="127.0.0.1", port=port, debug=False)

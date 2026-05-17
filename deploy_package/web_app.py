#!/usr/bin/env python3
"""Web 界面入口：Docker 部署版，基于现有程序实现网页上传和输出下载。"""

import os
import sys
import uuid
import shutil
import json
import time
import threading
import traceback
from pathlib import Path
from flask import Flask, request, render_template, send_file, jsonify, Response

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR / "src"))

try:
    from batch_processor import BatchProcessor
except Exception as import_err:
    print(f"导入 batch_processor 失败: {import_err}")
    traceback.print_exc()
    BatchProcessor = None

ALLOWED_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".pptx", ".docx"}
UPLOAD_ROOT = SCRIPT_DIR / "uploads"
OUTPUT_ROOT = SCRIPT_DIR / "web_output"
UPLOAD_ROOT.mkdir(exist_ok=True)
OUTPUT_ROOT.mkdir(exist_ok=True)

app = Flask(__name__, template_folder="templates")
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "pptx2jpg-secret-key")
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024

processing_jobs = {}
processing_lock = threading.Lock()


def allowed_file(filename: str) -> bool:
    suffix = Path(filename).suffix.lower()
    return suffix in ALLOWED_EXTENSIONS


def _guess_suffix(filename: str) -> str:
    """从文件名中尽力提取扩展名，防止 secure_filename 丢失扩展名"""
    path = Path(filename)
    if path.suffix:
        return path.suffix.lower()
    lower = filename.lower()
    if '.xlsx' in lower:
        return '.xlsx'
    if '.xls' in lower:
        return '.xls'
    if '.pptx' in lower:
        return '.pptx'
    if '.docx' in lower:
        return '.docx'
    return '.xlsx'


def safe_save_file(uploaded_file, upload_dir: Path, prefix: str) -> Path:
    """安全保存上传文件，确保扩展名一定存在"""
    original_suffix = _guess_suffix(uploaded_file.filename)
    safe_name = f"{prefix}{original_suffix}"
    dest = upload_dir / safe_name
    uploaded_file.save(dest)
    print(f"文件已保存: {dest} (原始: {uploaded_file.filename})")
    return dest


EXCEL_MAGIC_BYTES = {
    b'PK\x03\x04': '.xlsx',
    b'\xd0\xcf\x11\xe0': '.xls',
}


def validate_excel_format(filepath: str) -> str:
    with open(filepath, 'rb') as f:
        header = f.read(4)
    for magic, fmt in EXCEL_MAGIC_BYTES.items():
        if header[:len(magic)] == magic:
            return fmt
    raise ValueError(
        f"上传的文件不是有效的 Excel 格式。"
        f"支持格式: .xlsx (Office 2007+) 和 .xls (Office 97-2003)。"
        f"请勿将 CSV/HTML 等文件改后缀上传。"
    )


def _run_processing(run_id, excel_path, pptx_path, docx_path, output_docs_dir, output_images_dir, generate_png):
    try:
        processor = BatchProcessor(
            str(excel_path),
            str(pptx_path) if pptx_path else None,
            str(docx_path) if docx_path else None,
            str(output_docs_dir),
            str(output_images_dir),
        )
        processor.read_data()

        def on_progress(current, total, name):
            with processing_lock:
                processing_jobs[run_id] = {
                    "status": "processing",
                    "current": current,
                    "total": total,
                    "name": name,
                }

        processor.process(generate_png=generate_png, progress_callback=on_progress)

        zip_base = Path(output_docs_dir).parent / "invite_result"
        zip_path = shutil.make_archive(str(zip_base), "zip", root_dir=str(Path(output_docs_dir).parent), base_dir=".")

        with processing_lock:
            processing_jobs[run_id] = {
                "status": "done",
                "zip_path": zip_path,
                "current": processing_jobs.get(run_id, {}).get("total", 0),
                "total": processing_jobs.get(run_id, {}).get("total", 0),
                "name": "完成",
            }
    except Exception as exc:
        traceback.print_exc()
        with processing_lock:
            processing_jobs[run_id] = {
                "status": "error",
                "message": str(exc),
            }


@app.route("/", methods=["GET", "POST"])
def index():
    message = None
    if request.method == "POST":
        excel_file = request.files.get("excel_file")
        pptx_file = request.files.get("pptx_template")
        docx_file = request.files.get("docx_template")
        generate_png = request.form.get("generate_png") == "on"

        if not excel_file or not allowed_file(excel_file.filename):
            message = "请上传有效的 Excel 文件。"
        else:
            pptx_uploaded = pptx_file and allowed_file(pptx_file.filename)
            docx_uploaded = docx_file and allowed_file(docx_file.filename)
            if not pptx_uploaded and not docx_uploaded:
                message = "请至少上传一个 PPTX 或 DOCX 模板文件。"
            else:
                run_id = uuid.uuid4().hex
                run_dir = OUTPUT_ROOT / run_id
                upload_dir = run_dir / "uploads"
                output_docs_dir = run_dir / "output_documents"
                output_images_dir = run_dir / "output_images"
                run_dir.mkdir(parents=True, exist_ok=True)
                upload_dir.mkdir(parents=True, exist_ok=True)
                output_docs_dir.mkdir(parents=True, exist_ok=True)
                output_images_dir.mkdir(parents=True, exist_ok=True)

                excel_path = safe_save_file(excel_file, upload_dir, "data")
                pptx_path = safe_save_file(pptx_file, upload_dir, "template_pptx") if pptx_uploaded else None
                docx_path = safe_save_file(docx_file, upload_dir, "template_docx") if docx_uploaded else None

                try:
                    validate_excel_format(str(excel_path))
                except ValueError as ve:
                    shutil.rmtree(run_dir, ignore_errors=True)
                    message = str(ve)
                    return render_template("index.html", message=message)

                with processing_lock:
                    processing_jobs[run_id] = {"status": "starting", "current": 0, "total": 0, "name": "准备中..."}

                threading.Thread(
                    target=_run_processing,
                    args=(run_id, excel_path, pptx_path, docx_path, output_docs_dir, output_images_dir, generate_png),
                    daemon=True,
                ).start()

                return jsonify({"run_id": run_id})

    return render_template("index.html", message=message)


@app.route("/progress/<run_id>")
def progress(run_id):
    def generate():
        last_status = None
        while True:
            with processing_lock:
                job = processing_jobs.get(run_id, {})
            if job:
                data = json.dumps(job)
                yield f"data: {data}\n\n"
                status = job.get("status")
                if status in ("done", "error"):
                    break
                last_status = status
            time.sleep(0.5)
    return Response(generate(), mimetype="text/event-stream")


@app.route("/download/<run_id>")
def download(run_id):
    with processing_lock:
        job = processing_jobs.get(run_id, {})
    if job.get("status") == "done" and job.get("zip_path"):
        zip_path = job["zip_path"]
        if os.path.exists(zip_path):
            return send_file(
                zip_path,
                mimetype="application/zip",
                as_attachment=True,
                download_name="pptx2jpg_result.zip",
            )
    return jsonify({"error": "文件尚未生成完毕"}), 404


@app.route("/health")
def health_check():
    return "ok", 200


@app.errorhandler(500)
def internal_error(e):
    original = getattr(e, "original_exception", e)
    return f"500 Internal Server Error\n\n{original}\n\n{traceback.format_exc()}", 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
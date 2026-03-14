import datetime as dt
import uuid
from pathlib import Path

from flask import Flask, render_template, request, send_from_directory
from werkzeug.utils import secure_filename

from agent_service import optimize_resume_docx, save_optimized_resume_docx, save_word_units_snapshot

ALLOWED_RESUME_EXT = {"docx", "doc"}
ALLOWED_JOB_EXT = {"txt", "md", "png", "jpg", "jpeg", "webp", "bmp"}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "output"
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def _allowed(filename: str, allowed_ext: set[str]) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in allowed_ext


def _save_upload(file_storage, target_dir: Path) -> Path:
    original_name = (file_storage.filename or "file").strip()
    original_suffix = Path(original_name).suffix.lower()

    safe_stem = secure_filename(Path(original_name).stem) or "file"
    safe_suffix = original_suffix if original_suffix else Path(secure_filename(original_name)).suffix.lower()
    unique_name = f"{uuid.uuid4().hex}_{safe_stem}{safe_suffix}"

    target_path = target_dir / unique_name
    file_storage.save(target_path)
    return target_path


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/optimize")
def optimize():
    resume_file = request.files.get("resume")
    job_file = request.files.get("job_file")
    job_text = (request.form.get("job_text") or "").strip()

    if not resume_file or not resume_file.filename:
        return render_template("index.html", error="请上传 Word 简历（.docx 或 .doc）。")

    if not _allowed(resume_file.filename, ALLOWED_RESUME_EXT):
        return render_template("index.html", error="简历仅支持 .docx 或 .doc 格式。")

    has_job_file = bool(job_file and job_file.filename)
    has_job_text = bool(job_text)

    if has_job_file == has_job_text:
        return render_template("index.html", error="请在岗位文件和岗位文字中二选一。")

    if has_job_file and not _allowed(job_file.filename, ALLOWED_JOB_EXT):
        return render_template(
            "index.html",
            error="岗位文件仅支持 txt/md/png/jpg/jpeg/webp/bmp。",
        )

    try:
        resume_path = _save_upload(resume_file, UPLOAD_DIR)
        job_path = _save_upload(job_file, UPLOAD_DIR) if has_job_file else None

        optimized = optimize_resume_docx(
            resume_path=str(resume_path),
            job_path=str(job_path) if job_path else None,
            job_text=job_text if has_job_text else None,
        )

        ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
        out_docx_name = f"optimized_resume_{ts}.docx"
        out_units_name = f"word_units_{ts}.json"
        out_docx_path = OUTPUT_DIR / out_docx_name
        out_units_path = OUTPUT_DIR / out_units_name

        save_word_units_snapshot(optimized["word_units"], str(out_units_path))
        save_optimized_resume_docx(
            optimized_units=optimized["optimized_units"],
            source_resume_path=str(optimized.get("resolved_resume_path", resume_path)),
            output_docx_path=str(out_docx_path),
        )

        return render_template(
            "index.html",
            download_docx_name=out_docx_name,
            converted_from_doc=optimized.get("converted_from_doc", False),
            parse_report=optimized.get("parse_report", {}),
            verification=optimized.get("verification", {}),
            success="简历优化完成，已生成保持原样式的 Word 简历。",
        )
    except Exception as exc:
        return render_template("index.html", error=f"处理失败: {exc}")


@app.get("/download/<path:filename>")
def download(filename: str):
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True)


@app.get("/health")
def health():
    return {"status": "ok"}


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=True)

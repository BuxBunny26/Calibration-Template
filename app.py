"""
WearCheck ARC — Calibration Report Web API
============================================
Flask API for generating vibration analyzer calibration reports.
Integrates with Supabase for file storage and job history.

Usage:
    python app.py              # Development (localhost:5000)
    gunicorn app:app           # Production (Render)
"""

import os
import shutil
import tempfile
import uuid
import json
import re
from datetime import datetime, timezone

from flask import Flask, request, jsonify
from flask_cors import CORS
from supabase import create_client

from generate_full_report import (
    read_vib_analyzer,
    merge_info,
    generate_cover_page,
    merge_pdfs,
)

# ── App Setup ───────────────────────────────────────────────────────────

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB

CORS(app, origins=[
    "http://localhost:3000",
    r"https://.*\.netlify\.app",
], supports_credentials=False)

ALLOWED_EXTENSIONS = {"xlsx", "pdf"}

# ── Supabase Client ─────────────────────────────────────────────────────

SUPABASE_URL = os.environ.get("SUPABASE_URL", "https://dljknrumyawpvxdvjazn.supabase.co")
SUPABASE_KEY = os.environ.get("SUPABASE_SERVICE_KEY", os.environ.get("SUPABASE_ANON_KEY", ""))
STORAGE_BUCKET = "calibration-files"

sb = None
if SUPABASE_URL and SUPABASE_KEY:
    sb = create_client(SUPABASE_URL, SUPABASE_KEY)


# ── Helpers ──────────────────────────────────────────────────────────────

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def classify_file(filename):
    """Return ('acc'|'vel', 'xlsx'|'pdf') or (None, None)."""
    name_lower = filename.lower()
    ext = name_lower.rsplit(".", 1)[1] if "." in name_lower else ""
    if ext not in ALLOWED_EXTENSIONS:
        return None, None
    if "acc" in name_lower:
        return "acc", ext
    elif "vel" in name_lower:
        return "vel", ext
    return None, ext


def extract_group_key(filename):
    """Extract instrument identifier from filename.

    Examples:
        'B2140 1237749 Acc.xlsx' → 'b2140 1237749'
        'B2140 1238056 Vel.pdf'  → 'b2140 1238056'
    """
    name_lower = filename.lower()
    # Remove extension
    base = name_lower.rsplit(".", 1)[0] if "." in name_lower else name_lower
    # Remove 'acc' or 'vel' keyword and trim
    base = re.sub(r'\b(acc|vel)\b', '', base).strip()
    # Collapse multiple spaces
    base = re.sub(r'\s+', ' ', base)
    return base or "default"


def upload_to_storage(local_path, remote_path):
    """Upload a file to Supabase Storage. Returns public URL."""
    if not sb:
        return None
    with open(local_path, "rb") as f:
        content_type = "application/pdf" if remote_path.endswith(".pdf") else "application/octet-stream"
        sb.storage.from_(STORAGE_BUCKET).upload(
            remote_path, f.read(), {"content-type": content_type, "upsert": "true"}
        )
    res = sb.storage.from_(STORAGE_BUCKET).get_public_url(remote_path)
    return res


def save_job(job_id, status, merged_info, input_files_meta, output_path=None, output_name=None, error_msg=None):
    """Save or update a certificate job record in Supabase."""
    if not sb:
        return
    record = {
        "id": job_id,
        "status": status,
        "equipment_model": merged_info.get("model", ""),
        "serial_number": merged_info.get("serial", ""),
        "manufacturer": merged_info.get("manufacturer", ""),
        "calibration_date": merged_info.get("cal_date", ""),
        "calibration_due": merged_info.get("cal_due", ""),
        "calibration_tech": merged_info.get("cal_tech", ""),
        "customer": merged_info.get("customer", ""),
        "input_files": json.dumps(input_files_meta),
        "output_file_path": output_path,
        "output_file_name": output_name,
        "error_message": error_msg,
    }
    if status == "completed":
        record["completed_at"] = datetime.now(timezone.utc).isoformat()
    sb.table("certificate_jobs").upsert(record).execute()


# ── Routes ───────────────────────────────────────────────────────────────

def _process_group(group_key, saved, job_dir):
    """Process one instrument group and return a result dict."""
    job_id = str(uuid.uuid4())
    input_files_meta = []
    for key, path in saved.items():
        if path:
            input_files_meta.append({
                "name": os.path.basename(path),
                "type": key,
                "size": os.path.getsize(path),
            })

    try:
        has_acc = saved.get("acc_xlsx") or saved.get("acc_pdf")
        has_vel = saved.get("vel_xlsx") or saved.get("vel_pdf")

        if not has_acc and not has_vel:
            return {"job_id": job_id, "group": group_key, "status": "failed",
                    "error": "No ACC or VEL files found for this instrument."}

        if has_acc and not saved.get("acc_xlsx"):
            return {"job_id": job_id, "group": group_key, "status": "failed",
                    "error": "ACC PDF provided but ACC Excel (.xlsx) is missing."}
        if has_acc and not saved.get("acc_pdf"):
            return {"job_id": job_id, "group": group_key, "status": "failed",
                    "error": "ACC Excel provided but ACC PDF is missing."}
        if has_vel and not saved.get("vel_xlsx"):
            return {"job_id": job_id, "group": group_key, "status": "failed",
                    "error": "VEL PDF provided but VEL Excel (.xlsx) is missing."}
        if has_vel and not saved.get("vel_pdf"):
            return {"job_id": job_id, "group": group_key, "status": "failed",
                    "error": "VEL Excel provided but VEL PDF is missing."}

        acc_info = read_vib_analyzer(saved["acc_xlsx"]) if saved.get("acc_xlsx") else None
        vel_info = read_vib_analyzer(saved["vel_xlsx"]) if saved.get("vel_xlsx") else None
        merged = merge_info(acc_info, vel_info)

        cover_path = os.path.join(job_dir, f"_cover_{job_id[:8]}.pdf")
        generate_cover_page(merged, cover_path)

        data_pdfs = []
        if saved.get("acc_pdf"):
            data_pdfs.append(saved["acc_pdf"])
        if saved.get("vel_pdf"):
            data_pdfs.append(saved["vel_pdf"])

        model = merged.get("model", "Unknown").replace(" ", "_").replace("/", "-")
        serial = merged.get("serial", "Unknown").replace(" ", "_")
        cal_date = merged.get("cal_date", "").replace("/", "-") or "undated"
        output_name = f"FullReport_{model}_{serial}_{cal_date}.pdf"
        output_path = os.path.join(job_dir, output_name)

        merge_pdfs(cover_path, data_pdfs, output_path)

        remote_cert_path = f"certificates/{job_id}/{output_name}"
        download_url = upload_to_storage(output_path, remote_cert_path)

        for meta in input_files_meta:
            local = saved.get(meta["type"])
            if local:
                remote = f"uploads/{job_id}/{meta['name']}"
                upload_to_storage(local, remote)
                meta["path"] = remote

        save_job(job_id, "completed", merged, input_files_meta, remote_cert_path, output_name)

        return {
            "job_id": job_id,
            "group": group_key,
            "status": "completed",
            "output_file_name": output_name,
            "download_url": download_url or "",
            "model": merged.get("model", ""),
            "serial": merged.get("serial", ""),
        }

    except Exception as e:
        save_job(job_id, "failed", {}, input_files_meta, error_msg=str(e))
        app.logger.exception("Generation failed for group %s, job %s", group_key, job_id)
        return {"job_id": job_id, "group": group_key, "status": "failed",
                "error": "Certificate generation failed for this instrument."}


@app.route("/api/generate", methods=["POST"])
def generate():
    if "files" not in request.files:
        return jsonify({"error": "No files uploaded"}), 400

    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No files uploaded"}), 400

    batch_dir = os.path.join(tempfile.gettempdir(), f"wearcheck_{uuid.uuid4().hex[:12]}")
    os.makedirs(batch_dir, exist_ok=True)

    try:
        # ── Group files by instrument identifier ──
        groups = {}  # group_key -> {acc_xlsx, vel_xlsx, acc_pdf, vel_pdf}

        for f in files:
            if not f.filename or not allowed_file(f.filename):
                continue

            kind, ext = classify_file(f.filename)
            if kind is None:
                continue

            group_key = extract_group_key(f.filename)
            if group_key not in groups:
                groups[group_key] = {}

            slot = f"{kind}_{ext}"
            path = os.path.join(batch_dir, f"{group_key}_{f.filename}")
            f.save(path)
            groups[group_key][slot] = path

        if not groups:
            return jsonify({"error": "No valid ACC/VEL files found. Files must contain 'Acc' or 'Vel' in their name."}), 400

        # ── Process each group ──
        results = []
        for group_key, saved in groups.items():
            result = _process_group(group_key, saved, batch_dir)
            results.append(result)

        return jsonify({"results": results})

    except Exception as e:
        app.logger.exception("Batch generation failed")
        return jsonify({"error": "Certificate generation failed. Please check your files and try again."}), 500

    finally:
        shutil.rmtree(batch_dir, ignore_errors=True)


@app.route("/health")
def health():
    return jsonify({"status": "ok", "supabase": sb is not None})


if __name__ == "__main__":
    app.run(debug=True, port=5000)

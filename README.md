# WearCheck ARC — Calibration Certificate Generator

Web application for generating vibration analyzer calibration certificates.
Upload ACC and VEL Excel/PDF files, and the system produces a branded multi-page PDF report.

---

## Architecture

```
React Frontend (Netlify)  ──POST /api/generate──▶  Flask API (Render)
        │                                               │
        │  reads history                                │  generates PDF
        ▼                                               ▼
   Supabase DB                                   Supabase Storage
   (certificate_jobs)                            (calibration-files)
```

| Layer | Technology | Location |
|-------|-----------|----------|
| Frontend | React 18 | `frontend/` → Netlify |
| Backend | Flask + gunicorn | `app.py` → Render |
| Database | Supabase PostgreSQL | `certificate_jobs` table |
| Storage | Supabase Storage | `calibration-files` bucket |

---

## Local Development

### Prerequisites

- Python 3.12+
- Node.js 18+

### Backend

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python app.py
```

Runs on `http://localhost:5000`.

### Frontend

```powershell
cd frontend
npm install
npm start
```

Runs on `http://localhost:3000`.

### Run Tests

```powershell
python test_system.py
```

---

## Deployment

### Backend (Render)

1. Push to a Git repo
2. Create a new Web Service on Render pointing to the repo root
3. Set environment variables:
   - `SUPABASE_URL`
   - `SUPABASE_SERVICE_KEY`
4. Render uses `Procfile` and `requirements.txt` automatically

### Frontend (Netlify)

1. Connect repo to Netlify, set base directory to `frontend`
2. Set environment variables:
   - `REACT_APP_SUPABASE_URL`
   - `REACT_APP_SUPABASE_ANON_KEY`
   - `REACT_APP_API_URL` (Render service URL)

### Supabase Setup

1. Run `supabase_calibration_setup.sql` in the SQL Editor
2. Create a **public** storage bucket named `calibration-files`

---

## Project Files

| File | Purpose |
|------|---------|
| `app.py` | Flask API — `/api/generate` endpoint |
| `generate_full_report.py` | Reads ACC/VEL workbooks, builds cover page, merges PDFs |
| `generate_cover_page.py` | PVC workbook cover page generator |
| `generate_certificate.py` | Standalone 2-page PDF certificate from Excel template |
| `create_template.py` | Creates the Excel calibration template |
| `test_system.py` | End-to-end pipeline tests |
| `supabase_calibration_setup.sql` | Database schema for Supabase |
| `Data/` | Sample ACC/VEL test data files |
| `Signatures/` | Technician signature images |

"""
Test script: exercises the full report generation pipeline using the
sample data files in Data/.

Tests:
  1. Python imports for all modules
  2. Read Vib Analyzer from sample ACC/VEL Excel files
  3. Merge ACC + VEL info
  4. Generate cover page PDF
  5. Merge PDFs into full report
  6. Flask API /health endpoint
"""

import os
import sys
import tempfile
import shutil

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(SCRIPT_DIR, "Data")

# Sample files expected in Data/
ACC_XLSX = os.path.join(DATA_DIR, "B2140 1237749 Acc.xlsx")
VEL_XLSX = os.path.join(DATA_DIR, "B2140 1237749 Vel.xlsx")
ACC_PDF  = os.path.join(DATA_DIR, "B2140 1237749 Acc.pdf")
VEL_PDF  = os.path.join(DATA_DIR, "B2140 1237749 Vel.pdf")


def test_imports():
    """Verify all project modules import without errors."""
    print("[1] Testing imports...")
    from generate_full_report import (
        read_vib_analyzer, merge_info, generate_cover_page, merge_pdfs,
    )
    from app import app
    print("    All imports OK")
    return True


def test_read_excel():
    """Read ACC and VEL workbooks and check key fields are populated."""
    print("[2] Reading sample Excel workbooks...")
    from generate_full_report import read_vib_analyzer

    for label, path in [("ACC", ACC_XLSX), ("VEL", VEL_XLSX)]:
        if not os.path.exists(path):
            print(f"    SKIP: {label} file not found: {os.path.basename(path)}")
            return False
        info = read_vib_analyzer(path)
        model = info.get("model", "")
        serial = info.get("serial", "")
        print(f"    {label}: model={model}, serial={serial}")
        assert model, f"{label} workbook returned empty model"
        assert serial, f"{label} workbook returned empty serial"

    print("    Excel reading OK")
    return True


def test_merge_info():
    """Merge ACC + VEL info and check combined fields."""
    print("[3] Merging ACC + VEL info...")
    from generate_full_report import read_vib_analyzer, merge_info

    acc = read_vib_analyzer(ACC_XLSX)
    vel = read_vib_analyzer(VEL_XLSX)
    merged = merge_info(acc, vel)

    assert merged.get("model"), "Merged info missing model"
    assert merged.get("serial"), "Merged info missing serial"
    print(f"    Merged: {merged.get('manufacturer', '')} {merged['model']} S/N {merged['serial']}")
    print("    Merge OK")
    return merged


def test_generate_cover_page(merged):
    """Generate a cover page PDF and verify the file exists."""
    print("[4] Generating cover page...")
    from generate_full_report import generate_cover_page

    out_dir = tempfile.mkdtemp(prefix="wearcheck_test_")
    cover_path = os.path.join(out_dir, "test_cover.pdf")

    generate_cover_page(merged, cover_path)
    assert os.path.exists(cover_path), "Cover page PDF not created"
    size_kb = os.path.getsize(cover_path) / 1024
    print(f"    Cover page: {size_kb:.0f} KB")
    print("    Cover page OK")
    return out_dir, cover_path


def test_merge_pdfs(out_dir, cover_path):
    """Merge cover page + data PDFs into final report."""
    print("[5] Merging PDFs into full report...")
    from generate_full_report import merge_pdfs

    data_pdfs = [p for p in [ACC_PDF, VEL_PDF] if os.path.exists(p)]
    output_path = os.path.join(out_dir, "test_full_report.pdf")

    merge_pdfs(cover_path, data_pdfs, output_path)
    assert os.path.exists(output_path), "Full report PDF not created"
    size_kb = os.path.getsize(output_path) / 1024
    print(f"    Full report: {size_kb:.0f} KB ({1 + len(data_pdfs)} source PDFs merged)")
    print("    PDF merge OK")
    return output_path


def test_flask_health():
    """Confirm Flask app starts and /health returns JSON."""
    print("[6] Testing Flask /health endpoint...")
    from app import app

    with app.test_client() as client:
        resp = client.get("/health")
        assert resp.status_code == 200
        data = resp.get_json()
        assert data["status"] == "ok"
        print(f"    /health => status={data['status']}, supabase={data['supabase']}")
    print("    Flask health OK")
    return True


def run_tests():
    print("=" * 60)
    print("WearCheck ARC — Calibration System Tests")
    print("=" * 60)

    passed = 0
    failed = 0
    out_dir = None

    try:
        # 1. Imports
        test_imports()
        passed += 1

        # 2. Read Excel
        if test_read_excel():
            passed += 1

            # 3. Merge
            merged = test_merge_info()
            passed += 1

            # 4. Cover page
            out_dir, cover_path = test_generate_cover_page(merged)
            passed += 1

            # 5. Full report
            test_merge_pdfs(out_dir, cover_path)
            passed += 1
        else:
            print("    Skipping tests 3-5 (no sample data)")
            failed += 3

        # 6. Flask health
        test_flask_health()
        passed += 1

    except Exception as e:
        print(f"\n    FAILED: {e}")
        failed += 1

    finally:
        if out_dir:
            shutil.rmtree(out_dir, ignore_errors=True)

    print("\n" + "=" * 60)
    total = passed + failed
    print(f"Results: {passed}/{total} passed" + (f", {failed} failed" if failed else ""))
    print("=" * 60)
    return failed == 0


if __name__ == "__main__":
    success = run_tests()
    sys.exit(0 if success else 1)

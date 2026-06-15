"""
ArktechX Visual Regression Check Tool

Converts documents to images and compares them against golden references.
Generates an HTML diff report for human review.

Usage:
    # Register a golden reference
    python3 tests/visual_check.py --register output.docx

    # Compare against golden
    python3 tests/visual_check.py output.docx
    python3 tests/visual_check.py output.html --open   # auto-open HTML report

Requires: LibreOffice (for DOCX→PNG) or weasyprint+Pillow (for HTML→PNG)
"""

import os, sys, json, argparse, hashlib, datetime, base64, subprocess, shutil
from pathlib import Path
from typing import List, Optional, Tuple

_ROOT   = Path(__file__).parent.parent
GOLDEN  = _ROOT / "tests" / "golden"
REPORTS = _ROOT / ".generated" / "reports"
REPORTS.mkdir(parents=True, exist_ok=True)


# ── Image conversion ──────────────────────────────────────────────────────────

def _libreoffice_available() -> bool:
    return shutil.which("libreoffice") is not None or \
           shutil.which("soffice") is not None

def _weasyprint_available() -> bool:
    try:
        import weasyprint
        return True
    except ImportError:
        return False

def docx_to_png(docx_path: str, out_dir: str) -> List[str]:
    """Convert DOCX pages to PNG files using LibreOffice headless."""
    if not _libreoffice_available():
        print("  ⚠  LibreOffice not found — skipping DOCX→PNG conversion.")
        print("     Install: brew install libreoffice  or  apt install libreoffice")
        return []
    cmd = ["libreoffice", "--headless", "--convert-to", "png",
           "--outdir", out_dir, docx_path]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"  ✗ LibreOffice error: {result.stderr}")
        return []
    base = Path(docx_path).stem
    return sorted(str(p) for p in Path(out_dir).glob(f"{base}*.png"))

def html_to_png(html_path: str, out_dir: str) -> List[str]:
    """Convert HTML to PNG via WeasyPrint → PDF → PIL."""
    if not _weasyprint_available():
        print("  ⚠  weasyprint not found — pip3 install weasyprint")
        return []
    try:
        import weasyprint
        from PIL import Image
        pdf_path = os.path.join(out_dir, "_tmp.pdf")
        weasyprint.HTML(filename=html_path).write_pdf(pdf_path)
        # PDF → PNG via pdfplumber or PyMuPDF
        try:
            import fitz   # PyMuPDF
            doc = fitz.open(pdf_path)
            pngs = []
            for i, page in enumerate(doc):
                mat = fitz.Matrix(2, 2)   # 2× scale
                pix = page.get_pixmap(matrix=mat)
                out = os.path.join(out_dir, f"page_{i+1:02d}.png")
                pix.save(out)
                pngs.append(out)
            return pngs
        except ImportError:
            print("  ⚠  PyMuPDF not found — pip3 install pymupdf")
            return [pdf_path]   # return PDF path as fallback
    except Exception as e:
        print(f"  ✗ HTML→PNG error: {e}")
        return []

def to_images(file_path: str, out_dir: str) -> List[str]:
    """Dispatch to correct converter based on file extension."""
    os.makedirs(out_dir, exist_ok=True)
    ext = Path(file_path).suffix.lower()
    if ext == ".docx":
        return docx_to_png(file_path, out_dir)
    elif ext in (".html", ".htm"):
        return html_to_png(file_path, out_dir)
    elif ext == ".pptx":
        if _libreoffice_available():
            return docx_to_png(file_path, out_dir)  # LibreOffice handles PPTX too
        print("  ⚠  LibreOffice needed for PPTX→PNG conversion")
        return []
    else:
        print(f"  ⚠  No image converter for {ext}")
        return []


# ── Perceptual hash comparison ────────────────────────────────────────────────

def _phash(img_path: str, size: int = 16) -> int:
    """Compute a perceptual hash for an image."""
    from PIL import Image
    img  = Image.open(img_path).convert("L").resize((size, size))
    pixels = list(img.getdata())
    avg    = sum(pixels) / len(pixels)
    bits   = [1 if p >= avg else 0 for p in pixels]
    return int("".join(str(b) for b in bits), 2)

def _hamming(h1: int, h2: int) -> int:
    """Count differing bits between two hashes."""
    return bin(h1 ^ h2).count("1")

def compare_images(img1: str, img2: str) -> float:
    """
    Return visual similarity score 0–100.
    100 = identical, 0 = completely different.
    """
    try:
        bits = 16 * 16   # hash size
        h1 = _phash(img1)
        h2 = _phash(img2)
        diff = _hamming(h1, h2)
        return round((1 - diff / bits) * 100, 1)
    except Exception as e:
        print(f"  ✗ Image compare error: {e}")
        return 0.0


# ── Golden file management ────────────────────────────────────────────────────

def _golden_key(file_path: str) -> str:
    return Path(file_path).suffix.lower().lstrip(".")

def _golden_images_dir(key: str) -> Path:
    return GOLDEN / f"images_{key}"

def register_golden(file_path: str):
    """Convert document to images and save as golden reference."""
    key     = _golden_key(file_path)
    img_dir = _golden_images_dir(key)
    shutil.rmtree(img_dir, ignore_errors=True)
    img_dir.mkdir(parents=True, exist_ok=True)

    images = to_images(file_path, str(img_dir))
    if not images:
        print(f"  ⚠  No images produced — golden reference not registered.")
        print(f"     Saving document copy as fallback.")
        dest = GOLDEN / f"reference.{key}"
        shutil.copy2(file_path, dest)
        return

    # Save metadata
    meta = {
        "registered": datetime.datetime.now().isoformat(),
        "source": str(Path(file_path).name),
        "pages": len(images),
        "hashes": {os.path.basename(p): str(_phash(p)) for p in images},
    }
    (img_dir / "meta.json").write_text(json.dumps(meta, indent=2))
    print(f"  ✓ Registered {len(images)} page(s) as golden reference → {img_dir}")


# ── Comparison & report ───────────────────────────────────────────────────────

def compare_with_golden(file_path: str, open_report: bool = False) -> float:
    """
    Compare file against golden reference.
    Returns mean similarity score (0–100).
    """
    key     = _golden_key(file_path)
    img_dir = _golden_images_dir(key)

    if not img_dir.exists():
        print(f"  ⚠  No golden reference for .{key} documents.")
        print(f"     Register one with: python3 tests/visual_check.py --register {file_path}")
        return -1.0

    meta_path = img_dir / "meta.json"
    meta = json.loads(meta_path.read_text()) if meta_path.exists() else {}

    # Convert current file to images
    tmp_dir = REPORTS / "_tmp"
    shutil.rmtree(tmp_dir, ignore_errors=True)
    tmp_dir.mkdir()
    current_images = to_images(file_path, str(tmp_dir))

    if not current_images:
        print("  ⚠  Could not convert current file to images for comparison.")
        return -1.0

    golden_images = sorted(img_dir.glob("*.png"))
    if not golden_images:
        print("  ⚠  No golden PNG images found.")
        return -1.0

    # Compare page by page
    scores = []
    comparisons = []
    for i, (gimg, cimg) in enumerate(zip(golden_images, current_images)):
        score = compare_images(str(gimg), cimg)
        scores.append(score)
        comparisons.append({
            "page":    i + 1,
            "score":   score,
            "golden":  str(gimg),
            "current": cimg,
            "status":  "PASS" if score >= 85 else "WARN" if score >= 70 else "FAIL",
        })

    mean_score = sum(scores) / len(scores) if scores else 0

    # Generate HTML report
    report_path = _generate_report(file_path, comparisons, mean_score)
    print(f"\n  Visual comparison score: {mean_score:.1f}/100")
    print(f"  Report → {report_path}")

    if open_report:
        subprocess.run(["open", str(report_path)], check=False)

    return mean_score


def _img_to_b64(path: str) -> str:
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

def _generate_report(source: str, comparisons: list, mean_score: float) -> Path:
    now    = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    rpath  = REPORTS / f"visual_report_{now}.html"

    rows = []
    for c in comparisons:
        golden_b64  = _img_to_b64(c["golden"])  if os.path.exists(c["golden"])  else ""
        current_b64 = _img_to_b64(c["current"]) if os.path.exists(c["current"]) else ""
        color = {"PASS": "#53AF36", "WARN": "#FFB928", "FAIL": "#F12D2D"}[c["status"]]
        rows.append(f"""
        <tr>
          <td style="padding:8px;border:1px solid #D0D5DD;text-align:center">{c['page']}</td>
          <td style="padding:8px;border:1px solid #D0D5DD;text-align:center">
            <img src="data:image/png;base64,{golden_b64}"
                 style="max-width:300px;border:1px solid #eee">
          </td>
          <td style="padding:8px;border:1px solid #D0D5DD;text-align:center">
            <img src="data:image/png;base64,{current_b64}"
                 style="max-width:300px;border:1px solid #eee">
          </td>
          <td style="padding:8px;border:1px solid #D0D5DD;text-align:center;
                     color:{color};font-weight:600">{c['score']:.1f} — {c['status']}</td>
        </tr>""")

    status_color = "#53AF36" if mean_score >= 85 else "#FFB928" if mean_score >= 70 else "#F12D2D"
    status_text  = "PASS" if mean_score >= 85 else "WARN" if mean_score >= 70 else "FAIL"

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>Visual Regression Report — ArktechX</title>
<style>
  body {{ font-family: Inter, PingFang SC, sans-serif; color: #3D444A;
         background:#F8FAFC; margin:0; padding:32px; }}
  h1   {{ color:#0E1216; font-size:22px; margin:0 0 4px; }}
  .meta{{ color:#8A9199; font-size:13px; margin:0 0 24px; }}
  .score-box {{ display:inline-block; padding:12px 28px;
                background:{status_color}; color:#fff;
                border-radius:8px; font-size:28px; font-weight:700;
                margin-bottom:24px; }}
  table {{ border-collapse:collapse; width:100%; background:#fff;
           box-shadow:0 2px 8px rgba(0,0,0,.08); border-radius:8px;
           overflow:hidden; }}
  th {{ background:#3EC99E; color:#fff; padding:10px 12px;
        text-align:left; font-size:13px; }}
</style>
</head>
<body>
  <h1>ArktechX Visual Regression Report</h1>
  <p class="meta">
    Source: {Path(source).name} &nbsp;|&nbsp;
    Generated: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}
  </p>
  <div class="score-box">{mean_score:.1f}/100 — {status_text}</div>
  <p style="font-size:13px;color:#8A9199">
    Threshold: ≥85 = PASS, ≥70 = WARN, &lt;70 = FAIL (per page)
  </p>
  <table>
    <thead>
      <tr>
        <th style="width:60px">Page</th>
        <th>Golden Reference</th>
        <th>Current Output</th>
        <th style="width:140px">Score</th>
      </tr>
    </thead>
    <tbody>{"".join(rows)}</tbody>
  </table>
</body>
</html>"""

    rpath.write_text(html, encoding="utf-8")
    return rpath


# ── CLI ───────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="ArktechX visual regression check")
    parser.add_argument("file", nargs="?",
                        help="Document to compare against golden")
    parser.add_argument("--register", metavar="FILE",
                        help="Register FILE as new golden reference")
    parser.add_argument("--open", action="store_true",
                        help="Auto-open HTML report after comparison")
    args = parser.parse_args()

    if args.register:
        register_golden(args.register)
    elif args.file:
        score = compare_with_golden(args.file, open_report=args.open)
        sys.exit(0 if score >= 85 else 1)
    else:
        parser.print_help()

if __name__ == "__main__":
    main()

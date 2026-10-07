#!/usr/bin/env python3
"""
render-pptx.py — Render a .pptx to PNG images (one per slide) and/or a PDF.

Used for:
  - style previews in Phase 2 (render the three candidate title slides so the user
    can compare them visually),
  - visual QA in Phase 5 (look at every slide for overflow, overlap, misalignment),
  - PDF export in Phase 6.

Backends, tried in this order:
  1. Microsoft PowerPoint via COM automation (Windows with Office installed).
     Pixel-accurate: this is exactly what the user will see. Also doubles as a
     validity check — if PowerPoint refuses the file, the script fails loudly.
  2. LibreOffice (soffice) headless → PDF, then PDF → PNG via PyMuPDF or pdftoppm
     when either is available. Fonts not installed locally are substituted.
  3. Microsoft PowerPoint via AppleScript (macOS with Office installed) → PDF only.

Usage:
    python scripts/render-pptx.py deck.pptx                      # PNGs into deck-render/
    python scripts/render-pptx.py deck.pptx --out previews/      # choose folder
    python scripts/render-pptx.py deck.pptx --pdf                # also write deck.pdf
    python scripts/render-pptx.py deck.pptx --pdf-only           # only the PDF
    python scripts/render-pptx.py deck.pptx --width 1280         # smaller PNGs
    python scripts/render-pptx.py deck.pptx --grid               # also a labeled contact sheet (needs Pillow)
    python scripts/render-pptx.py deck.pptx --backend soffice    # force a backend

Prints the absolute path of every file it wrote so they can be opened or viewed.
"""

import argparse
import glob
import os
import platform
import shutil
import subprocess
import sys
import tempfile


def log(msg):
    print(msg, flush=True)


# ---------------------------------------------------------------------------
# Backend 1: PowerPoint COM (Windows)
# ---------------------------------------------------------------------------

PS_TEMPLATE = r"""
$ErrorActionPreference = 'Stop'
$src = '{src}'
$out = '{out}'
$pdf = '{pdf}'
$w = {width}
$h = {height}
$doPng = ${do_png}
$app = New-Object -ComObject PowerPoint.Application
$app.DisplayAlerts = 1   # ppAlertsNone: never show a repair/upgrade dialog, throw instead
try {{
    $pres = $app.Presentations.Open($src, $true, $false, $false)
    if ($doPng) {{
        $i = 1
        foreach ($s in $pres.Slides) {{
            $name = 'slide-' + ('{{0:D3}}' -f $i) + '.png'
            $s.Export((Join-Path $out $name), 'PNG', $w, $h)
            $i++
        }}
        Write-Output ("SLIDES=" + ($i - 1))
    }}
    if ($pdf -ne '') {{
        $pres.SaveCopyAs($pdf, 32)   # ppSaveAsPDF
    }}
    $pres.Close()
}} finally {{
    $app.Quit()
}}
"""


def render_powerpoint_com(src, out_dir, width, height, pdf_path, do_png):
    ps = PS_TEMPLATE.format(
        src=src.replace("'", "''"),
        out=out_dir.replace("'", "''"),
        pdf=(pdf_path or "").replace("'", "''"),
        width=width,
        height=height,
        do_png="true" if do_png else "false",
    )
    fd, script = tempfile.mkstemp(suffix=".ps1")
    with os.fdopen(fd, "w", encoding="utf-8-sig") as f:
        f.write(ps)
    try:
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", script],
            capture_output=True, text=True, timeout=600,
        )
    finally:
        os.remove(script)
    if proc.returncode != 0:
        raise RuntimeError("PowerPoint COM export failed:\n" + (proc.stderr or proc.stdout).strip())
    return proc.stdout


# ---------------------------------------------------------------------------
# Backend 2: LibreOffice
# ---------------------------------------------------------------------------

def find_soffice():
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice",
        "/usr/local/bin/soffice",
        "/opt/homebrew/bin/soffice",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def render_soffice(soffice, src, out_dir, width, pdf_path, do_png):
    tmp = tempfile.mkdtemp(prefix="pptx-render-")
    try:
        profile = os.path.join(tmp, "profile")
        proc = subprocess.run(
            [soffice, "--headless", "--norestore", f"-env:UserInstallation=file:///{profile.replace(os.sep, '/')}",
             "--convert-to", "pdf", "--outdir", tmp, src],
            capture_output=True, text=True, timeout=600,
        )
        pdfs = glob.glob(os.path.join(tmp, "*.pdf"))
        if proc.returncode != 0 or not pdfs:
            raise RuntimeError("LibreOffice conversion failed:\n" + (proc.stderr or proc.stdout).strip())
        tmp_pdf = pdfs[0]
        if pdf_path:
            shutil.copyfile(tmp_pdf, pdf_path)
        count = 0
        if do_png:
            count = pdf_to_pngs(tmp_pdf, out_dir, width)
        return count
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def pdf_to_pngs(pdf, out_dir, width):
    try:
        import fitz  # PyMuPDF
    except ImportError:
        fitz = None
    if fitz is not None:
        doc = fitz.open(pdf)
        for i, page in enumerate(doc):
            zoom = width / page.rect.width
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            pix.save(os.path.join(out_dir, f"slide-{i + 1:03d}.png"))
        return len(doc)
    pdftoppm = shutil.which("pdftoppm")
    if pdftoppm:
        subprocess.run([pdftoppm, "-png", "-scale-to-x", str(width), "-scale-to-y", "-1", pdf,
                        os.path.join(out_dir, "slide")], check=True)
        files = sorted(glob.glob(os.path.join(out_dir, "slide-*.png")))
        for i, f in enumerate(files):
            os.replace(f, os.path.join(out_dir, f"slide-{i + 1:03d}.png"))
        return len(files)
    log("  (no PyMuPDF or pdftoppm found: PDF written, but no per-slide PNGs. `pip install pymupdf` to enable.)")
    return 0


# ---------------------------------------------------------------------------
# Backend 3: PowerPoint on macOS via AppleScript (PDF only)
# ---------------------------------------------------------------------------

def render_mac_powerpoint(src, out_dir, width, pdf_path, do_png):
    target = pdf_path or os.path.join(out_dir, "deck.pdf")
    script = f'''
    tell application "Microsoft PowerPoint"
        open POSIX file "{src}"
        set thePres to active presentation
        save thePres in POSIX file "{target}" as save as PDF
        close thePres saving no
    end tell'''
    proc = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=600)
    if proc.returncode != 0:
        raise RuntimeError("PowerPoint (macOS) export failed:\n" + proc.stderr.strip())
    if do_png:
        return pdf_to_pngs(target, out_dir, width)
    return 0


# ---------------------------------------------------------------------------
# Contact sheet
# ---------------------------------------------------------------------------

def make_grid(out_dir, cols=3, thumb_w=640):
    try:
        from PIL import Image, ImageDraw
    except ImportError:
        log("  (Pillow not installed: skipping --grid. `pip install Pillow` to enable.)")
        return None
    files = sorted(glob.glob(os.path.join(out_dir, "slide-*.png")))
    if not files:
        return None
    thumbs = []
    for f in files:
        im = Image.open(f).convert("RGB")
        ratio = thumb_w / im.width
        thumbs.append(im.resize((thumb_w, int(im.height * ratio))))
    th = thumbs[0].height
    pad, label_h = 16, 28
    rows = (len(thumbs) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * (thumb_w + pad) + pad, rows * (th + label_h + pad) + pad), "#202020")
    draw = ImageDraw.Draw(sheet)
    for i, t in enumerate(thumbs):
        r, c = divmod(i, cols)
        x = pad + c * (thumb_w + pad)
        y = pad + r * (th + label_h + pad)
        draw.text((x, y + 6), f"slide {i + 1}", fill="#ffffff")
        sheet.paste(t, (x, y + label_h))
    path = os.path.join(out_dir, "contact-sheet.png")
    sheet.save(path)
    return path


# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pptx")
    ap.add_argument("--out", help="output folder for PNGs (default: <deck>-render/)")
    ap.add_argument("--pdf", action="store_true", help="also write <deck>.pdf next to the deck")
    ap.add_argument("--pdf-only", action="store_true", help="write only the PDF")
    ap.add_argument("--pdf-path", help="explicit PDF output path")
    ap.add_argument("--width", type=int, default=1920, help="PNG width in pixels (default 1920)")
    ap.add_argument("--grid", action="store_true", help="also build contact-sheet.png")
    ap.add_argument("--backend", choices=["auto", "powerpoint", "soffice"], default="auto")
    args = ap.parse_args()

    src = os.path.abspath(args.pptx)
    if not os.path.isfile(src):
        raise SystemExit(f"File not found: {src}")
    base = os.path.splitext(os.path.basename(src))[0]
    out_dir = os.path.abspath(args.out or os.path.join(os.path.dirname(src), f"{base}-render"))
    do_png = not args.pdf_only
    pdf_path = None
    if args.pdf or args.pdf_only or args.pdf_path:
        pdf_path = os.path.abspath(args.pdf_path or os.path.join(os.path.dirname(src), f"{base}.pdf"))
    if do_png:
        os.makedirs(out_dir, exist_ok=True)
        for old in glob.glob(os.path.join(out_dir, "slide-*.png")):
            os.remove(old)
    height = round(args.width * 9 / 16)

    system = platform.system()
    backends = []
    if args.backend in ("auto", "powerpoint"):
        if system == "Windows":
            backends.append("powerpoint-com")
        elif system == "Darwin" and os.path.isdir("/Applications/Microsoft PowerPoint.app"):
            backends.append("powerpoint-mac")
    if args.backend in ("auto", "soffice"):
        backends.append("soffice")

    errors = []
    for be in backends:
        try:
            log(f"Rendering with {be} ...")
            if be == "powerpoint-com":
                render_powerpoint_com(src, out_dir, args.width, height, pdf_path, do_png)
            elif be == "powerpoint-mac":
                render_mac_powerpoint(src, out_dir, args.width, pdf_path, do_png)
            else:
                soffice = find_soffice()
                if not soffice:
                    raise RuntimeError("LibreOffice (soffice) not found. Install it from https://www.libreoffice.org/download/")
                render_soffice(soffice, src, out_dir, args.width, pdf_path, do_png)
            break
        except Exception as e:  # noqa: BLE001
            errors.append(f"[{be}] {e}")
            log(f"  {be} failed: {e}")
    else:
        raise SystemExit("All render backends failed:\n" + "\n".join(errors))

    written = []
    if do_png:
        written += sorted(glob.glob(os.path.join(out_dir, "slide-*.png")))
        if args.grid:
            g = make_grid(out_dir)
            if g:
                written.append(g)
    if pdf_path and os.path.exists(pdf_path):
        written.append(pdf_path)

    log("")
    log(f"Rendered {len([w for w in written if w.endswith('.png') and 'contact-sheet' not in w])} slide image(s)")
    for w in written:
        log(w)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
check-pptx.py — Structural validation plus a content outline for a .pptx.

Run it after every build, before rendering. It catches the defects that make
PowerPoint show "needs repair" or silently drop content, and prints an outline
you can read to confirm nothing is missing or out of order.

Checks:
  - every part is well-formed XML
  - every relationship target exists inside the package
  - every slide is registered in [Content_Types].xml and presentation.xml
  - charts: every axis id a plot references is declared (pptxgenjs combo-chart trap)
  - shapes that extend past the slide edge
  - text that probably overflows its box (heuristic, flagged as "possible")
  - leftover placeholder / prompt text ("Click to add", "lorem ipsum", "TODO", "[insert")
  - empty slides, slides without a title
  - hex-looking colors with a leading '#' or 8 digits (pptxgenjs corrupts these)

Usage:
    python scripts/check-pptx.py deck.pptx           # checks + outline
    python scripts/check-pptx.py deck.pptx --quiet   # checks only
    python scripts/check-pptx.py deck.pptx --json    # machine-readable

Exit code 1 if any ERROR was found; warnings do not fail the run. Only stdlib.
"""

import argparse
import json
import math
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from posixpath import normpath, join as pjoin, dirname

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
}
EMU_PER_IN = 914400
PLACEHOLDER_RX = re.compile(r"click to (add|edit)|lorem ipsum|\bTODO\b|\[insert|\bx{3,}\b|placeholder text", re.I)


class Report:
    def __init__(self):
        self.errors, self.warnings, self.slides = [], [], []

    def error(self, msg):
        self.errors.append(msg)

    def warn(self, msg):
        self.warnings.append(msg)


def parse(zf, name):
    return ET.fromstring(zf.read(name))


def resolve(base_part, target):
    if target.startswith("/"):
        return target.lstrip("/")
    return normpath(pjoin(dirname(base_part), target))


# ---------------------------------------------------------------------------
# package-level checks
# ---------------------------------------------------------------------------

def check_package(zf, rep):
    names = set(zf.namelist())
    for name in names:
        if name.endswith(".xml") or name.endswith(".rels"):
            try:
                parse(zf, name)
            except ET.ParseError as e:
                rep.error(f"{name}: malformed XML ({e})")

    for name in [n for n in names if n.endswith(".rels")]:
        try:
            root = parse(zf, name)
        except ET.ParseError:
            continue
        src_part = name.replace("_rels/", "").replace(".rels", "")
        for rel in root.findall("rel:Relationship", NS):
            if rel.get("TargetMode") == "External":
                continue
            target = resolve(src_part, rel.get("Target", ""))
            if target not in names:
                rep.error(f"{name}: relationship {rel.get('Id')} points to missing part '{target}'")

    try:
        ct = parse(zf, "[Content_Types].xml")
        overrides = {o.get("PartName").lstrip("/") for o in ct.findall("ct:Override", NS)}
        for name in names:
            if re.fullmatch(r"ppt/slides/slide\d+\.xml", name) and name not in overrides:
                rep.error(f"[Content_Types].xml has no Override for {name}")
    except KeyError:
        rep.error("[Content_Types].xml missing")


def slide_parts_in_order(zf, rep):
    pres = parse(zf, "ppt/presentation.xml")
    rels = parse(zf, "ppt/_rels/presentation.xml.rels")
    rid_map = {r.get("Id"): resolve("ppt/presentation.xml", r.get("Target")) for r in rels.findall("rel:Relationship", NS)}
    order = []
    for sld in pres.findall("p:sldIdLst/p:sldId", NS):
        rid = sld.get(f"{{{NS['r']}}}id")
        part = rid_map.get(rid)
        if not part:
            rep.error(f"presentation.xml: sldId r:id={rid} has no relationship")
            continue
        order.append(part)
    registered = set(order)
    for name in zf.namelist():
        if re.fullmatch(r"ppt/slides/slide\d+\.xml", name) and name not in registered:
            rep.warn(f"{name} exists but is not listed in presentation.xml (orphan slide; invisible)")
    sz = pres.find("p:sldSz", NS)
    cx = int(sz.get("cx")) if sz is not None else 12192000
    cy = int(sz.get("cy")) if sz is not None else 6858000
    return order, cx, cy


# ---------------------------------------------------------------------------
# chart checks
# ---------------------------------------------------------------------------

def check_charts(zf, rep):
    for name in zf.namelist():
        if not re.fullmatch(r"ppt/charts/chart\d+\.xml", name):
            continue
        try:
            root = parse(zf, name)
        except ET.ParseError:
            continue
        plot = root.find(".//c:plotArea", NS)
        if plot is None:
            continue
        declared = set()
        for ax_tag in ("c:catAx", "c:valAx", "c:dateAx", "c:serAx"):
            for ax in plot.findall(ax_tag, NS):
                aid = ax.find("c:axId", NS)
                if aid is not None:
                    declared.add(aid.get("val"))
        # Each 2-D chart group must have at least two of its axis ids declared (cat + val).
        # pptxgenjs may emit a third, undeclared series-axis id; PowerPoint tolerates that,
        # so only flag groups whose primary pair is incomplete.
        for child in plot:
            tag = child.tag.split("}")[1]
            if not tag.endswith("Chart") or tag in ("pieChart", "pie3DChart", "doughnutChart", "ofPieChart"):
                continue
            referenced = [aid.get("val") for aid in child.findall("c:axId", NS)]
            if not referenced:
                continue
            found = [a for a in referenced if a in declared]
            if len(found) < min(2, len(referenced)):
                rep.error(f"{name}: <c:{tag}> references axis id(s) {referenced} but only {found} are declared — "
                          "PowerPoint will discard this chart. In pptxgenjs, a secondary-axis series needs both "
                          "`valAxes` and `catAxes` (two entries each) on the chart options.")


# ---------------------------------------------------------------------------
# slide checks + outline
# ---------------------------------------------------------------------------

def shape_text(sp):
    paras = []
    for p in sp.findall(".//a:p", NS):
        runs = [t.text or "" for t in p.findall(".//a:t", NS)]
        paras.append("".join(runs))
    return paras


def run_sizes(sp):
    sizes = [int(r.get("sz")) / 100 for r in sp.findall(".//a:rPr[@sz]", NS)]
    sizes += [int(r.get("sz")) / 100 for r in sp.findall(".//a:defRPr[@sz]", NS)]
    return sizes


def estimate_overflow(sp, paras, w_in, h_in):
    """Very rough: average glyph width ≈ 0.5 em, line height ≈ 1.2 em."""
    sizes = run_sizes(sp)
    pt = max(sizes) if sizes else 18.0
    if w_in <= 0 or h_in <= 0:
        return None
    body = sp.find(".//a:bodyPr", NS)
    autofit = body is not None and (body.find("a:normAutofit", NS) is not None or body.find("a:spAutoFit", NS) is not None)
    if autofit:
        return None
    char_w_in = 0.5 * pt / 72
    chars_per_line = max(1, int(w_in / char_w_in))
    lines = sum(max(1, math.ceil(len(t) / chars_per_line)) for t in paras)
    need_in = lines * pt * 1.2 / 72
    if need_in > h_in * 1.15 and need_in - h_in > 0.15:
        return need_in, pt, lines
    return None


def check_slides(zf, rep, order, cx, cy, want_outline):
    for idx, part in enumerate(order, start=1):
        try:
            root = parse(zf, part)
        except (ET.ParseError, KeyError):
            continue
        info = {"slide": idx, "part": part, "title": "", "texts": [], "notes": "", "shapes": 0, "pictures": 0, "charts": 0, "tables": 0}
        tree = root.find(".//p:spTree", NS)
        if tree is None:
            rep.error(f"slide {idx}: no spTree")
            continue

        for el in tree.iter():
            tag = el.tag.split("}")[1]
            if tag not in ("sp", "pic", "graphicFrame", "cxnSp"):
                continue
            info["shapes"] += 1
            if tag == "pic":
                info["pictures"] += 1
            if tag == "graphicFrame":
                if el.find(".//c:chart", NS) is not None or el.find(".//{http://schemas.openxmlformats.org/drawingml/2006/chart}chart") is not None:
                    info["charts"] += 1
                if el.find(".//a:tbl", NS) is not None:
                    info["tables"] += 1
                    for cell in el.findall(".//a:tc", NS):
                        txt = "".join(t.text or "" for t in cell.findall(".//a:t", NS))
                        if txt:
                            info["texts"].append(txt)

            name = ""
            cnv = el.find(".//p:cNvPr", NS)
            if cnv is not None:
                name = cnv.get("name", "")
            ph = el.find(".//p:ph", NS)
            ph_type = ph.get("type", "body") if ph is not None else None

            xfrm = el.find(".//a:xfrm", NS)
            off = xfrm.find("a:off", NS) if xfrm is not None else None
            ext = xfrm.find("a:ext", NS) if xfrm is not None else None
            x = y = w = h = None
            if off is not None and ext is not None:
                x, y, w, h = (int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy")))
                if x < -EMU_PER_IN // 20 or y < -EMU_PER_IN // 20 or x + w > cx + EMU_PER_IN // 20 or y + h > cy + EMU_PER_IN // 20:
                    rep.warn(f"slide {idx}: '{name}' extends past the slide edge "
                             f"(x={x/EMU_PER_IN:.2f} y={y/EMU_PER_IN:.2f} w={w/EMU_PER_IN:.2f} h={h/EMU_PER_IN:.2f} in)")

            paras = shape_text(el) if tag == "sp" else []
            joined = " ".join(p for p in paras if p).strip()
            if ph_type in ("sldNum", "ftr", "dt"):
                continue
            if joined:
                if ph_type in ("title", "ctrTitle") and not info["title"]:
                    info["title"] = joined
                else:
                    info["texts"].append(joined)
                if PLACEHOLDER_RX.search(joined):
                    rep.warn(f"slide {idx}: leftover placeholder text in '{name}': {joined[:60]!r}")
                if w and h:
                    est = estimate_overflow(el, [p for p in paras], w / EMU_PER_IN, h / EMU_PER_IN)
                    if est:
                        need, pt, lines = est
                        rep.warn(f"slide {idx}: possible text overflow in '{name}' "
                                 f"(~{lines} lines at {pt:g}pt need ~{need:.2f} in, box is {h/EMU_PER_IN:.2f} in). "
                                 "Verify in the render; shorten, enlarge the box, or split the slide.")
            elif ph is not None and ph_type in ("title", "ctrTitle", "body", "subTitle") and tag == "sp":
                rep.warn(f"slide {idx}: empty placeholder '{name}' will show prompt text in edit view — remove it or fill it")

        # speaker notes
        rel_name = part.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
        if rel_name in zf.namelist():
            rels = parse(zf, rel_name)
            for r in rels.findall("rel:Relationship", NS):
                if r.get("Type", "").endswith("/notesSlide"):
                    notes_part = resolve(part, r.get("Target"))
                    if notes_part in zf.namelist():
                        nroot = parse(zf, notes_part)
                        texts = []
                        for sp in nroot.findall(".//p:sp", NS):
                            ph = sp.find(".//p:ph", NS)
                            if ph is not None and ph.get("type") == "body":
                                texts += [t for t in shape_text(sp) if t]
                        info["notes"] = "\n".join(texts)

        if info["shapes"] == 0:
            rep.warn(f"slide {idx}: empty slide")
        rep.slides.append(info)

    # whole-deck consistency
    titles = [s["title"] for s in rep.slides]
    if len(set(t for t in titles if t)) < len([t for t in titles if t]):
        dupes = {t for t in titles if t and titles.count(t) > 1}
        rep.warn("duplicate slide titles: " + "; ".join(sorted(dupes)))


def check_color_literals(zf, rep):
    bad = re.compile(r'val="(#[0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})"')
    for name in zf.namelist():
        if re.fullmatch(r"ppt/(slides|slideLayouts|slideMasters)/[^/]+\.xml", name):
            xml = zf.read(name).decode("utf-8", "replace")
            m = bad.search(xml)
            if m:
                rep.error(f"{name}: invalid color literal {m.group(1)!r} (use 6-digit hex without '#'; "
                          "pptxgenjs needs `transparency` for alpha)")


# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pptx")
    ap.add_argument("--quiet", action="store_true", help="skip the outline")
    ap.add_argument("--json", action="store_true", help="print JSON instead of text")
    args = ap.parse_args()

    if not zipfile.is_zipfile(args.pptx):
        print(f"ERROR: not a .pptx (zip) file: {args.pptx}")
        sys.exit(1)

    rep = Report()
    with zipfile.ZipFile(args.pptx) as zf:
        check_package(zf, rep)
        try:
            order, cx, cy = slide_parts_in_order(zf, rep)
        except (KeyError, ET.ParseError) as e:
            rep.error(f"cannot read presentation.xml: {e}")
            order, cx, cy = [], 12192000, 6858000
        check_charts(zf, rep)
        check_color_literals(zf, rep)
        check_slides(zf, rep, order, cx, cy, not args.quiet)

    if args.json:
        print(json.dumps({"errors": rep.errors, "warnings": rep.warnings, "slides": rep.slides,
                          "slide_size_in": [round(cx / EMU_PER_IN, 3), round(cy / EMU_PER_IN, 3)]}, indent=2, ensure_ascii=False))
        sys.exit(1 if rep.errors else 0)

    print(f"{args.pptx}: {len(order)} slide(s), canvas {cx / EMU_PER_IN:.3f} x {cy / EMU_PER_IN:.3f} in")
    if not args.quiet:
        print("")
        for s in rep.slides:
            extras = []
            if s["pictures"]:
                extras.append(f"{s['pictures']} image(s)")
            if s["charts"]:
                extras.append(f"{s['charts']} chart(s)")
            if s["tables"]:
                extras.append(f"{s['tables']} table(s)")
            print(f"--- slide {s['slide']}: {s['title'] or '(no title)'}" + (f"  [{', '.join(extras)}]" if extras else ""))
            for t in s["texts"]:
                print(f"    {t[:160]}")
            if s["notes"]:
                print(f"    notes: {s['notes'][:160]}")
        print("")
    for w in rep.warnings:
        print(f"WARNING: {w}")
    for e in rep.errors:
        print(f"ERROR: {e}")
    print("")
    print(f"{len(rep.errors)} error(s), {len(rep.warnings)} warning(s)")
    sys.exit(1 if rep.errors else 0)


if __name__ == "__main__":
    main()

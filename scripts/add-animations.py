#!/usr/bin/env python3
"""
add-animations.py — Add slide transitions and entrance animations to a .pptx.

pptxgenjs cannot write animations. This script post-processes the finished deck
and injects native PowerPoint timing XML, so the result is fully editable in
PowerPoint's Animations pane (nothing is baked into images).

Usage:
    python scripts/add-animations.py deck.pptx [options]

Transitions (between slides):
    --transition fade|push|wipe|cover|split|morph|none   (default: fade)
    --speed slow|med|fast                                 (default: med)
    --direction l|r|u|d                                   push/wipe/cover direction (default: l)
    --advance-after MS      auto-advance every slide after MS milliseconds (kiosk mode)

Entrance animations (objects on each slide):
    --entrance fade|rise|wipe|fly|zoom|none   (default: rise — fade + small upward drift,
                                               the PowerPoint equivalent of a staggered reveal)
    --mode auto|click       auto = everything plays on slide entry, staggered (default)
                            click = each object appears on its own click, in z-order
    --duration MS           effect length (default: 500)
    --stagger MS            delay between objects in auto mode (default: 120)
    --initial-delay MS      delay before the first object in auto mode (default: 150)
    --include-large         also animate shapes covering >= 85% of the slide (backgrounds)
    --slides 2-9,12         only touch these slide numbers (1-based); default: all
    --skip 1                never touch these slide numbers (default: none)
    --clear                 remove existing transitions/animations first

Examples:
    python scripts/add-animations.py deck.pptx
    python scripts/add-animations.py deck.pptx --transition push --entrance fade --stagger 200
    python scripts/add-animations.py deck.pptx --entrance none --transition morph
    python scripts/add-animations.py deck.pptx --mode click --skip 1

Only stdlib is used. Re-running replaces what the previous run added.
"""

import argparse
import os
import re
import shutil
import sys
import tempfile
import zipfile
from xml.dom import minidom

P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
SHAPE_TAGS = ("p:sp", "p:pic", "p:graphicFrame", "p:grpSp", "p:cxnSp")
CHROME_PH_TYPES = {"sldNum", "ftr", "dt"}

TRANSITIONS = {
    "fade": lambda d: "<p:fade/>",
    "push": lambda d: f'<p:push dir="{d}"/>',
    "wipe": lambda d: f'<p:wipe dir="{d}"/>',
    "cover": lambda d: f'<p:cover dir="{d}"/>',
    "split": lambda d: '<p:split orient="horz" dir="out"/>',
}

PRESETS = {
    # name: (presetID, presetSubtype)
    "fade": (10, 0),
    "rise": (10, 0),
    "wipe": (22, 8),
    "fly": (2, 4),
    "zoom": (23, 16),
}


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def parse_slide_list(spec, count):
    if not spec:
        return set()
    out = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            out.update(range(int(a), int(b) + 1))
        else:
            out.add(int(part))
    return {n for n in out if 1 <= n <= count}


def slide_order(zf):
    """Return slide part names in presentation order."""
    pres = zf.read("ppt/presentation.xml").decode("utf-8")
    rels = zf.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
    rid_to_target = dict(re.findall(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    # attribute order varies, so also try Target before Id
    for target, rid in re.findall(r'<Relationship[^>]*Target="([^"]+)"[^>]*Id="([^"]+)"', rels):
        rid_to_target.setdefault(rid, target)
    order = []
    for rid in re.findall(r'<p:sldId [^>]*r:id="([^"]+)"', pres):
        target = rid_to_target.get(rid)
        if target:
            order.append("ppt/" + target.lstrip("/").replace("ppt/", "", 1) if not target.startswith("ppt/") else target)
    return order


def slide_size(zf):
    pres = zf.read("ppt/presentation.xml").decode("utf-8")
    m = re.search(r'<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"', pres)
    if not m:
        return 12192000, 6858000
    return int(m.group(1)), int(m.group(2))


def first_child(node, name):
    for c in node.childNodes:
        if c.nodeType == c.ELEMENT_NODE and c.tagName == name:
            return c
    return None


def find_desc(node, name):
    found = node.getElementsByTagName(name)
    return found[0] if found else None


def collect_shapes(xml, slide_cx, slide_cy, include_large):
    """Return list of (spid, has_text, kind) for animatable top-level shapes."""
    doc = minidom.parseString(xml.encode("utf-8"))
    sp_tree = doc.getElementsByTagName("p:spTree")
    if not sp_tree:
        return []
    shapes = []
    for node in sp_tree[0].childNodes:
        if node.nodeType != node.ELEMENT_NODE or node.tagName not in SHAPE_TAGS:
            continue
        cnv = find_desc(node, "p:cNvPr")
        if cnv is None or not cnv.getAttribute("id"):
            continue
        spid = cnv.getAttribute("id")

        ph = find_desc(node, "p:ph")
        if ph is not None and ph.getAttribute("type") in CHROME_PH_TYPES:
            continue

        has_text = False
        tx = find_desc(node, "p:txBody")
        if tx is not None:
            has_text = any(t.firstChild and t.firstChild.nodeValue.strip() for t in tx.getElementsByTagName("a:t"))

        # skip empty, unfilled text boxes
        if node.tagName == "p:sp" and not has_text:
            sppr = first_child(node, "p:spPr")
            if sppr is not None and first_child(sppr, "a:noFill") is not None:
                continue

        # skip full-bleed background panels unless asked
        if not include_large:
            xfrm = find_desc(node, "a:xfrm")
            ext = find_desc(xfrm, "a:ext") if xfrm is not None else None
            if ext is not None:
                try:
                    cx, cy = int(ext.getAttribute("cx")), int(ext.getAttribute("cy"))
                    if cx * cy >= 0.85 * slide_cx * slide_cy:
                        continue
                except ValueError:
                    pass

        shapes.append((spid, has_text, node.tagName))
    return shapes


# ---------------------------------------------------------------------------
# XML builders
# ---------------------------------------------------------------------------

class IdGen:
    def __init__(self, start=1):
        self.n = start - 1

    def next(self):
        self.n += 1
        return self.n


def effect_xml(ids, spid, kind, duration, delay, node_type):
    preset_id, subtype = PRESETS[kind]
    eid = ids.next()
    parts = [
        f'<p:par><p:cTn id="{eid}" presetID="{preset_id}" presetClass="entr" presetSubtype="{subtype}" '
        f'fill="hold" grpId="0" nodeType="{node_type}"><p:stCondLst><p:cond delay="{delay}"/></p:stCondLst><p:childTnLst>'
    ]
    # visibility set (every entrance effect starts with this)
    parts.append(
        f'<p:set><p:cBhvr><p:cTn id="{ids.next()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>'
        f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>'
        f'</p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
    )

    def anim(attr, v0, v1, additive=True):
        add = ' additive="base"' if additive else ""
        return (
            f'<p:anim calcmode="lin" valueType="num"><p:cBhvr{add}><p:cTn id="{ids.next()}" dur="{duration}" fill="hold"/>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl><p:attrNameLst><p:attrName>{attr}</p:attrName></p:attrNameLst></p:cBhvr>'
            f'<p:tavLst><p:tav tm="0"><p:val><p:strVal val="{v0}"/></p:val></p:tav>'
            f'<p:tav tm="100000"><p:val><p:strVal val="{v1}"/></p:val></p:tav></p:tavLst></p:anim>'
        )

    def anim_effect(filt):
        return (
            f'<p:animEffect transition="in" filter="{filt}"><p:cBhvr><p:cTn id="{ids.next()}" dur="{duration}"/>'
            f'<p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl></p:cBhvr></p:animEffect>'
        )

    if kind == "fade":
        parts.append(anim_effect("fade"))
    elif kind == "rise":
        parts.append(anim_effect("fade"))
        parts.append(anim("ppt_y", "#ppt_y+0.04", "#ppt_y"))
    elif kind == "wipe":
        parts.append(anim_effect("wipe(left)"))
    elif kind == "fly":
        parts.append(anim("ppt_x", "#ppt_x", "#ppt_x"))
        parts.append(anim("ppt_y", "1+#ppt_h/2", "#ppt_y"))
    elif kind == "zoom":
        parts.append(anim("ppt_w", "0", "#ppt_w"))
        parts.append(anim("ppt_h", "0", "#ppt_h"))
        parts.append(anim_effect("fade"))

    parts.append("</p:childTnLst></p:cTn></p:par>")
    return "".join(parts)


def timing_xml(shapes, kind, mode, duration, stagger, initial_delay):
    ids = IdGen()
    root_id = ids.next()   # 1
    seq_id = ids.next()    # 2
    groups = []

    if mode == "auto":
        click_id = ids.next()
        inner_id = ids.next()
        effects = []
        for i, (spid, _, _) in enumerate(shapes):
            node_type = "afterEffect" if i == 0 else "withEffect"
            delay = initial_delay + i * stagger if i else initial_delay
            effects.append(effect_xml(ids, spid, kind, duration, delay, node_type))
        groups.append(
            f'<p:par><p:cTn id="{click_id}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
            f'<p:cond evt="onBegin" delay="0"><p:tn val="{seq_id}"/></p:cond></p:stCondLst><p:childTnLst>'
            f'<p:par><p:cTn id="{inner_id}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
            + "".join(effects)
            + "</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>"
        )
    else:  # click
        for spid, _, _ in shapes:
            click_id = ids.next()
            inner_id = ids.next()
            eff = effect_xml(ids, spid, kind, duration, 0, "clickEffect")
            groups.append(
                f'<p:par><p:cTn id="{click_id}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>'
                f'<p:par><p:cTn id="{inner_id}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>'
                + eff
                + "</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>"
            )

    bld = []
    for spid, has_text, tag in shapes:
        if tag == "p:graphicFrame":
            bld.append(f'<p:bldGraphic spid="{spid}" grpId="0"><p:bldAsOne/></p:bldGraphic>')
        elif tag == "p:sp" and has_text:
            bld.append(f'<p:bldP spid="{spid}" grpId="0"/>')

    return (
        f'<p:timing><p:tnLst><p:par><p:cTn id="{root_id}" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
        f'<p:seq concurrent="1" nextAc="seek"><p:cTn id="{seq_id}" dur="indefinite" nodeType="mainSeq"><p:childTnLst>'
        + "".join(groups)
        + "</p:childTnLst></p:cTn>"
        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
        "</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>"
        + (f"<p:bldLst>{''.join(bld)}</p:bldLst>" if bld else "")
        + "</p:timing>"
    )


def transition_xml(kind, speed, direction, advance_after):
    adv = ""
    if advance_after:
        adv = f' advAuto="1" advTm="{advance_after}"'
    if kind == "morph":
        # Morph needs the p159 namespace and an AlternateContent wrapper for older clients.
        return (
            '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
            '<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">'
            f'<p:transition spd="{speed}"{adv}><p159:morph option="byObject"/></p:transition></mc:Choice>'
            f'<mc:Fallback><p:transition spd="{speed}"{adv}><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>'
        )
    body = TRANSITIONS[kind](direction)
    return f'<p:transition spd="{speed}"{adv}>{body}</p:transition>'


# ---------------------------------------------------------------------------
# slide patching
# ---------------------------------------------------------------------------

STRIP_PATTERNS = [
    re.compile(r"<mc:AlternateContent[^>]*>\s*<mc:Choice[^>]*>\s*<p:transition.*?</mc:AlternateContent>", re.S),
    re.compile(r"<p:transition\b[^>]*/>", re.S),
    re.compile(r"<p:transition\b.*?</p:transition>", re.S),
    re.compile(r"<p:timing>.*?</p:timing>", re.S),
]


def strip_existing(xml):
    for pat in STRIP_PATTERNS:
        xml = pat.sub("", xml)
    return xml


def insert_after_clrmap(xml, fragment):
    """p:sld children order: cSld, clrMapOvr, transition, timing, extLst."""
    m = re.search(r"</p:clrMapOvr>", xml)
    if m:
        return xml[: m.end()] + fragment + xml[m.end():]
    m = re.search(r"</p:cSld>", xml)
    if not m:
        raise SystemExit("slide XML has no </p:cSld>; cannot patch")
    return xml[: m.end()] + fragment + xml[m.end():]


def patch_slide(xml, args, slide_cx, slide_cy):
    xml = strip_existing(xml)
    fragment = ""
    if args.transition != "none":
        fragment += transition_xml(args.transition, args.speed, args.direction, args.advance_after)
    count = 0
    if args.entrance != "none":
        shapes = collect_shapes(xml, slide_cx, slide_cy, args.include_large)
        if shapes:
            fragment += timing_xml(shapes, args.entrance, args.mode, args.duration, args.stagger, args.initial_delay)
            count = len(shapes)
    if fragment:
        xml = insert_after_clrmap(xml, fragment)
    return xml, count


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("pptx")
    ap.add_argument("--transition", default="fade", choices=["fade", "push", "wipe", "cover", "split", "morph", "none"])
    ap.add_argument("--speed", default="med", choices=["slow", "med", "fast"])
    ap.add_argument("--direction", default="l", choices=["l", "r", "u", "d"])
    ap.add_argument("--advance-after", type=int, default=0, metavar="MS")
    ap.add_argument("--entrance", default="rise", choices=["fade", "rise", "wipe", "fly", "zoom", "none"])
    ap.add_argument("--mode", default="auto", choices=["auto", "click"])
    ap.add_argument("--duration", type=int, default=500, metavar="MS")
    ap.add_argument("--stagger", type=int, default=120, metavar="MS")
    ap.add_argument("--initial-delay", type=int, default=150, metavar="MS")
    ap.add_argument("--include-large", action="store_true")
    ap.add_argument("--slides", default="", help="only these slide numbers, e.g. 2-9,12")
    ap.add_argument("--skip", default="", help="never touch these slide numbers")
    ap.add_argument("--clear", action="store_true", help="only remove existing transitions and animations")
    args = ap.parse_args()

    if not zipfile.is_zipfile(args.pptx):
        raise SystemExit(f"Not a .pptx file: {args.pptx}")

    with zipfile.ZipFile(args.pptx) as zf:
        order = slide_order(zf)
        cx, cy = slide_size(zf)
    total = len(order)
    only = parse_slide_list(args.slides, total) or set(range(1, total + 1))
    skip = parse_slide_list(args.skip, total)
    targets = {name: i + 1 for i, name in enumerate(order) if (i + 1) in only and (i + 1) not in skip}

    if args.clear:
        args.transition = "none"
        args.entrance = "none"

    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".pptx", dir=os.path.dirname(os.path.abspath(args.pptx)))
    os.close(tmp_fd)
    report = []
    try:
        with zipfile.ZipFile(args.pptx, "r") as zin, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename in targets:
                    xml, n = patch_slide(data.decode("utf-8"), args, cx, cy)
                    data = xml.encode("utf-8")
                    report.append((targets[item.filename], n))
                zout.writestr(item, data)
        shutil.move(tmp_path, args.pptx)
    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)

    report.sort()
    if args.clear:
        print(f"Cleared transitions and animations on {len(report)} slide(s) in {args.pptx}")
        return
    print(f"Patched {len(report)}/{total} slide(s) in {args.pptx}")
    print(f"  transition: {args.transition} ({args.speed})" if args.transition != "none" else "  transition: none")
    if args.entrance != "none":
        print(f"  entrance: {args.entrance}, mode={args.mode}, duration={args.duration}ms, stagger={args.stagger}ms")
        print("  animated objects per slide: " + ", ".join(f"{n}:{c}" for n, c in report))


if __name__ == "__main__":
    main()

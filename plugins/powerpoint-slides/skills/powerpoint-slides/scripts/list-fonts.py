#!/usr/bin/env python3
"""
list-fonts.py — Check which fonts are installed on this machine.

PowerPoint renders whatever font *name* is written into the deck with whatever
font the viewer's computer has. A distinctive Google Font that is not installed
is silently replaced (usually by Calibri or Arial), which destroys the design.
Run this before committing to a font pair, and again on the presenting machine.

Usage:
    python scripts/list-fonts.py                              # list every installed family
    python scripts/list-fonts.py "Source Serif 4" "DM Sans"   # check specific families
    python scripts/list-fonts.py --grep serif                 # filter the list

Exit code 1 if any requested family is missing. Only stdlib.
"""

import argparse
import os
import platform
import re
import subprocess
import sys


def windows_fonts():
    families = set()
    try:
        import winreg
        for hive, path in (
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
            (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
        ):
            try:
                key = winreg.OpenKey(hive, path)
            except OSError:
                continue
            i = 0
            while True:
                try:
                    name, _, _ = winreg.EnumValue(key, i)
                except OSError:
                    break
                i += 1
                families.add(normalize(name))
    except ImportError:
        pass
    if families:
        return families  # registry names are clean family names; file stems are not
    for folder in (os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
                   os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts")):
        if os.path.isdir(folder):
            for f in os.listdir(folder):
                stem, ext = os.path.splitext(f)
                if ext.lower() not in (".ttf", ".otf", ".ttc"):
                    continue  # skip legacy bitmap .fon files and non-font files
                families.add(normalize(stem))
    return families


def fc_list_fonts():
    try:
        out = subprocess.run(["fc-list", ":", "family"], capture_output=True, text=True, timeout=30).stdout
    except (OSError, subprocess.TimeoutExpired):
        return set()
    families = set()
    for line in out.splitlines():
        for fam in line.split(","):
            if fam.strip():
                families.add(fam.strip())
    return families


def mac_fonts():
    fams = fc_list_fonts()
    if fams:
        return fams
    try:
        out = subprocess.run(["system_profiler", "SPFontsDataType"], capture_output=True, text=True, timeout=120).stdout
        return {m.strip() for m in re.findall(r"Family:\s*(.+)", out)}
    except (OSError, subprocess.TimeoutExpired):
        return set()


def normalize(name):
    # "Source Serif 4 Bold (TrueType)" -> "Source Serif 4 Bold"
    name = re.sub(r"\s*\((TrueType|OpenType|All res)\)\s*$", "", name, flags=re.I)
    return name.strip()


STYLE_WORDS = r"(?:\s+(?:Thin|Extra ?Light|Light|Regular|Medium|Semi ?Bold|Demi ?Bold|Bold|Extra ?Bold|Black|Heavy|Italic|Oblique|Condensed|Narrow|Wide|Variable|VF|[0-9]{3}))+$"


def family_of(name):
    base = re.sub(STYLE_WORDS, "", name, flags=re.I).strip()
    base = re.sub(r"[-_]+(?:Regular|Bold|Italic|Light|Medium|Black|Variable|VF|Semi|Extra|Thin)[A-Za-z]*$", "", base, flags=re.I)
    return base.replace("-", " ").replace("_", " ").strip()


def installed_families():
    system = platform.system()
    if system == "Windows":
        raw = windows_fonts()
    elif system == "Darwin":
        raw = mac_fonts()
    else:
        raw = fc_list_fonts()
    return {family_of(n) for n in raw if n}, raw


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("fonts", nargs="*", help="font family names to check")
    ap.add_argument("--grep", help="case-insensitive filter when listing")
    args = ap.parse_args()

    families, raw = installed_families()
    lower = {f.lower(): f for f in families}
    lower_raw = {r.lower() for r in raw}

    if args.fonts:
        missing = []
        for want in args.fonts:
            key = want.lower()
            hit = key in lower or key in lower_raw or any(r.startswith(key) for r in lower_raw)
            print(f"{'OK      ' if hit else 'MISSING '} {want}")
            if not hit:
                missing.append(want)
        if missing:
            print("")
            print("Missing fonts fall back to Calibri/Arial in PowerPoint. Options:")
            print("  - install them (Google Fonts: https://fonts.google.com/ ; Fontshare: https://www.fontshare.com/)")
            print("  - or choose the preset's Office-safe fallback pair from STYLE_PRESETS.md")
            sys.exit(1)
        return

    names = sorted(families, key=str.lower)
    if args.grep:
        names = [n for n in names if args.grep.lower() in n.lower()]
    print(f"{len(names)} font famil{'y' if len(names) == 1 else 'ies'} installed on {platform.system()}")
    for n in names:
        print("  " + n)


if __name__ == "__main__":
    main()

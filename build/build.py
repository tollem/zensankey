#!/usr/bin/env python3
"""Builds the ZenSankey add-ins from the VBA sources.

Steps: ASCII check of the sources -> VBA project compilation with LibreOffice
(build/lo_vba_export.py) -> packaging (build/package.py) into dist/.

Usage: python3 build/build.py
Requires: LibreOffice (soffice), python3 with uno, openpyxl, python-pptx.
"""
import os, subprocess, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "src"); OUT = os.path.join(ROOT, "build", "out"); DIST = os.path.join(ROOT, "dist")
PROJECT = "ZenSankey"


def check_ascii(path):
    data = open(path, "rb").read()
    bad = [i for i, b in enumerate(data) if b > 127]
    if bad:
        sys.exit(f"{path}: non-ASCII byte at offset {bad[0]}. Sources must be pure ASCII (use ChrW() for other characters).")
    for n, line in enumerate(data.split(b"\n"), 1):
        if len(line) > 1000:
            sys.exit(f"{path}:{n}: line longer than 1000 characters")


def run(args):
    print("+", " ".join(args))
    subprocess.run(args, check=True)


def main():
    os.makedirs(OUT, exist_ok=True); os.makedirs(DIST, exist_ok=True)
    for m in ("SankeyCore", "SankeyExcel", "SankeyPowerPoint"):
        check_ascii(os.path.join(SRC, m + ".bas"))
    py = sys.executable
    exp = os.path.join(ROOT, "build", "lo_vba_export.py")
    run([py, exp, os.path.join(OUT, "vba_excel.bin"), PROJECT,
         f"SankeyCore={os.path.join(SRC, 'SankeyCore.bas')}", f"SankeyExcel={os.path.join(SRC, 'SankeyExcel.bas')}"])
    run([py, exp, os.path.join(OUT, "vba_powerpoint.bin"), PROJECT,
         f"SankeyCore={os.path.join(SRC, 'SankeyCore.bas')}", f"SankeyPowerPoint={os.path.join(SRC, 'SankeyPowerPoint.bas')}"])
    run([py, os.path.join(ROOT, "build", "package.py"), os.path.join(OUT, "vba_excel.bin"), os.path.join(OUT, "vba_powerpoint.bin"), DIST])
    print("done:", DIST)


if __name__ == "__main__":
    main()

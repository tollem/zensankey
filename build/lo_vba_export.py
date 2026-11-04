#!/usr/bin/env python3
"""Compiles VBA modules into a vbaProject.bin using LibreOffice (headless, UNO).

LibreOffice Calc can export a Basic library flagged as VBA-compatible to the
binary MS-OVBA format when saving as .xlsm. This script creates a hidden Calc
document, loads the given .bas files into a library, saves it as .xlsm and
extracts xl/vbaProject.bin. The resulting project is host-agnostic and is
embedded into both the Excel (.xlam) and PowerPoint (.ppam) packages.

Usage: lo_vba_export.py OUT.bin PROJECT_NAME Module1=path.bas [Module2=path2.bas ...]
Requires: soffice on PATH, Python with the uno module (python3-uno).
"""
import sys, os, subprocess, time, zipfile, tempfile
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.script import ModuleInfo


def pv(name, value):
    p = PropertyValue(); p.Name = name; p.Value = value; return p


def main():
    out = os.path.abspath(sys.argv[1])
    project = sys.argv[2]
    mods = [a.split("=", 1) for a in sys.argv[3:]]
    port = 2095
    proc = subprocess.Popen(["soffice", "--headless", "--invisible", "--nologo", "--norestore",
                             f"--accept=socket,host=127.0.0.1,port={port};urp;"],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    ctx = None
    for _ in range(60):
        try:
            local = uno.getComponentContext()
            resolver = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
            ctx = resolver.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext")
            break
        except Exception:
            time.sleep(1)
    if ctx is None:
        proc.kill(); sys.exit("LibreOffice did not start")
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (pv("Hidden", True),))
    libs = doc.BasicLibraries
    libs.createLibrary(project)
    lib = libs.getByName(project)
    for name, path in mods:
        src = open(path, encoding="ascii").read()
        lib.insertByName(name, src)
        info = ModuleInfo(); info.ModuleType = 1
        lib.insertModuleInfo(name, info)
        print(f"  module {name}: {len(src)} chars")
    libs.VBACompatibilityMode = True
    libs.ProjectName = project
    tmp = tempfile.mkdtemp()
    xlsm = os.path.join(tmp, "export.xlsm")
    doc.storeToURL(uno.systemPathToFileUrl(xlsm), (pv("FilterName", "Calc MS Excel 2007 VBA XML"), pv("Overwrite", True)))
    doc.close(True)
    try:
        desktop.terminate()
    except Exception:
        pass
    proc.wait(timeout=30)
    data = zipfile.ZipFile(xlsm).read("xl/vbaProject.bin")
    open(out, "wb").write(data)
    print(f"  vbaProject.bin: {len(data)} bytes -> {out}")


if __name__ == "__main__":
    main()

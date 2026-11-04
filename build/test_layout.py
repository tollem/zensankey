#!/usr/bin/env python3
"""Compiles the VBA sources in LibreOffice (VBA compatibility mode) and runs tests/SkTest.bas.
Checks syntax (as far as LibreOffice Basic goes) and the layout engine. Drawing calls
(BuildFreeform, text boxes) are not supported by LibreOffice and are not exercised here.

Usage: python3 build/test_layout.py
"""
import sys, os, subprocess, time
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.script import ModuleInfo

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def pv(n, v):
    p = PropertyValue(); p.Name = n; p.Value = v; return p


def main():
    mods = [("SankeyCore", "src/SankeyCore.bas"), ("SankeyExcel", "src/SankeyExcel.bas"),
            ("SankeyPowerPoint", "src/SankeyPowerPoint.bas"), ("SkTest", "tests/SkTest.bas")]
    port = 2090
    proc = subprocess.Popen(["soffice", "--headless", "--invisible", "--nologo", "--norestore", f"--accept=socket,host=127.0.0.1,port={port};urp;"],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    ctx = None
    for _ in range(60):
        try:
            local = uno.getComponentContext(); r = local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver", local)
            ctx = r.resolve(f"uno:socket,host=127.0.0.1,port={port};urp;StarOffice.ComponentContext"); break
        except Exception:
            time.sleep(1)
    desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0, (pv("Hidden", True),))
    libs = doc.BasicLibraries
    libs.createLibrary("ZenSankey"); lib = libs.getByName("ZenSankey")
    for name, path in mods:
        src = open(os.path.join(ROOT, path), encoding="ascii").read().replace("IRibbonControl", "Object")
        lib.insertByName(name, "Option VBASupport 1\n" + src)
        info = ModuleInfo(); info.ModuleType = 1; lib.insertModuleInfo(name, info)
    libs.VBACompatibilityMode = True
    sp = doc.getScriptProvider()
    code = 0
    try:
        script = sp.getScript("vnd.sun.star.script:ZenSankey.SkTest.SkTestRun?language=Basic&location=document")
        out = script.invoke((), (), ())[0]
        print(out)
    except Exception as e:
        print("ERROR:", e); code = 1
    doc.close(True)
    try:
        desktop.terminate()
    except Exception:
        pass
    proc.wait(timeout=30)
    sys.exit(code)


if __name__ == "__main__":
    main()

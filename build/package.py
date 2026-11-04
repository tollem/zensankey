#!/usr/bin/env python3
"""Packages the add-ins: ZenSankey.xlam (Excel) and ZenSankey.ppam (PowerPoint),
plus macro-enabled test documents (.xlsm / .pptm) that need no installation.

Usage: package.py VBA_EXCEL.bin VBA_PPT.bin OUT_DIR
"""
import sys, zipfile, re, io, os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from pptx import Presentation

VBA_REL = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
UI_REL = "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
NAME = "ZenSankey"


def custom_ui(host):
    tip = "Draw a Sankey diagram (native shapes) from the selected range: Source | Target | Amount"
    if host == "ppt":
        tip = "Draw a Sankey diagram (native shapes) from the selected table or text box"
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabZenSankey" label="ZenSankey">
        <group id="grpZenSankey" label="ZenSankey">
          <button id="btnSankeyDraw" label="Draw Sankey" imageMso="ShapesInsertGallery" size="large" onAction="Ribbon_Draw" screentip="{tip}"/>
          <button id="btnSankeyExample" label="Example" imageMso="TableInsert" size="large" onAction="Ribbon_Example" screentip="Insert an example flow table"/>
          <button id="btnSankeyOptions" label="Options" imageMso="PropertySheet" size="large" onAction="Ribbon_Options" screentip="Add the options list (palette, links, labels, titles, size)"/>
          <button id="btnSankeyHelp" label="Help" imageMso="Help" size="large" onAction="Ribbon_Help" screentip="How to use ZenSankey"/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
'''


def rewrite_zip(src_bytes, edits, adds):
    zin = zipfile.ZipFile(io.BytesIO(src_bytes))
    out = io.BytesIO(); zout = zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED)
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename in edits:
            data = edits[item.filename](data.decode("utf-8")).encode("utf-8")
        zout.writestr(item, data)
    for name, data in adds.items():
        zout.writestr(name, data)
    zout.close()
    return out.getvalue()


def add_rel(xml, rid, rtype, target):
    return xml.replace("</Relationships>", f'<Relationship Id="{rid}" Type="{rtype}" Target="{target}"/></Relationships>')


def excel_base():
    wb = Workbook(); ws = wb.active; ws.title = NAME
    ws.column_dimensions["A"].width = 110
    lines = [f"{NAME} for Excel", "",
             "This workbook is an add-in (.xlam). Once installed, a Sankey tab appears in the ribbon.",
             "1. Prepare a table Source | Target | Amount (optional columns: Color #RRGGBB, Source column, Target column).",
             "2. Select the range and click Draw Sankey. The diagram is a group of native shapes.",
             "3. Click Options to add a Sankey_Options sheet and customise palette, links, labels, titles and size.",
             "Documentation and updates: see the README of the project."]
    for i, t in enumerate(lines, 1):
        c = ws.cell(row=i, column=1, value=t)
        c.font = Font(name="Arial", size=11 if i > 1 else 14, bold=(i == 1)); c.alignment = Alignment(wrap_text=True)
    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()


def ct_xl(x, addin):
    ct = "application/vnd.ms-excel.addin.macroEnabled.main+xml" if addin else "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
    x = x.replace("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml", ct)
    return x.replace("</Types>", '<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>')


def wbxml(x):
    if "codeName=" in x:
        return x
    if "<workbookPr" in x:
        return re.sub(r"<workbookPr", '<workbookPr codeName="ThisWorkbook"', x, count=1)
    return re.sub(r"(<workbook[^>]*>)", r'\1<workbookPr codeName="ThisWorkbook"/>', x, count=1)


def sheet_cn(x):
    if "<sheetPr" in x:
        return re.sub(r"<sheetPr", '<sheetPr codeName="Sheet1"', x, count=1) if "codeName" not in x else x
    return re.sub(r"(<worksheet[^>]*>)", r'\1<sheetPr codeName="Sheet1"/>', x, count=1)


def ct_pp(x, addin):
    ct = "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml" if addin else "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
    x = x.replace("application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml", ct)
    return x.replace("</Types>", '<Override PartName="/ppt/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>')


def main():
    vba_xl, vba_pp, out_dir = sys.argv[1], sys.argv[2], sys.argv[3]
    xlsx = excel_base()
    prs = Presentation(); prs.slide_width = 12192000; prs.slide_height = 6858000
    buf = io.BytesIO(); prs.save(buf); pptx = buf.getvalue()
    xl_bin = open(vba_xl, "rb").read(); pp_bin = open(vba_pp, "rb").read()
    for addin, ext in ((True, "xlam"), (False, "xlsm")):
        data = rewrite_zip(xlsx, {
            "[Content_Types].xml": lambda x, a=addin: ct_xl(x, a),
            "xl/workbook.xml": wbxml,
            "xl/worksheets/sheet1.xml": sheet_cn,
            "xl/_rels/workbook.xml.rels": lambda x: add_rel(x, "rIdVBA", VBA_REL, "vbaProject.bin"),
            "_rels/.rels": lambda x: add_rel(x, "rIdUI", UI_REL, "customUI/customUI14.xml"),
        }, {"xl/vbaProject.bin": xl_bin, "customUI/customUI14.xml": custom_ui("xl").encode("utf-8")})
        fn = f"{NAME}.{ext}" if addin else f"{NAME}-test.{ext}"
        open(os.path.join(out_dir, fn), "wb").write(data); print("  wrote", fn)
    for addin, ext in ((True, "ppam"), (False, "pptm")):
        data = rewrite_zip(pptx, {
            "[Content_Types].xml": lambda x, a=addin: ct_pp(x, a),
            "ppt/_rels/presentation.xml.rels": lambda x: add_rel(x, "rIdVBA", VBA_REL, "vbaProject.bin"),
            "_rels/.rels": lambda x: add_rel(x, "rIdUI", UI_REL, "customUI/customUI14.xml"),
        }, {"ppt/vbaProject.bin": pp_bin, "customUI/customUI14.xml": custom_ui("ppt").encode("utf-8")})
        fn = f"{NAME}.{ext}" if addin else f"{NAME}-test.{ext}"
        open(os.path.join(out_dir, fn), "wb").write(data); print("  wrote", fn)


if __name__ == "__main__":
    main()

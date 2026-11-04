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


def dropdown(cid, label, items, tip, size_string=None):
    ss = f' sizeString="{size_string}"' if size_string else ""
    xml = f'          <dropDown id="{cid}" label="{label}"{ss} onAction="Ribbon_DropDown" getSelectedItemIndex="Ribbon_DropDownIndex" screentip="{tip}">\n'
    for i, it in enumerate(items):
        xml += f'            <item id="{cid}_{i}" label="{it}"/>\n'
    return xml + "          </dropDown>\n"


def checkbox(cid, label, tip):
    return f'          <checkBox id="{cid}" label="{label}" onAction="Ribbon_Check" getPressed="Ribbon_CheckState" screentip="{tip}"/>\n'


def editbox(cid, label, tip, size_string):
    return f'          <editBox id="{cid}" label="{label}" sizeString="{size_string}" onChange="Ribbon_EditChange" getText="Ribbon_EditText" screentip="{tip}"/>\n'


def custom_ui(host):
    tip = "Draw a Sankey diagram (native shapes) from the selected range: Source | Target | Amount. With a diagram selected: redraw it."
    if host == "ppt":
        tip = "Draw a Sankey diagram (native shapes) from the selected table or text box. With a diagram selected: redraw it."
    live = "Applies immediately to the selected diagram, and to the next diagrams drawn."
    x = '<?xml version="1.0" encoding="UTF-8"?>\n'
    x += '<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="Ribbon_OnLoad">\n  <ribbon>\n    <tabs>\n'
    x += '      <tab id="tabZenSankey" label="ZenSankey">\n'
    x += '        <group id="grpDiagram" label="Diagram">\n'
    x += f'          <button id="btnDraw" label="Draw Sankey" imageMso="ShapesInsertGallery" size="large" onAction="Ribbon_Draw" screentip="{tip}"/>\n'
    x += '          <button id="btnUpdate" label="Update" imageMso="Refresh" size="large" onAction="Ribbon_Update" screentip="Redraw the selected diagram from its source table, with the current style"/>\n'
    x += '          <button id="btnExample" label="Example" imageMso="TableInsert" size="large" onAction="Ribbon_Example" screentip="Insert an example flow table and draw it"/>\n'
    x += '        </group>\n'
    x += '        <group id="grpStyle" label="Style">\n'
    x += dropdown("ddPalette", "Palette", ["Teal", "Vivid", "Pastel", "Muted", "Ocean", "Warm", "Mono"], f"Color palette. {live}", "Pastel  ")
    x += dropdown("ddColorBy", "Colors", ["Follow the flow", "One per node", "One per column"],
                  f"Follow the flow: a node keeps the color of its main source, branches get new colors. {live}", "Follow the flow ")
    x += dropdown("ddLinkColor", "Links", ["Target color", "Source color", "Gradient", "Gray"], f"Color of the flow bands. {live}", "Target color ")
    x += dropdown("ddOpacity", "Opacity", ["25 %", "35 %", "45 %", "55 %", "70 %", "85 %", "100 %"], f"Opacity of the flow bands. {live}", "100 %  ")
    x += dropdown("ddShape", "Shape", ["Curved", "Soft curve", "Straight"], f"Shape of the flow bands. {live}", "Soft curve ")
    x += dropdown("ddNodeWidth", "Nodes", ["Thin", "Normal", "Wide", "Extra wide"], f"Width of the node bars. {live}", "Extra wide ")
    x += checkbox("cbRounded", "Rounded nodes", f"Node bars with rounded corners. {live}")
    x += checkbox("cbOutline", "Band outline", f"Thin white outline around the flow bands. {live}")
    x += '        </group>\n'
    x += '        <group id="grpLabels" label="Labels">\n'
    x += dropdown("ddLabels", "Show", ["Name and value", "Name only", "Value only", "No labels"], f"Content of the node labels. {live}", "Name and value ")
    x += dropdown("ddFontSize", "Size", ["8", "9", "10", "11", "12", "14", "16", "18"], f"Font size of the labels (titles scale with it). {live}", "18  ")
    x += dropdown("ddFont", "Font", ["Arial", "Calibri", "Helvetica", "Verdana", "Georgia", "Times New Roman"], f"Font of the labels and titles. {live}", "Times New Roman ")
    x += dropdown("ddFormat", "Format", ["Standard", "Thousands (k)", "Millions (M)", "Percent"], f"Number format of the values. {live}", "Thousands (k) ")
    x += dropdown("ddDecimals", "Decimals", ["0", "1", "2"], f"Decimals shown in the values. {live}", "0  ")
    x += editbox("ebUnit", "Unit", f"Unit appended to the values, e.g. CHF or kWh. Press Enter to apply. {live}", "CHF  ")
    x += checkbox("cbOutside", "Outside labels", f"Labels of the first and last columns outside the diagram. {live}")
    x += checkbox("cbLinkValues", "Values on bands", f"Write the amount on bands that are tall enough. {live}")
    x += checkbox("cbLabelBg", "Label background", f"Translucent white box behind each label (readability over the bands). {live}")
    x += '        </group>\n'
    x += '        <group id="grpTitles" label="Titles">\n'
    x += editbox("ebTitle", "Title", f"Title above the diagram. Press Enter to apply. {live}", "WWWWWWWWWWWWWWWWWW")
    x += editbox("ebSubtitle", "Subtitle", f"Subtitle in smaller gray italics. Press Enter to apply. {live}", "WWWWWWWWWWWWWWWWWW")
    x += editbox("ebColumns", "Sections", f"Section titles above the columns, separated by ; (e.g. Sources;Allocation;Result). Press Enter to apply. {live}", "WWWWWWWWWWWWWWWWWW")
    x += checkbox("cbBackground", "Background", f"Light rounded background behind the diagram. {live}")
    x += '        </group>\n'
    x += '        <group id="grpMore" label="More">\n'
    x += dropdown("ddOrder", "Order", ["Automatic", "As in table"], f"Vertical order of the nodes: automatic (fewer crossings) or the order of the table. {live}", "As in table ")
    x += '          <button id="btnReset" label="Reset style" imageMso="ClearFormatting" onAction="Ribbon_Reset" screentip="Restore the default style (and apply it to the selected diagram)"/>\n'
    x += '          <button id="btnHelp" label="Help" imageMso="Help" onAction="Ribbon_Help" screentip="How to use ZenSankey"/>\n'
    x += '        </group>\n'
    x += '      </tab>\n    </tabs>\n  </ribbon>\n</customUI>\n'
    return x


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
             "This workbook is an add-in (.xlam). Once installed, a ZenSankey tab appears in the ribbon.",
             "1. Prepare a table Source | Target | Amount (optional columns: Color #RRGGBB, Source column, Target column).",
             "2. Select the range and click Draw Sankey. The diagram is a group of native shapes.",
             "3. Adjust the style with the ribbon controls (palette, links, labels, titles): the selected diagram updates immediately.",
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

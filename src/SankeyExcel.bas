Option Explicit
' =====================================================================
'  ZenSankey - native-shape Sankey diagrams for Excel and PowerPoint
'  SankeyExcel: Excel entry points (ribbon callbacks, selection reading,
'  options sheet, example data).
'
'  Copyright (c) 2026 Samuel Mellot
'  Licensed under the Apache License, Version 2.0. See LICENSE and NOTICE.
' =====================================================================

Private Const OPTIONS_SHEET As String = "Sankey_Options"

' Draws a Sankey from the selected range: Source | Target | Amount [| Color | Source column | Target column]
Public Sub DrawSankey()
    Dim rng As Object, r As Long, r0 As Long, nRows As Long, nCols As Long
    Dim src As String, dst As String, v As Double, colHex As String, cs As Long, cd As Long
    Dim ws As Object, grp As Object, L As Single, T As Single, W As Single, H As Single
    On Error GoTo Failed
    If TypeName(Application.Selection) <> "Range" Then
        MsgBox "Select the range that contains the flows first: three columns Source | Target | Amount (with or without a header row).", vbExclamation, "ZenSankey"
        Exit Sub
    End If
    Set rng = Application.Selection
    If rng.Columns.Count < 3 Then
        MsgBox "The selection must contain at least three columns: Source | Target | Amount.", vbExclamation, "ZenSankey"
        Exit Sub
    End If
    Set ws = rng.Parent
    nRows = rng.Rows.Count: nCols = rng.Columns.Count
    r0 = 1
    If Not IsNumeric(rng.Cells(1, 3).Value) Or Len(Trim$(CStr(rng.Cells(1, 3).Value))) = 0 Then r0 = 2
    SkReset
    For r = r0 To nRows
        src = Trim$(CStr(rng.Cells(r, 1).Value)): dst = Trim$(CStr(rng.Cells(r, 2).Value))
        v = ParseNumber(rng.Cells(r, 3).Value)
        colHex = "": cs = -1: cd = -1
        If nCols >= 4 Then If IsHexColor(CStr(rng.Cells(r, 4).Value)) Then colHex = CStr(rng.Cells(r, 4).Value)
        If nCols >= 5 Then If IsNumeric(rng.Cells(r, 5).Value) And Len(CStr(rng.Cells(r, 5).Value)) > 0 Then cs = CLng(rng.Cells(r, 5).Value)
        If nCols >= 6 Then If IsNumeric(rng.Cells(r, 6).Value) And Len(CStr(rng.Cells(r, 6).Value)) > 0 Then cd = CLng(rng.Cells(r, 6).Value)
        If Len(src) > 0 And Len(dst) > 0 And v > 0 Then SkAddLink src, dst, v, colHex, cs, cd
    Next r
    If SkLinkCount = 0 Then
        MsgBox "No valid flow found (empty cells or zero amounts).", vbExclamation, "ZenSankey"
        Exit Sub
    End If
    W = 640: H = 400
    ApplyOptionsSheet ws.Parent, W, H
    L = rng.Left + rng.Width + 30: T = rng.Top
    Set grp = SkDraw(ws.Shapes, L, T, W, H, "")
    grp.Select
    Exit Sub
Failed:
    MsgBox "The Sankey could not be drawn: " & Err.Description, vbCritical, "ZenSankey"
End Sub

' Reads the Sankey_Options sheet of the workbook (key in column A, value in column B) if it exists
Private Sub ApplyOptionsSheet(ByVal wb As Object, ByRef W As Single, ByRef H As Single)
    Dim ws As Object, r As Long, k As String, v As String, txt As String
    On Error Resume Next
    Set ws = wb.Worksheets(OPTIONS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    For r = 2 To 200
        k = Trim$(CStr(ws.Cells(r, 1).Value)): v = Trim$(CStr(ws.Cells(r, 2).Value))
        If Len(k) > 0 Then
            Select Case LCase$(k)
                Case "width": If ParseNumber(v) > 50 Then W = ParseNumber(v)
                Case "height": If ParseNumber(v) > 50 Then H = ParseNumber(v)
                Case Else: txt = txt & k & " = " & v & vbCr
            End Select
        End If
    Next r
    SkApplyOptionText txt
End Sub

' Creates (or activates) the Sankey_Options sheet with every option, its default value and a description
Public Sub InsertOptionsSheet()
    Dim wb As Object, ws As Object, lst As Variant, i As Long, parts() As String, r As Long
    On Error GoTo Failed
    Set wb = Application.ActiveWorkbook
    If wb Is Nothing Then MsgBox "Open a workbook first.", vbExclamation, "ZenSankey": Exit Sub
    On Error Resume Next
    Set ws = wb.Worksheets(OPTIONS_SHEET)
    On Error GoTo Failed
    If Not ws Is Nothing Then
        ws.Activate
        MsgBox "The sheet " & OPTIONS_SHEET & " already exists. Edit the values in column B; they apply to the next Sankey drawn in this workbook.", vbInformation, "ZenSankey"
        Exit Sub
    End If
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = OPTIONS_SHEET
    ws.Cells(1, 1).Value = "Option": ws.Cells(1, 2).Value = "Value": ws.Cells(1, 3).Value = "Description / choices"
    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A1:C1").Interior.Color = RGB(2, 128, 144)
    ws.Range("A1:C1").Font.Color = RGB(255, 255, 255)
    r = 2
    ws.Cells(r, 1).Value = "width": ws.Cells(r, 2).Value = 640: ws.Cells(r, 3).Value = "Diagram width in points (Excel only)": r = r + 1
    ws.Cells(r, 1).Value = "height": ws.Cells(r, 2).Value = 400: ws.Cells(r, 3).Value = "Diagram height in points (Excel only)": r = r + 1
    lst = SkOptionsList()
    For i = LBound(lst) To UBound(lst)
        parts = Split(lst(i), "|")
        ws.Cells(r, 1).Value = parts(0)
        ws.Cells(r, 2).Value = parts(1)
        ws.Cells(r, 2).NumberFormat = "@"
        If Len(parts(3)) > 0 Then
            ws.Cells(r, 3).Value = parts(2) & " [" & Replace(parts(3), ",", ", ") & "]"
            On Error Resume Next
            ws.Cells(r, 2).Validation.Delete
            ws.Cells(r, 2).Validation.Add 3, 1, 1, parts(3)
            ws.Cells(r, 2).Validation.ShowError = False
            On Error GoTo Failed
        Else
            ws.Cells(r, 3).Value = parts(2)
        End If
        r = r + 1
    Next i
    ws.Range("B2:B" & r).Font.Color = RGB(0, 0, 255)
    ws.Range("B2:B" & r).Interior.Color = RGB(255, 255, 153)
    ws.Columns(1).ColumnWidth = 24: ws.Columns(2).ColumnWidth = 28: ws.Columns(3).ColumnWidth = 95
    ws.Cells(r + 1, 1).Value = "Edit column B. Keys in column A must not change. Values apply to the next Sankey drawn in this workbook."
    ws.Cells(r + 1, 1).Font.Italic = True
    ws.Activate
    ws.Cells(2, 2).Select
    Exit Sub
Failed:
    MsgBox "The options sheet could not be created: " & Err.Description, vbCritical, "ZenSankey"
End Sub

' Inserts an example flow table at the active cell (income statement)
Public Sub InsertExample()
    Dim c As Object, i As Long, d As Variant
    On Error GoTo Failed
    Set c = Application.ActiveCell
    d = Array( _
        Array("Source", "Target", "Amount", "Color (optional)", "Source col. (opt.)", "Target col. (opt.)"), _
        Array("Revenue", "Cost of sales", 450, "#E34948", "", ""), _
        Array("Revenue", "Gross profit", 550, "#1BAF7A", "", ""), _
        Array("Gross profit", "Operating expenses", 300, "#E34948", "", ""), _
        Array("Gross profit", "Operating profit", 250, "#1BAF7A", "", ""), _
        Array("Operating expenses", "R&D", 120, "#E34948", "", ""), _
        Array("Operating expenses", "Sales & marketing", 110, "#E34948", "", ""), _
        Array("Operating expenses", "General & admin", 70, "#E34948", "", ""), _
        Array("Operating profit", "Taxes", 60, "#E34948", "", ""), _
        Array("Operating profit", "Net profit", 190, "#1BAF7A", "", ""))
    For i = 0 To UBound(d)
        c.Offset(i, 0).Resize(1, 6).Value = d(i)
    Next i
    c.Resize(1, 6).Font.Bold = True
    c.Parent.Columns(c.Column).ColumnWidth = 24
    c.Parent.Columns(c.Column + 1).ColumnWidth = 24
    c.Resize(UBound(d) + 1, 4).Select
    MsgBox "Example inserted and selected: click Draw Sankey. Columns 4 to 6 are optional (color #RRGGBB, forced column of the source and of the target, 0 = first column).", vbInformation, "ZenSankey"
    Exit Sub
Failed:
    MsgBox "The example could not be inserted: " & Err.Description, vbCritical, "ZenSankey"
End Sub

Public Sub ShowHelp()
    MsgBox "ZenSankey for Excel" & vbCrLf & vbCrLf & _
           "1. Prepare a table: Source | Target | Amount (one row per flow). Optional columns: Color (#RRGGBB), Source column, Target column (0 = first column)." & vbCrLf & _
           "2. Select the range (with or without header) and click Draw Sankey. The diagram is placed to the right of the selection." & vbCrLf & _
           "3. Click Options to add a Sankey_Options sheet: palette, colors, link shape, labels, titles, section titles, value format, size..." & vbCrLf & _
           "4. The diagram is a group of native shapes: resize it, edit colors and texts, copy and paste it into PowerPoint." & vbCrLf & vbCrLf & _
           "Amounts such as 3 000, 3'000 or 2 650,5 are accepted. Documentation: see the README of the project.", vbInformation, "ZenSankey"
End Sub

' ---- Ribbon callbacks --------------------------------------------------
Public Sub Ribbon_Draw(control As IRibbonControl)
    DrawSankey
End Sub

Public Sub Ribbon_Example(control As IRibbonControl)
    InsertExample
End Sub

Public Sub Ribbon_Options(control As IRibbonControl)
    InsertOptionsSheet
End Sub

Public Sub Ribbon_Help(control As IRibbonControl)
    ShowHelp
End Sub

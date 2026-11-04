Option Explicit
' =====================================================================
'  ZenSankey - native-shape Sankey diagrams for Excel and PowerPoint
'  SankeyExcel: Excel entry points (ribbon callbacks, range reading,
'  in-place update of a diagram, example data).
'
'  Copyright (c) 2026 Samuel Mellot
'  Licensed under the Apache License, Version 2.0. See LICENSE and NOTICE.
' =====================================================================

Private Const META_TAG As String = "ZenSankey"
Public gRibbon As IRibbonUI

' Draws a Sankey from the selected range: Source | Target | Amount [| Color | Source column | Target column].
' If a ZenSankey diagram is selected instead, it is redrawn with the current style.
Public Sub DrawSankey()
    Dim rng As Object, ws As Object, grp As Object, L As Single, T As Single, W As Single, H As Single
    On Error GoTo Failed
    Set grp = SelectedDiagram()
    If Not grp Is Nothing Then UpdateDiagram grp: Exit Sub
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
    SkReset
    ReadRange rng
    If SkLinkCount = 0 Then
        MsgBox "No valid flow found (empty cells or zero amounts).", vbExclamation, "ZenSankey"
        Exit Sub
    End If
    W = 620: H = 360
    L = rng.Left + rng.Width + 30: T = rng.Top
    Set grp = SkDraw(ws.Shapes, L, T, W, H, "", True)
    grp.AlternativeText = MetaText(ws.Name, rng.Address(False, False), grp)
    grp.Select
    Exit Sub
Failed:
    MsgBox "The Sankey could not be drawn: " & Err.Description, vbCritical, "ZenSankey"
End Sub

' Reads the flows of a range into the engine
Private Sub ReadRange(ByVal rng As Object)
    Dim r As Long, r0 As Long, nRows As Long, nCols As Long
    Dim src As String, dst As String, v As Double, colHex As String, cs As Long, cd As Long
    nRows = rng.Rows.Count: nCols = rng.Columns.Count
    r0 = 1
    If Not IsNumeric(rng.Cells(1, 3).Value) Or Len(Trim$(CStr(rng.Cells(1, 3).Value))) = 0 Then r0 = 2
    For r = r0 To nRows
        src = Trim$(CStr(rng.Cells(r, 1).Value)): dst = Trim$(CStr(rng.Cells(r, 2).Value))
        v = ParseNumber(rng.Cells(r, 3).Value)
        colHex = "": cs = -1: cd = -1
        If nCols >= 4 Then If IsHexColor(CStr(rng.Cells(r, 4).Value)) Then colHex = CStr(rng.Cells(r, 4).Value)
        If nCols >= 5 Then If IsNumeric(rng.Cells(r, 5).Value) And Len(CStr(rng.Cells(r, 5).Value)) > 0 Then cs = CLng(rng.Cells(r, 5).Value)
        If nCols >= 6 Then If IsNumeric(rng.Cells(r, 6).Value) And Len(CStr(rng.Cells(r, 6).Value)) > 0 Then cd = CLng(rng.Cells(r, 6).Value)
        If Len(src) > 0 And Len(dst) > 0 And v > 0 Then SkAddLink src, dst, v, colHex, cs, cd
    Next r
End Sub

' Metadata stored with the diagram: source range, size and style (so that it can be redrawn)
Private Function MetaText(ByVal sheetName As String, ByVal addr As String, ByVal grp As Object) As String
    MetaText = META_TAG & " 2.1" & vbLf & "sheet=" & sheetName & vbLf & "range=" & addr & vbLf & _
               "size=" & Format$(SkDrawnWidth(), "0.0") & ";" & Format$(SkDrawnHeight(), "0.0") & vbLf & _
               "box=" & Format$(grp.Width, "0.0") & ";" & Format$(grp.Height, "0.0") & vbLf & _
               "style:" & vbLf & SkSettingsText()
End Function

Private Function MetaValue(ByVal meta As String, ByVal key As String) As String
    Dim rows() As String, i As Long
    meta = Replace(meta, vbCrLf, vbLf): meta = Replace(meta, vbCr, vbLf)
    rows = Split(meta, vbLf)
    For i = 0 To UBound(rows)
        If Left$(rows(i), Len(key) + 1) = key & "=" Then MetaValue = Mid$(rows(i), Len(key) + 2): Exit Function
        If rows(i) = "style:" Then Exit Function
    Next i
End Function

Private Function MetaStyle(ByVal meta As String) As String
    Dim p As Long
    meta = Replace(meta, vbCrLf, vbLf): meta = Replace(meta, vbCr, vbLf)
    p = InStr(meta, "style:" & vbLf)
    If p > 0 Then MetaStyle = Mid$(meta, p + 7)
End Function

' The ZenSankey group currently selected (or containing the selected shape), or Nothing
Public Function SelectedDiagram() As Object
    Dim sr As Object, s As Object, p As Object, k As Long
    On Error Resume Next
    Set sr = Application.Selection.ShapeRange
    If sr Is Nothing Then Exit Function
    Set s = sr(1)
    For k = 1 To 6
        If s Is Nothing Then Exit Function
        If Left$(s.AlternativeText, Len(META_TAG)) = META_TAG Then Set SelectedDiagram = s: Exit Function
        Set p = Nothing
        Set p = s.ParentGroup
        If p Is Nothing Then Exit Function
        Set s = p
    Next k
End Function

' Redraws a diagram from its source range, at the same place, with the current style
Public Sub UpdateDiagram(ByVal grp As Object)
    Dim ws As Object, wsSrc As Object, rng As Object, meta As String, sheetName As String, addr As String
    Dim L As Single, T As Single, W As Single, H As Single, parts() As String, newGrp As Object, sameBox As Boolean
    On Error GoTo Failed
    meta = grp.AlternativeText
    sheetName = MetaValue(meta, "sheet"): addr = MetaValue(meta, "range")
    Set ws = grp.Parent
    Set wsSrc = Nothing
    On Error Resume Next
    Set wsSrc = ws.Parent.Worksheets(sheetName)
    If wsSrc Is Nothing Then Set wsSrc = ws
    Set rng = wsSrc.Range(addr)
    On Error GoTo Failed
    If rng Is Nothing Then
        MsgBox "The source table of this diagram was not found (" & sheetName & "!" & addr & "). Select the table and click Draw Sankey.", vbExclamation, "ZenSankey"
        Exit Sub
    End If
    SkReset
    ReadRange rng
    If SkLinkCount = 0 Then
        MsgBox "No valid flow found in " & sheetName & "!" & addr & ".", vbExclamation, "ZenSankey"
        Exit Sub
    End If
    L = grp.Left: T = grp.Top
    W = grp.Width + 16: H = grp.Height + 16
    sameBox = False
    parts = Split(MetaValue(meta, "box"), ";")
    If UBound(parts) = 1 Then
        If Abs(ParseNumber(parts(0)) - grp.Width) < 1 And Abs(ParseNumber(parts(1)) - grp.Height) < 1 Then sameBox = True
    End If
    If sameBox Then
        parts = Split(MetaValue(meta, "size"), ";")
        If UBound(parts) = 1 Then W = ParseNumber(parts(0)): H = ParseNumber(parts(1))
    End If
    grp.Delete
    Set newGrp = SkDraw(ws.Shapes, L, T, W, H, "", sameBox)
    newGrp.AlternativeText = MetaText(wsSrc.Name, addr, newGrp)
    newGrp.Select
    Exit Sub
Failed:
    MsgBox "The diagram could not be updated: " & Err.Description, vbCritical, "ZenSankey"
End Sub

' Loads the style stored with the selected diagram (if any) into the engine
Private Function LoadSelectedStyle() As Object
    Dim grp As Object, st As String
    Set grp = SelectedDiagram()
    If grp Is Nothing Then Exit Function
    st = MetaStyle(grp.AlternativeText)
    If Len(st) > 0 Then SkApplyOptionText st
    Set LoadSelectedStyle = grp
End Function

Private Sub AfterStyleChange(ByVal grp As Object)
    If Not grp Is Nothing Then UpdateDiagram grp
    RefreshRibbon
End Sub

Private Sub RefreshRibbon()
    On Error Resume Next
    If Not gRibbon Is Nothing Then gRibbon.Invalidate
End Sub

' Inserts an example flow table at the active cell (income statement) and draws it
Public Sub InsertExample()
    Dim c As Object, i As Long, d As Variant
    On Error GoTo Failed
    Set c = Application.ActiveCell
    d = Array( _
        Array("Source", "Target", "Amount", "Color (optional)"), _
        Array("Revenue", "Cost of sales", 450, "#E34948"), _
        Array("Revenue", "Gross profit", 550, "#1BAF7A"), _
        Array("Gross profit", "Operating expenses", 300, "#E34948"), _
        Array("Gross profit", "Operating profit", 250, "#1BAF7A"), _
        Array("Operating expenses", "R&D", 120, "#E34948"), _
        Array("Operating expenses", "Sales & marketing", 110, "#E34948"), _
        Array("Operating expenses", "General & admin", 70, "#E34948"), _
        Array("Operating profit", "Taxes", 60, "#E34948"), _
        Array("Operating profit", "Net profit", 190, "#1BAF7A"))
    For i = 0 To UBound(d)
        c.Offset(i, 0).Resize(1, 4).Value = d(i)
    Next i
    c.Resize(1, 4).Font.Bold = True
    c.Parent.Columns(c.Column).ColumnWidth = 22
    c.Parent.Columns(c.Column + 1).ColumnWidth = 22
    c.Resize(UBound(d) + 1, 4).Select
    DrawSankey
    Exit Sub
Failed:
    MsgBox "The example could not be inserted: " & Err.Description, vbCritical, "ZenSankey"
End Sub

Public Sub ShowHelp()
    MsgBox "ZenSankey for Excel" & vbCrLf & vbCrLf & _
           "1. Prepare a table: Source | Target | Amount (one row per flow). Optional columns: Color (#RRGGBB), Source column, Target column (0 = first column)." & vbCrLf & _
           "2. Select the range (with or without header) and click Draw Sankey. The diagram is placed to the right of the selection." & vbCrLf & _
           "3. Adjust the style with the ribbon controls (palette, links, labels, titles...): a selected diagram is updated immediately." & vbCrLf & _
           "4. Update redraws the selected diagram after the table changed. Reset style restores the defaults." & vbCrLf & _
           "5. The diagram is a group of native shapes: resize it, edit colors and texts, copy and paste it into PowerPoint." & vbCrLf & vbCrLf & _
           "Amounts such as 3 000, 3'000 or 2 650,5 are accepted. Documentation: see the README of the project.", vbInformation, "ZenSankey"
End Sub

' ---- Ribbon callbacks --------------------------------------------------
Public Sub Ribbon_OnLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
End Sub

Public Sub Ribbon_Draw(control As IRibbonControl)
    DrawSankey
End Sub

Public Sub Ribbon_Update(control As IRibbonControl)
    Dim grp As Object
    Set grp = SelectedDiagram()
    If grp Is Nothing Then
        MsgBox "Select a ZenSankey diagram first (click on it), then Update.", vbExclamation, "ZenSankey"
    Else
        UpdateDiagram grp
    End If
End Sub

Public Sub Ribbon_Example(control As IRibbonControl)
    InsertExample
End Sub

Public Sub Ribbon_Help(control As IRibbonControl)
    ShowHelp
End Sub

Public Sub Ribbon_Reset(control As IRibbonControl)
    Dim grp As Object
    Set grp = SelectedDiagram()
    SkResetStyle
    AfterStyleChange grp
End Sub

Public Sub Ribbon_DropDown(control As IRibbonControl, id As String, index As Integer)
    Dim grp As Object
    Set grp = LoadSelectedStyle()
    SkRibbonSelect control.Id, index
    AfterStyleChange grp
End Sub

Public Sub Ribbon_DropDownIndex(control As IRibbonControl, ByRef index As Variant)
    index = SkRibbonIndex(control.Id)
End Sub

Public Sub Ribbon_Check(control As IRibbonControl, pressed As Boolean)
    Dim grp As Object
    Set grp = LoadSelectedStyle()
    SkRibbonSetPressed control.Id, pressed
    AfterStyleChange grp
End Sub

Public Sub Ribbon_CheckState(control As IRibbonControl, ByRef returnValue As Variant)
    returnValue = SkRibbonPressed(control.Id)
End Sub

Public Sub Ribbon_EditChange(control As IRibbonControl, text As String)
    Dim grp As Object
    Set grp = LoadSelectedStyle()
    SkRibbonSetText control.Id, text
    AfterStyleChange grp
End Sub

Public Sub Ribbon_EditText(control As IRibbonControl, ByRef returnedVal As Variant)
    returnedVal = SkRibbonText(control.Id)
End Sub

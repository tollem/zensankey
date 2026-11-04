Option Explicit
' =====================================================================
'  ZenSankey - native-shape Sankey diagrams for Excel and PowerPoint
'  SankeyPowerPoint: PowerPoint entry points (ribbon callbacks, table or
'  text reading, in-place update of a diagram, example table).
'
'  Copyright (c) 2026 Samuel Mellot
'  Licensed under the Apache License, Version 2.0. See LICENSE and NOTICE.
' =====================================================================

Private Const PP_SELECTION_SHAPES As Long = 2
Private Const PP_SELECTION_TEXT As Long = 3
Private Const META_TAG As String = "ZenSankey"
Public gRibbon As IRibbonUI

' Draws a Sankey from the selected table (Source | Target | Amount ...) or text box ("Source [amount] Target" per line).
' If a ZenSankey diagram is selected instead, it is redrawn with the current style.
Public Sub DrawSankey()
    Dim sel As Object, shp As Object, sld As Object, grp As Object
    Dim L As Single, T As Single, W As Single, H As Single, slW As Single, slH As Single
    On Error GoTo Failed
    Set grp = SelectedDiagram()
    If Not grp Is Nothing Then UpdateDiagram grp: Exit Sub
    Set sel = Application.ActiveWindow.Selection
    If sel.Type <> PP_SELECTION_SHAPES And sel.Type <> PP_SELECTION_TEXT Then
        MsgBox "Select a table (Source | Target | Amount) or a text box with one flow per line (Source [amount] Target) first.", vbExclamation, "ZenSankey"
        Exit Sub
    End If
    Set shp = sel.ShapeRange(1)
    Set sld = shp.Parent
    SkReset
    ReadShape shp
    If SkLinkCount = 0 Then
        MsgBox "No valid flow found. Expected: a table with three columns Source | Target | Amount, or text lines such as Source [amount] Target.", vbExclamation, "ZenSankey"
        Exit Sub
    End If
    slW = Application.ActivePresentation.PageSetup.SlideWidth
    slH = Application.ActivePresentation.PageSetup.SlideHeight
    ' default placement: below the source shape when there is room, otherwise the content area
    L = 36: W = slW - 72
    If shp.Top + shp.Height + 40 < slH - 200 Then
        T = shp.Top + shp.Height + 24: H = slH - T - 40
    Else
        T = 90: H = slH - 130
    End If
    Set grp = SkDraw(sld.Shapes, L, T, W, H, "", False)
    grp.AlternativeText = MetaText(sld, shp.Name, grp)
    grp.Select
    Exit Sub
Failed:
    MsgBox "The Sankey could not be drawn: " & Err.Description, vbCritical, "ZenSankey"
End Sub

' Reads the flows of a table or text shape into the engine
Private Sub ReadShape(ByVal shp As Object)
    If shp.HasTable Then
        ReadTable shp.Table
    ElseIf shp.HasTextFrame Then
        ReadText shp.TextFrame.TextRange.Text
    End If
End Sub

' Metadata stored with the diagram: source shape, size and style (so that it can be redrawn)
Private Function MetaText(ByVal sld As Object, ByVal srcName As String, ByVal grp As Object) As String
    MetaText = META_TAG & " 2.1" & vbLf & "slide=" & sld.SlideID & vbLf & "shape=" & srcName & vbLf & _
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
    Dim sel As Object, s As Object, p As Object, k As Long
    On Error Resume Next
    Set sel = Application.ActiveWindow.Selection
    If sel Is Nothing Then Exit Function
    If sel.Type <> PP_SELECTION_SHAPES Then Exit Function
    Set s = sel.ShapeRange(1)
    For k = 1 To 6
        If s Is Nothing Then Exit Function
        If Left$(s.AlternativeText, Len(META_TAG)) = META_TAG Then Set SelectedDiagram = s: Exit Function
        Set p = Nothing
        Set p = s.ParentGroup
        If p Is Nothing Then Exit Function
        Set s = p
    Next k
End Function

' Redraws a diagram from its source table or text box, at the same place, with the current style
Public Sub UpdateDiagram(ByVal grp As Object)
    Dim sld As Object, src As Object, meta As String, srcName As String
    Dim L As Single, T As Single, W As Single, H As Single, parts() As String, newGrp As Object, sameBox As Boolean
    On Error GoTo Failed
    meta = grp.AlternativeText
    srcName = MetaValue(meta, "shape")
    Set sld = grp.Parent
    Set src = Nothing
    On Error Resume Next
    Set src = sld.Shapes(srcName)
    On Error GoTo Failed
    If src Is Nothing Then
        MsgBox "The source table of this diagram (" & srcName & ") was not found on the slide. Select the table and click Draw Sankey.", vbExclamation, "ZenSankey"
        Exit Sub
    End If
    SkReset
    ReadShape src
    If SkLinkCount = 0 Then
        MsgBox "No valid flow found in " & srcName & ".", vbExclamation, "ZenSankey"
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
    Set newGrp = SkDraw(sld.Shapes, L, T, W, H, "", False)
    newGrp.AlternativeText = MetaText(sld, srcName, newGrp)
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

' Inserts an example table on the current slide and draws it
Public Sub InsertExample()
    Dim sld As Object, tb As Object, d As Variant, i As Long, j As Long, slW As Single
    On Error GoTo Failed
    Set sld = Application.ActiveWindow.View.Slide
    slW = Application.ActivePresentation.PageSetup.SlideWidth
    d = Array( _
        Array("Source", "Target", "Amount", "Color"), _
        Array("Revenue", "Cost of sales", "450", "#E34948"), _
        Array("Revenue", "Gross profit", "550", "#1BAF7A"), _
        Array("Gross profit", "Operating expenses", "300", "#E34948"), _
        Array("Gross profit", "Operating profit", "250", "#1BAF7A"), _
        Array("Operating expenses", "R&D", "120", "#E34948"), _
        Array("Operating expenses", "Sales & marketing", "110", "#E34948"), _
        Array("Operating expenses", "General & admin", "70", "#E34948"), _
        Array("Operating profit", "Taxes", "60", "#E34948"), _
        Array("Operating profit", "Net profit", "190", "#1BAF7A"))
    Set tb = sld.Shapes.AddTable(UBound(d) + 1, 4, 36, 40, slW / 2 - 50, 16 * (UBound(d) + 1))
    For i = 0 To UBound(d)
        For j = 0 To 3
            tb.Table.Cell(i + 1, j + 1).Shape.TextFrame.TextRange.Text = d(i)(j)
            tb.Table.Cell(i + 1, j + 1).Shape.TextFrame.TextRange.Font.Size = 9
        Next j
    Next i
    tb.Select
    DrawSankey
    Exit Sub
Failed:
    MsgBox "The example could not be inserted: " & Err.Description, vbCritical, "ZenSankey"
End Sub

Public Sub ShowHelp()
    MsgBox "ZenSankey for PowerPoint" & vbCrLf & vbCrLf & _
           "1. Insert a table Source | Target | Amount (Insert > Table, or paste Excel cells), or a text box with one flow per line: Source [amount] Target." & vbCrLf & _
           "2. Select the table or the text box, then click Draw Sankey." & vbCrLf & _
           "3. Adjust the style with the ribbon controls (palette, links, labels, titles...): a selected diagram is updated immediately." & vbCrLf & _
           "4. Update redraws the selected diagram after the table changed. Reset style restores the defaults." & vbCrLf & _
           "5. The diagram is a group of native shapes: move it, resize it, change colors and fonts. Keep the source table (it can be moved off the slide) if you want to update the diagram later." & vbCrLf & vbCrLf & _
           "Optional table columns: Color (#RRGGBB), Source column, Target column (0 = first column).", vbInformation, "ZenSankey"
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

' ---- Data reading --------------------------------------------------------
Private Sub ReadTable(ByVal tbl As Object)
    Dim r As Long, r0 As Long, nR As Long, nC As Long, src As String, dst As String, v As Double
    Dim colHex As String, cs As Long, cd As Long, t3 As String
    nR = tbl.Rows.Count: nC = tbl.Columns.Count
    If nC < 3 Then Exit Sub
    r0 = 1
    t3 = Trim$(CellText(tbl, 1, 3))
    If ParseNumber(t3) = 0 Then r0 = 2
    For r = r0 To nR
        src = Trim$(CellText(tbl, r, 1)): dst = Trim$(CellText(tbl, r, 2))
        v = ParseNumber(CellText(tbl, r, 3))
        colHex = "": cs = -1: cd = -1
        If nC >= 4 Then If IsHexColor(CellText(tbl, r, 4)) Then colHex = Trim$(CellText(tbl, r, 4))
        If nC >= 5 Then If Len(Trim$(CellText(tbl, r, 5))) > 0 Then cs = CLng(ParseNumber(CellText(tbl, r, 5)))
        If nC >= 6 Then If Len(Trim$(CellText(tbl, r, 6))) > 0 Then cd = CLng(ParseNumber(CellText(tbl, r, 6)))
        If Len(src) > 0 And Len(dst) > 0 And v > 0 Then SkAddLink src, dst, v, colHex, cs, cd
    Next r
End Sub

Private Function CellText(ByVal tbl As Object, ByVal r As Long, ByVal c As Long) As String
    On Error Resume Next
    CellText = tbl.Cell(r, c).Shape.TextFrame.TextRange.Text
End Function

' Free text: "Source [amount] Target" (SankeyMATIC syntax), "Source ; Target ; amount", tabs or commas
Private Sub ReadText(ByVal txt As String)
    Dim rows() As String, i As Long, ln As String, p1 As Long, p2 As Long, parts() As String, sep As String
    Dim src As String, dst As String, v As Double
    txt = Replace(txt, vbCrLf, vbCr): txt = Replace(txt, vbLf, vbCr): txt = Replace(txt, ChrW(11), vbCr)
    rows = Split(txt, vbCr)
    For i = 0 To UBound(rows)
        ln = Trim$(rows(i))
        If Len(ln) > 0 And Left$(ln, 1) <> "'" Then
            p1 = InStr(ln, "["): p2 = InStr(ln, "]")
            src = "": dst = "": v = 0
            If p1 > 1 And p2 > p1 Then
                src = Trim$(Left$(ln, p1 - 1)): dst = Trim$(Mid$(ln, p2 + 1)): v = ParseNumber(Mid$(ln, p1 + 1, p2 - p1 - 1))
            Else
                sep = ""
                If InStr(ln, vbTab) > 0 Then
                    sep = vbTab
                ElseIf InStr(ln, ";") > 0 Then
                    sep = ";"
                ElseIf InStr(ln, ",") > 0 Then
                    sep = ","
                End If
                If Len(sep) > 0 Then
                    parts = Split(ln, sep)
                    If UBound(parts) >= 2 Then src = Trim$(parts(0)): dst = Trim$(parts(1)): v = ParseNumber(parts(2))
                End If
            End If
            If Len(src) > 0 And Len(dst) > 0 And v > 0 Then SkAddLink src, dst, v
        End If
    Next i
End Sub

Option Explicit
' =====================================================================
'  ZenSankey - native-shape Sankey diagrams for Excel and PowerPoint
'  SankeyPowerPoint: PowerPoint entry points (ribbon callbacks, table or
'  text reading, options text box, example table).
'
'  Copyright (c) 2026 Samuel Mellot
'  Licensed under the Apache License, Version 2.0. See LICENSE and NOTICE.
' =====================================================================

Private Const PP_SELECTION_SHAPES As Long = 2
Private Const PP_SELECTION_TEXT As Long = 3
Private Const OPTIONS_SHAPE As String = "Sankey_Options"

' Draws a Sankey from the selected table (Source | Target | Amount ...) or text box ("Source [amount] Target" per line)
Public Sub DrawSankey()
    Dim sel As Object, shp As Object, sld As Object, grp As Object
    Dim L As Single, T As Single, W As Single, H As Single, slW As Single, slH As Single, ok As Boolean
    On Error GoTo Failed
    Set sel = Application.ActiveWindow.Selection
    If sel.Type <> PP_SELECTION_SHAPES And sel.Type <> PP_SELECTION_TEXT Then
        MsgBox "Select a table (Source | Target | Amount) or a text box with one flow per line (Source [amount] Target) first.", vbExclamation, "ZenSankey"
        Exit Sub
    End If
    Set shp = sel.ShapeRange(1)
    Set sld = shp.Parent
    SkReset
    If shp.HasTable Then
        ok = ReadTable(shp.Table)
    ElseIf shp.HasTextFrame Then
        ok = ReadText(shp.TextFrame.TextRange.Text)
    End If
    If Not ok Or SkLinkCount = 0 Then
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
    ApplyOptionsShape sld, L, T, W, H
    Set grp = SkDraw(sld.Shapes, L, T, W, H, "")
    grp.Select
    Exit Sub
Failed:
    MsgBox "The Sankey could not be drawn: " & Err.Description, vbCritical, "ZenSankey"
End Sub

' Reads the options text box named Sankey_Options on the slide, if present
Private Sub ApplyOptionsShape(ByVal sld As Object, ByRef L As Single, ByRef T As Single, ByRef W As Single, ByRef H As Single)
    Dim s As Object, txt As String, rows() As String, i As Long, p As Long, k As String, v As String, rest As String
    On Error Resume Next
    Set s = sld.Shapes(OPTIONS_SHAPE)
    On Error GoTo 0
    If s Is Nothing Then Exit Sub
    If Not s.HasTextFrame Then Exit Sub
    txt = s.TextFrame.TextRange.Text
    txt = Replace(txt, vbCrLf, vbCr): txt = Replace(txt, vbLf, vbCr): txt = Replace(txt, ChrW(11), vbCr)
    rows = Split(txt, vbCr)
    For i = 0 To UBound(rows)
        p = InStr(rows(i), "=")
        If p > 1 Then
            k = LCase$(Trim$(Left$(rows(i), p - 1))): v = Trim$(Mid$(rows(i), p + 1))
            Select Case k
                Case "left": If Len(v) > 0 Then L = ParseNumber(v)
                Case "top": If Len(v) > 0 Then T = ParseNumber(v)
                Case "width": If ParseNumber(v) > 50 Then W = ParseNumber(v)
                Case "height": If ParseNumber(v) > 50 Then H = ParseNumber(v)
                Case Else: rest = rest & rows(i) & vbCr
            End Select
        End If
    Next i
    SkApplyOptionText rest
End Sub

' Inserts (or selects) the Sankey_Options text box next to the current slide, listing every option
Public Sub InsertOptionsBox()
    Dim sld As Object, s As Object, lst As Variant, i As Long, parts() As String, txt As String, slW As Single
    On Error GoTo Failed
    Set sld = Application.ActiveWindow.View.Slide
    slW = Application.ActivePresentation.PageSetup.SlideWidth
    On Error Resume Next
    Set s = sld.Shapes(OPTIONS_SHAPE)
    On Error GoTo Failed
    If Not s Is Nothing Then
        s.Select
        MsgBox "The options box already exists on this slide (selected). Edit its lines; they apply to the next Sankey drawn on this slide.", vbInformation, "ZenSankey"
        Exit Sub
    End If
    txt = "' ZenSankey options - one key = value per line, applies to this slide" & vbCr
    txt = txt & "' left = 36" & vbCr & "' top = 90" & vbCr & "' width = " & Format$(slW - 72, "0") & vbCr & "' height = 360" & vbCr
    lst = SkOptionsList()
    For i = LBound(lst) To UBound(lst)
        parts = Split(lst(i), "|")
        txt = txt & parts(0) & " = " & parts(1)
        If Len(parts(3)) > 0 Then txt = txt & "    ' " & Replace(parts(3), ",", " | ") Else txt = txt & "    ' " & parts(2)
        txt = txt & vbCr
    Next i
    Set s = sld.Shapes.AddTextbox(1, slW + 24, 20, 420, 500)
    With s
        .Name = OPTIONS_SHAPE
        .Fill.Visible = -1: .Fill.ForeColor.RGB = RGB(255, 255, 230)
        .Line.Visible = -1: .Line.ForeColor.RGB = RGB(2, 128, 144)
        .TextFrame.WordWrap = 0
        .TextFrame.AutoSize = 1
        .TextFrame.TextRange.Text = txt
        .TextFrame.TextRange.Font.Name = "Consolas"
        .TextFrame.TextRange.Font.Size = 9
    End With
    s.Select
    MsgBox "An options box named " & OPTIONS_SHAPE & " was added to the right of the slide (outside the slide, so it is not shown in slide shows)." & vbCrLf & _
           "Edit the values after the = signs; lines starting with ' are comments. Remove the quote before left/top/width/height to set the position and size.", vbInformation, "ZenSankey"
    Exit Sub
Failed:
    MsgBox "The options box could not be created: " & Err.Description, vbCritical, "ZenSankey"
End Sub

' Inserts an example table on the current slide
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
    Set tb = sld.Shapes.AddTable(UBound(d) + 1, 4, 36, 60, slW / 2 - 50, 20 * (UBound(d) + 1))
    For i = 0 To UBound(d)
        For j = 0 To 3
            tb.Table.Cell(i + 1, j + 1).Shape.TextFrame.TextRange.Text = d(i)(j)
            tb.Table.Cell(i + 1, j + 1).Shape.TextFrame.TextRange.Font.Size = 10
        Next j
    Next i
    tb.Select
    MsgBox "Example table inserted and selected: click Draw Sankey. You can also paste cells copied from Excel (PowerPoint tables are recognized).", vbInformation, "ZenSankey"
    Exit Sub
Failed:
    MsgBox "The example could not be inserted: " & Err.Description, vbCritical, "ZenSankey"
End Sub

Public Sub ShowHelp()
    MsgBox "ZenSankey for PowerPoint" & vbCrLf & vbCrLf & _
           "1. Insert a table Source | Target | Amount (Insert > Table, or paste Excel cells), or a text box with one flow per line: Source [amount] Target." & vbCrLf & _
           "2. Select the table or the text box, then click Draw Sankey." & vbCrLf & _
           "3. Click Options to add an options box next to the slide: palette, colors, link shape, labels, titles, section titles, position and size..." & vbCrLf & _
           "4. The diagram is a group of native shapes: move it, resize it, change colors and fonts. Delete the source table if you wish." & vbCrLf & vbCrLf & _
           "Optional table columns: Color (#RRGGBB), Source column, Target column (0 = first column).", vbInformation, "ZenSankey"
End Sub

' ---- Ribbon callbacks --------------------------------------------------
Public Sub Ribbon_Draw(control As IRibbonControl)
    DrawSankey
End Sub

Public Sub Ribbon_Example(control As IRibbonControl)
    InsertExample
End Sub

Public Sub Ribbon_Options(control As IRibbonControl)
    InsertOptionsBox
End Sub

Public Sub Ribbon_Help(control As IRibbonControl)
    ShowHelp
End Sub

' ---- Data reading --------------------------------------------------------
Private Function ReadTable(ByVal tbl As Object) As Boolean
    Dim r As Long, r0 As Long, nR As Long, nC As Long, src As String, dst As String, v As Double
    Dim colHex As String, cs As Long, cd As Long, t3 As String
    nR = tbl.Rows.Count: nC = tbl.Columns.Count
    If nC < 3 Then Exit Function
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
    ReadTable = True
End Function

Private Function CellText(ByVal tbl As Object, ByVal r As Long, ByVal c As Long) As String
    On Error Resume Next
    CellText = tbl.Cell(r, c).Shape.TextFrame.TextRange.Text
End Function

' Free text: "Source [amount] Target" (SankeyMATIC syntax), "Source ; Target ; amount", tabs or commas
Private Function ReadText(ByVal txt As String) As Boolean
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
    ReadText = True
End Function

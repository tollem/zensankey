Option Explicit
' =====================================================================
'  ZenSankey - native-shape Sankey diagrams for Excel and PowerPoint
'  SankeyCore: style settings, layout engine and drawing (shared by both
'  hosts; it only relies on the Office Shapes object model).
'
'  Copyright (c) 2026 Samuel Mellot
'  Licensed under the Apache License, Version 2.0. See LICENSE and NOTICE.
' =====================================================================

' ---- Office constants (numeric, so no type library is required) -----
Private Const C_msoEditingAuto As Long = 0
Private Const C_msoEditingCorner As Long = 1
Private Const C_msoSegmentLine As Long = 0
Private Const C_msoSegmentCurve As Long = 1
Private Const C_msoShapeRectangle As Long = 1
Private Const C_msoShapeRoundedRectangle As Long = 5
Private Const C_msoTextOrientationHorizontal As Long = 1
Private Const C_msoAutoSizeShapeToFitText As Long = 1
Private Const C_msoAlignLeft As Long = 1
Private Const C_msoAlignCenter As Long = 2
Private Const C_msoAlignRight As Long = 3
Private Const C_msoGradientVertical As Long = 2
Private Const C_msoBringToFront As Long = 0
Private Const C_msoSendToBack As Long = 1
Private Const C_msoFalse As Long = 0
Private Const C_msoTrue As Long = -1

Private Type SkNode
    Label As String
    Col As Long
    ColForced As Long
    ValIn As Double
    ValOut As Double
    Amount As Double
    X As Single
    Y As Single
    H As Single
    OutOff As Single
    InOff As Single
    Fill As Long
    FillSet As Boolean
    Rank As Double
    HasIn As Boolean
    HasOut As Boolean
    OnLeft As Boolean
    LW As Single
    LH As Single
    LX As Single
    LY As Single
    LblName As String
End Type

Private Type SkLink
    Src As Long
    Dst As Long
    Amount As Double
    Fill As Long
    HasColor As Boolean
    Y0 As Single
    Y1 As Single
    H As Single
End Type

' ---- Style options (defaults in SkDefaults) ---------------------------
Private oPalette As String
Private oColors As String
Private oColorBy As String
Private oNodeColor As String
Private oLinkColor As String
Private oLinkShape As String
Private oCurvature As Single
Private oLinkOpacity As Single
Private oLinkOutline As Boolean
Private oNodeWidth As Single
Private oNodePadding As Single
Private oRoundedNodes As Boolean
Private oFont As String
Private oFontSize As Single
Private oTextColor As Long
Private oLabels As String
Private oLabelLines As Long
Private oOutsideLabels As Boolean
Private oLabelBackground As Boolean
Private oLinkValues As Boolean
Private oValueFormat As String
Private oDecimals As Long
Private oUnit As String
Private oTitle As String
Private oSubtitle As String
Private oColumnTitles As String
Private oColumnTitleSize As Single
Private oBackground As Boolean
Private oBackgroundColor As Long
Private oAlignment As String
Private oOrder As String
Private oMargin As Single
Private mStyleReady As Boolean

Private mNodes() As SkNode
Private mLinks() As SkLink
Private mNodeCount As Long
Private mLinkCount As Long
Private mMaxCol As Long
Private mTotalRef As Double
Private mPal() As Long
Private mPalCount As Long
Private mStamp As String
Private mColX() As Single
Private mPadEff As Single
Private mAreaL As Single
Private mAreaT As Single
Private mAreaW As Single
Private mAreaH As Single
Private mBandTitle As Single
Private mBandCol As Single
Private mBoxW As Single
Private mBoxH As Single

' =====================================================================
'  Public API: data and style
' =====================================================================
' Clears the flows (the style is kept)
Public Sub SkReset()
    If Not mStyleReady Then SkDefaults
    mNodeCount = 0: mLinkCount = 0
    ReDim mNodes(1 To 1): ReDim mLinks(1 To 1)
End Sub

' Restores the default style
Public Sub SkResetStyle()
    SkDefaults
End Sub

Public Sub SkDefaults()
    oPalette = "Teal": oColors = "": oColorBy = "flow"
    oNodeColor = "palette": oLinkColor = "target"
    oLinkShape = "curved": oCurvature = 0.5: oLinkOpacity = 0.45: oLinkOutline = False
    oNodeWidth = 12: oNodePadding = 16: oRoundedNodes = False
    oFont = "Arial": oFontSize = 10: oTextColor = RGB(28, 43, 45)
    oLabels = "name_value": oLabelLines = 2: oOutsideLabels = True: oLabelBackground = False
    oLinkValues = False: oValueFormat = "standard": oDecimals = 0: oUnit = ""
    oTitle = "": oSubtitle = "": oColumnTitles = "": oColumnTitleSize = 10
    oBackground = False: oBackgroundColor = RGB(247, 251, 251)
    oAlignment = "justify": oOrder = "auto": oMargin = 8
    mStyleReady = True
End Sub

' Option catalogue: array of "key|default|description|choices" strings
Public Function SkOptionsList() As Variant
    Dim a(1 To 32) As String
    a(1) = "title||Title shown above the diagram|"
    a(2) = "subtitle||Subtitle (smaller, gray)|"
    a(3) = "column_titles||Section titles above each column, separated by ; (e.g. Sources;Allocation;Outcome)|"
    a(4) = "palette|Teal|Color palette|Teal,Vivid,Pastel,Muted,Ocean,Warm,Mono,Custom"
    a(5) = "colors||Custom colors #RRGGBB separated by ; (used when palette = Custom)|"
    a(6) = "color_by|flow|flow: a node keeps the color of its main source, branches get new colors; node: one color per node; column: one per column|flow,node,column"
    a(7) = "node_color|palette|Node color: palette, gray, or a #RRGGBB code|palette,gray"
    a(8) = "link_color|target|Link color: target, source, gradient (source to target), gray, or #RRGGBB|target,source,gradient,gray"
    a(9) = "link_shape|curved|Link shape|curved,straight"
    a(10) = "curvature|0.5|Link curvature (0.2 to 0.8)|"
    a(11) = "link_opacity|45|Link opacity in % (100 = opaque)|"
    a(12) = "link_outline|no|Thin white outline around links|yes,no"
    a(13) = "node_width|12|Node width (points)|"
    a(14) = "node_padding|16|Minimum vertical gap between nodes (points)|"
    a(15) = "rounded_nodes|no|Nodes with rounded corners|yes,no"
    a(16) = "font|Arial|Font name|"
    a(17) = "font_size|10|Label font size|"
    a(18) = "text_color|#1C2B2D|Text color #RRGGBB|"
    a(19) = "labels|name_value|Node label content|name_value,name,value,none"
    a(20) = "label_lines|2|Label on 2 lines (name / value) or 1 line (name (value))|1,2"
    a(21) = "outside_labels|yes|Labels outside the diagram for the first and last columns|yes,no"
    a(22) = "label_background|no|Translucent white box behind labels (readability over links)|yes,no"
    a(23) = "link_values|no|Show the amount in the middle of links that are tall enough|yes,no"
    a(24) = "value_format|standard|Amount format: standard, k (thousands), M (millions), percent (of total inflow)|standard,k,M,percent"
    a(25) = "decimals|0|Number of decimals for amounts|"
    a(26) = "unit||Unit appended to amounts (e.g. CHF, kWh)|"
    a(27) = "column_title_size|10|Font size of column titles|"
    a(28) = "background|no|Rounded background rectangle behind the diagram|yes,no"
    a(29) = "background_color|#F7FBFB|Background color #RRGGBB|"
    a(30) = "alignment|justify|justify: nodes without outflow are pushed to the last column; left: columns by depth only|justify,left"
    a(31) = "order|auto|Vertical order of nodes: auto (fewer crossings) or table (order of appearance)|auto,table"
    a(32) = "margin|8|Inner margin (points)|"
    SkOptionsList = a
End Function

' Applies option text, one "key = value" per line
Public Sub SkApplyOptionText(ByVal txt As String)
    Dim rows() As String, i As Long, p As Long, k As String, v As String
    If Not mStyleReady Then SkDefaults
    txt = Replace(txt, vbCrLf, vbCr): txt = Replace(txt, vbLf, vbCr): txt = Replace(txt, ChrW(11), vbCr)
    rows = Split(txt, vbCr)
    For i = 0 To UBound(rows)
        p = InStr(rows(i), "=")
        If p > 1 Then
            k = Trim$(Left$(rows(i), p - 1)): v = Trim$(Mid$(rows(i), p + 1))
            SkSetOption k, v
        End If
    Next i
End Sub

' Sets one option (key is case-insensitive; unknown keys are ignored)
Public Sub SkSetOption(ByVal k As String, ByVal v As String)
    If Not mStyleReady Then SkDefaults
    k = LCase$(Trim$(k)): v = Trim$(v)
    Select Case k
        Case "title": oTitle = v
        Case "subtitle": oSubtitle = v
        Case "column_titles": oColumnTitles = v
        Case "palette": If Len(v) > 0 Then oPalette = v
        Case "colors": oColors = v
        Case "color_by": If Len(v) > 0 Then oColorBy = LCase$(v)
        Case "node_color": If Len(v) > 0 Then oNodeColor = v
        Case "link_color": If Len(v) > 0 Then oLinkColor = v
        Case "link_shape": If Len(v) > 0 Then oLinkShape = LCase$(v)
        Case "curvature": If Len(v) > 0 Then oCurvature = Clamp(ParseNumber(v), 0.1, 0.9)
        Case "link_opacity": If Len(v) > 0 Then oLinkOpacity = Clamp(ParseNumber(v), 5, 100) / 100
        Case "link_outline": oLinkOutline = YesNo(v, False)
        Case "node_width": If Len(v) > 0 Then oNodeWidth = Clamp(ParseNumber(v), 2, 80)
        Case "node_padding": If Len(v) > 0 Then oNodePadding = Clamp(ParseNumber(v), 0, 100)
        Case "rounded_nodes": oRoundedNodes = YesNo(v, False)
        Case "font": If Len(v) > 0 Then oFont = v
        Case "font_size": If Len(v) > 0 Then oFontSize = Clamp(ParseNumber(v), 5, 40)
        Case "text_color": If IsHexColor(v) Then oTextColor = HexToRGB(v)
        Case "labels": If Len(v) > 0 Then oLabels = LCase$(v)
        Case "label_lines": If Len(v) > 0 Then oLabelLines = IIf(ParseNumber(v) < 2, 1, 2)
        Case "outside_labels": oOutsideLabels = YesNo(v, True)
        Case "label_background": oLabelBackground = YesNo(v, False)
        Case "link_values": oLinkValues = YesNo(v, False)
        Case "value_format": If Len(v) > 0 Then oValueFormat = LCase$(v)
        Case "decimals": If Len(v) > 0 Then oDecimals = CLng(Clamp(ParseNumber(v), 0, 4))
        Case "unit": oUnit = v
        Case "column_title_size": If Len(v) > 0 Then oColumnTitleSize = Clamp(ParseNumber(v), 5, 40)
        Case "background": oBackground = YesNo(v, False)
        Case "background_color": If IsHexColor(v) Then oBackgroundColor = HexToRGB(v)
        Case "alignment": If Len(v) > 0 Then oAlignment = LCase$(v)
        Case "order": If Len(v) > 0 Then oOrder = LCase$(v)
        Case "margin": If Len(v) > 0 Then oMargin = Clamp(ParseNumber(v), 0, 100)
    End Select
End Sub

' Current value of one option, as text
Public Function SkGetOption(ByVal k As String) As String
    If Not mStyleReady Then SkDefaults
    Select Case LCase$(Trim$(k))
        Case "title": SkGetOption = oTitle
        Case "subtitle": SkGetOption = oSubtitle
        Case "column_titles": SkGetOption = oColumnTitles
        Case "palette": SkGetOption = oPalette
        Case "colors": SkGetOption = oColors
        Case "color_by": SkGetOption = oColorBy
        Case "node_color": SkGetOption = oNodeColor
        Case "link_color": SkGetOption = oLinkColor
        Case "link_shape": SkGetOption = oLinkShape
        Case "curvature": SkGetOption = Replace(Format$(oCurvature, "0.0#"), ",", ".")
        Case "link_opacity": SkGetOption = CStr(CLng(oLinkOpacity * 100))
        Case "link_outline": SkGetOption = IIf(oLinkOutline, "yes", "no")
        Case "node_width": SkGetOption = CStr(CLng(oNodeWidth))
        Case "node_padding": SkGetOption = CStr(CLng(oNodePadding))
        Case "rounded_nodes": SkGetOption = IIf(oRoundedNodes, "yes", "no")
        Case "font": SkGetOption = oFont
        Case "font_size": SkGetOption = CStr(CLng(oFontSize))
        Case "text_color": SkGetOption = RGBToHex(oTextColor)
        Case "labels": SkGetOption = oLabels
        Case "label_lines": SkGetOption = CStr(oLabelLines)
        Case "outside_labels": SkGetOption = IIf(oOutsideLabels, "yes", "no")
        Case "label_background": SkGetOption = IIf(oLabelBackground, "yes", "no")
        Case "link_values": SkGetOption = IIf(oLinkValues, "yes", "no")
        Case "value_format": SkGetOption = oValueFormat
        Case "decimals": SkGetOption = CStr(oDecimals)
        Case "unit": SkGetOption = oUnit
        Case "column_title_size": SkGetOption = CStr(CLng(oColumnTitleSize))
        Case "background": SkGetOption = IIf(oBackground, "yes", "no")
        Case "background_color": SkGetOption = RGBToHex(oBackgroundColor)
        Case "alignment": SkGetOption = oAlignment
        Case "order": SkGetOption = oOrder
        Case "margin": SkGetOption = CStr(CLng(oMargin))
    End Select
End Function

' Every option as "key = value" lines (used to store the style with a diagram)
Public Function SkSettingsText() As String
    Dim lst As Variant, i As Long, parts() As String, s As String
    lst = SkOptionsList()
    For i = LBound(lst) To UBound(lst)
        parts = Split(lst(i), "|")
        s = s & parts(0) & " = " & SkGetOption(parts(0)) & vbLf
    Next i
    SkSettingsText = s
End Function

' Adds a flow. colorHex: "#RRGGBB" or ""; srcCol / dstCol: forced column (0 = first), -1 = automatic
Public Sub SkAddLink(ByVal srcName As String, ByVal dstName As String, ByVal amount As Double, _
                     Optional ByVal colorHex As String = "", Optional ByVal srcCol As Long = -1, Optional ByVal dstCol As Long = -1)
    Dim s As Long, d As Long
    srcName = Trim$(srcName): dstName = Trim$(dstName)
    If Len(srcName) = 0 Or Len(dstName) = 0 Or amount <= 0 Then Exit Sub
    s = NodeIndex(srcName): d = NodeIndex(dstName)
    If srcCol >= 0 Then mNodes(s).ColForced = srcCol
    If dstCol >= 0 Then mNodes(d).ColForced = dstCol
    mLinkCount = mLinkCount + 1
    ReDim Preserve mLinks(1 To mLinkCount)
    With mLinks(mLinkCount)
        .Src = s: .Dst = d: .Amount = amount
        .HasColor = IsHexColor(colorHex)
        If .HasColor Then .Fill = HexToRGB(colorHex)
    End With
    mNodes(s).ValOut = mNodes(s).ValOut + amount: mNodes(s).HasOut = True
    mNodes(d).ValIn = mNodes(d).ValIn + amount: mNodes(d).HasIn = True
End Sub

Public Function SkLinkCount() As Long
    SkLinkCount = mLinkCount
End Function

' =====================================================================
'  Public API: ribbon helpers (control id -> option)
' =====================================================================
Public Function SkRibbonKey(ByVal id As String) As String
    Select Case id
        Case "ddPalette": SkRibbonKey = "palette"
        Case "ddColorBy": SkRibbonKey = "color_by"
        Case "ddLinkColor": SkRibbonKey = "link_color"
        Case "ddOpacity": SkRibbonKey = "link_opacity"
        Case "ddShape": SkRibbonKey = "shape"
        Case "ddNodeWidth": SkRibbonKey = "node_width"
        Case "ddLabels": SkRibbonKey = "labels"
        Case "ddFontSize": SkRibbonKey = "font_size"
        Case "ddFont": SkRibbonKey = "font"
        Case "ddFormat": SkRibbonKey = "value_format"
        Case "ddDecimals": SkRibbonKey = "decimals"
        Case "ddOrder": SkRibbonKey = "order"
        Case "cbRounded": SkRibbonKey = "rounded_nodes"
        Case "cbOutline": SkRibbonKey = "link_outline"
        Case "cbOutside": SkRibbonKey = "outside_labels"
        Case "cbLinkValues": SkRibbonKey = "link_values"
        Case "cbLabelBg": SkRibbonKey = "label_background"
        Case "cbBackground": SkRibbonKey = "background"
        Case "ebTitle": SkRibbonKey = "title"
        Case "ebSubtitle": SkRibbonKey = "subtitle"
        Case "ebColumns": SkRibbonKey = "column_titles"
        Case "ebUnit": SkRibbonKey = "unit"
    End Select
End Function

' Item values of a ribbon drop-down, in item order
Public Function SkRibbonValues(ByVal id As String) As String
    Select Case id
        Case "ddPalette": SkRibbonValues = "Teal;Vivid;Pastel;Muted;Ocean;Warm;Mono"
        Case "ddColorBy": SkRibbonValues = "flow;node;column"
        Case "ddLinkColor": SkRibbonValues = "target;source;gradient;gray"
        Case "ddOpacity": SkRibbonValues = "25;35;45;55;70;85;100"
        Case "ddShape": SkRibbonValues = "curved;soft;straight"
        Case "ddNodeWidth": SkRibbonValues = "8;12;18;26"
        Case "ddLabels": SkRibbonValues = "name_value;name;value;none"
        Case "ddFontSize": SkRibbonValues = "8;9;10;11;12;14;16;18"
        Case "ddFont": SkRibbonValues = "Arial;Calibri;Helvetica;Verdana;Georgia;Times New Roman"
        Case "ddFormat": SkRibbonValues = "standard;k;M;percent"
        Case "ddDecimals": SkRibbonValues = "0;1;2"
        Case "ddOrder": SkRibbonValues = "auto;table"
    End Select
End Function

' Index of the current value in a drop-down (nearest value for numeric lists)
Public Function SkRibbonIndex(ByVal id As String) As Long
    Dim vals() As String, cur As String, i As Long, best As Long, bestD As Double, d As Double, k As String
    vals = Split(SkRibbonValues(id), ";")
    k = SkRibbonKey(id)
    If k = "shape" Then
        If oLinkShape = "straight" Then
            cur = "straight"
        ElseIf oCurvature < 0.4 Then
            cur = "soft"
        Else
            cur = "curved"
        End If
    Else
        cur = SkGetOption(k)
    End If
    For i = 0 To UBound(vals)
        If StrComp(vals(i), cur, vbTextCompare) = 0 Then SkRibbonIndex = i: Exit Function
    Next i
    If IsNumeric(cur) And IsNumeric(vals(0)) Then
        best = 0: bestD = -1
        For i = 0 To UBound(vals)
            d = Abs(Val(vals(i)) - Val(cur))
            If bestD < 0 Or d < bestD Then best = i: bestD = d
        Next i
        SkRibbonIndex = best
    Else
        SkRibbonIndex = 0
    End If
End Function

' Applies the item chosen in a drop-down
Public Sub SkRibbonSelect(ByVal id As String, ByVal idx As Long)
    Dim vals() As String, v As String
    vals = Split(SkRibbonValues(id), ";")
    If idx < 0 Or idx > UBound(vals) Then Exit Sub
    v = vals(idx)
    If id = "ddShape" Then
        Select Case v
            Case "straight": SkSetOption "link_shape", "straight"
            Case "soft": SkSetOption "link_shape", "curved": SkSetOption "curvature", "0.3"
            Case Else: SkSetOption "link_shape", "curved": SkSetOption "curvature", "0.5"
        End Select
    Else
        SkSetOption SkRibbonKey(id), v
    End If
End Sub

Public Function SkRibbonPressed(ByVal id As String) As Boolean
    SkRibbonPressed = YesNo(SkGetOption(SkRibbonKey(id)), False)
End Function

Public Sub SkRibbonSetPressed(ByVal id As String, ByVal pressed As Boolean)
    SkSetOption SkRibbonKey(id), IIf(pressed, "yes", "no")
End Sub

Public Function SkRibbonText(ByVal id As String) As String
    SkRibbonText = SkGetOption(SkRibbonKey(id))
End Function

Public Sub SkRibbonSetText(ByVal id As String, ByVal txt As String)
    SkSetOption SkRibbonKey(id), txt
End Sub

' =====================================================================
'  Public API: drawing
' =====================================================================
' Draws the diagram into the given Shapes collection (worksheet or slide) and returns the group.
' pGrow: the width and height may grow so that every label fits (Excel); otherwise the
' font is reduced to fit the given box (PowerPoint).
Public Function SkDraw(ByVal shp As Object, ByVal pLeft As Single, ByVal pTop As Single, _
                       ByVal pWidth As Single, ByVal pHeight As Single, Optional ByVal pTitle As String = "", _
                       Optional ByVal pGrow As Boolean = False) As Object
    Dim i As Long, names() As Variant, n As Long, grp As Object, c As Long, titles() As String
    Dim yTitle As Single, lbl As Object, fs0 As Single, cs0 As Single, errNo As Long, errDesc As String
    If mLinkCount = 0 Then Err.Raise vbObjectError + 1, "SankeyCore", "No flows to draw."
    If Len(pTitle) > 0 And Len(oTitle) = 0 Then oTitle = pTitle
    mStamp = Format$(Now, "hhnnss")
    fs0 = oFontSize: cs0 = oColumnTitleSize
    On Error GoTo Failed
    PrepareLayout shp, pLeft, pTop, pWidth, pHeight, pGrow
    ReDim names(0 To mLinkCount * 2 + 2 * mNodeCount + mMaxCol + 4)
    n = -1
    If oBackground Then n = n + 1: names(n) = DrawBackground(shp, pLeft, pTop, mBoxW, mBoxH)
    For i = 1 To mLinkCount
        n = n + 1: names(n) = DrawLink(shp, i, n)
    Next i
    For i = 1 To mNodeCount
        n = n + 1: names(n) = DrawNode(shp, i, n)
    Next i
    If oLabels <> "none" Then
        For i = 1 To mNodeCount
            Set lbl = shp.Item(mNodes(i).LblName)
            lbl.Left = mNodes(i).LX: lbl.Top = mNodes(i).LY
            lbl.ZOrder C_msoBringToFront
            n = n + 1: names(n) = mNodes(i).LblName
        Next i
    End If
    If oLinkValues Then
        For i = 1 To mLinkCount
            If mLinks(i).H >= oFontSize * 1.3 Then n = n + 1: names(n) = DrawLinkValue(shp, i, n)
        Next i
    End If
    If Len(oColumnTitles) > 0 Then
        titles = Split(oColumnTitles, ";")
        For c = 0 To mMaxCol
            If c <= UBound(titles) Then
                If Len(Trim$(titles(c))) > 0 Then n = n + 1: names(n) = DrawColumnTitle(shp, c, Trim$(titles(c)), pTop + oMargin + mBandTitle, n)
            End If
        Next c
    End If
    yTitle = pTop + oMargin
    If Len(oTitle) > 0 Then
        n = n + 1: names(n) = DrawTitle(shp, oTitle, pLeft + oMargin, yTitle, mBoxW - 2 * oMargin, oFontSize * 1.6, False, n)
        yTitle = yTitle + oFontSize * 1.6 * 1.5
    End If
    If Len(oSubtitle) > 0 Then n = n + 1: names(n) = DrawTitle(shp, oSubtitle, pLeft + oMargin, yTitle, mBoxW - 2 * oMargin, oFontSize * 1.1, True, n)
    ReDim Preserve names(0 To n)
    Set grp = shp.Range(names).Group
    grp.Name = "Sankey " & mStamp
    oFontSize = fs0: oColumnTitleSize = cs0
    Set SkDraw = grp
    Exit Function
Failed:
    errNo = Err.Number: errDesc = Err.Description
    oFontSize = fs0: oColumnTitleSize = cs0
    Err.Raise errNo, "SankeyCore", errDesc
End Function

' Size of the box actually used by the last SkDraw (points)
Public Function SkDrawnWidth() As Single
    SkDrawnWidth = mBoxW
End Function

Public Function SkDrawnHeight() As Single
    SkDrawnHeight = mBoxH
End Function

' Computes the layout without drawing (label sizes are estimated) and returns a text description
Public Function SkDebugLayout(ByVal pWidth As Single, ByVal pHeight As Single, Optional ByVal pGrow As Boolean = True) As String
    Dim i As Long, l As Long, s As String, c1 As Long, c2 As Long, g As Boolean, c As Long, fs0 As Single, cs0 As Single
    fs0 = oFontSize: cs0 = oColumnTitleSize
    PrepareLayout Nothing, 0, 0, pWidth, pHeight, pGrow
    s = "BOX|" & Format$(mBoxW, "0.0") & "|" & Format$(mBoxH, "0.0") & "|" & Format$(mAreaL, "0.0") & "|" & Format$(mAreaT, "0.0") & "|" & Format$(mAreaW, "0.0") & "|" & Format$(mAreaH, "0.0") & "|" & Format$(oFontSize, "0.0") & "|" & Format$(mPadEff, "0.0") & vbLf
    For c = 0 To mMaxCol
        s = s & "C|" & c & "|" & Format$(mColX(c), "0.0") & vbLf
    Next c
    For i = 1 To mNodeCount
        s = s & "N|" & mNodes(i).Label & "|" & mNodes(i).Col & "|" & Format$(mNodes(i).X, "0.0") & "|" & Format$(mNodes(i).Y, "0.0") & "|" & Format$(mNodes(i).H, "0.0") & "|" & mNodes(i).Amount & "|" & FormatAmount(mNodes(i).Amount) & "|" & RGBToHex(NodeFillColor(i)) & "|" & Format$(mNodes(i).LX, "0.0") & "|" & Format$(mNodes(i).LY, "0.0") & "|" & Format$(mNodes(i).LW, "0.0") & "|" & Format$(mNodes(i).LH, "0.0") & "|" & IIf(mNodes(i).OnLeft, "L", "R") & vbLf
    Next i
    For l = 1 To mLinkCount
        LinkColors l, c1, c2, g
        s = s & "L|" & mNodes(mLinks(l).Src).Label & "|" & mNodes(mLinks(l).Dst).Label & "|" & Format$(mLinks(l).Y0, "0.0") & "|" & Format$(mLinks(l).Y1, "0.0") & "|" & Format$(mLinks(l).H, "0.0") & "|" & RGBToHex(c1) & "|" & RGBToHex(c2) & "|" & mLinks(l).Amount & vbLf
    Next l
    s = s & "OPT|" & oPalette & "|" & oColorBy & "|" & oLinkColor & "|" & oLinkShape & "|" & oLinkOpacity & "|" & oLabels & "|" & oValueFormat & "|" & oColumnTitles & "|" & oTitle & "|" & oSubtitle & "|" & oNodeWidth & "|" & oCurvature & "|" & oFont & "|" & oColumnTitleSize & vbLf
    oFontSize = fs0: oColumnTitleSize = cs0
    SkDebugLayout = s
End Function

' "#RRGGBB" or "RRGGBB" -> VBA RGB Long
Public Function HexToRGB(ByVal h As String) As Long
    Dim r As Long, g As Long, b As Long
    h = Trim$(h)
    If Left$(h, 1) = "#" Then h = Mid$(h, 2)
    If Len(h) <> 6 Then HexToRGB = RGB(2, 128, 144): Exit Function
    r = CLng("&H" & Mid$(h, 1, 2)): g = CLng("&H" & Mid$(h, 3, 2)): b = CLng("&H" & Mid$(h, 5, 2))
    HexToRGB = RGB(r, g, b)
End Function

' VBA RGB Long -> "#RRGGBB"
Public Function RGBToHex(ByVal c As Long) As String
    Dim r As Long, g As Long, b As Long
    r = c Mod 256: g = (c \ 256) Mod 256: b = (c \ 65536) Mod 256
    RGBToHex = "#" & Right$("0" & Hex$(r), 2) & Right$("0" & Hex$(g), 2) & Right$("0" & Hex$(b), 2)
End Function

' Is the string a valid "#RRGGBB" color?
Public Function IsHexColor(ByVal s As String) As Boolean
    Dim i As Long, c As String
    s = Trim$(s)
    If Left$(s, 1) = "#" Then s = Mid$(s, 2)
    If Len(s) <> 6 Then Exit Function
    For i = 1 To 6
        c = UCase$(Mid$(s, i, 1))
        If InStr("0123456789ABCDEF", c) = 0 Then Exit Function
    Next i
    IsHexColor = True
End Function

' Tolerant number parsing: "3 000", "3'000", "2 650,5", "1,234.5", "55 %"
Public Function ParseNumber(ByVal v As Variant) As Double
    Dim s As String, i As Long, c As String, out As String, hasDot As Boolean
    If IsNumeric(v) And VarType(v) <> vbString Then ParseNumber = CDbl(v): Exit Function
    s = CStr(v)
    For i = 1 To Len(s)
        If Mid$(s, i, 1) = "." Then hasDot = True
    Next i
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        Select Case c
            Case "0" To "9", "-": out = out & c
            Case ".": out = out & "."
            Case ",": If Not hasDot Then out = out & "."
            Case Else ' spaces, apostrophes, letters: ignored
        End Select
    Next i
    If Len(out) = 0 Then ParseNumber = 0 Else ParseNumber = Val(out)
End Function

' Formats an amount according to the options (standard / k / M / percent, decimals, unit)
Public Function FormatAmount(ByVal v As Double) As String
    Dim f As String, s As String, d As Long
    If Not mStyleReady Then SkDefaults
    d = oDecimals
    If d > 0 Then f = "#,##0." & String$(d, "0") Else f = "#,##0"
    Select Case oValueFormat
        Case "k": s = Format$(v / 1000, f) & " k"
        Case "m": s = Format$(v / 1000000, f) & " M"
        Case "percent", "%"
            If mTotalRef > 0 Then s = Format$(100 * v / mTotalRef, f) & " %" Else s = Format$(v, f)
        Case Else
            If d = 0 And Abs(v - Int(v)) > 0.0001 Then f = "#,##0.##"
            s = Format$(v, f)
            If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
    End Select
    If Len(oUnit) > 0 Then s = s & " " & oUnit
    FormatAmount = s
End Function

' Palette color (index 0..n-1, cyclic)
Public Function PaletteColor(ByVal idx As Long) As Long
    If mPalCount = 0 Then BuildPalette
    PaletteColor = mPal(idx Mod mPalCount)
End Function

' =====================================================================
'  Palette and colors
' =====================================================================
Private Sub BuildPalette()
    Dim lst As String, parts() As String, i As Long, n As Long, dflt As String
    dflt = "#028090;#6250D6;#E34948;#1BAF7A;#E88E3B;#2A6F97;#D9A441;#7F8C8D;#B86B2C;#4FB0BC"
    Select Case LCase$(oPalette)
        Case "vivid": lst = "#2A78D6;#EB6834;#1BAF7A;#EDA100;#E87BA4;#008300;#6250D6;#E34948;#9AB3B3"
        Case "pastel": lst = "#8FBFD9;#F4B183;#A9D8B8;#FFE699;#F8CBAD;#C9B8E8;#B4C7E7;#BFBFBF;#F2E2A0"
        Case "muted": lst = "#2F4858;#33658A;#86BBD8;#4A6670;#7A8C94;#A8B6BD;#5C7F8D;#9FB7C2;#C5D5DD"
        Case "ocean": lst = "#03045E;#0077B6;#00B4D8;#90E0EF;#023E8A;#0096C7;#48CAE4;#5390D9;#80B3FF"
        Case "warm": lst = "#C0392B;#E67E22;#F1C40F;#8E44AD;#2C3E50;#16A085;#D35400;#7F8C8D;#27AE60"
        Case "mono": lst = "#0B4F6C;#2F8FA6;#7CCBD6;#1B6F8A;#A9E0E6;#4FAFC0;#5B7C87;#8CA3AA;#34606E"
        Case "custom": lst = oColors
        Case Else: lst = dflt
    End Select
    If Len(Trim$(lst)) = 0 Then lst = dflt
    lst = Replace(lst, ",", ";")
    parts = Split(lst, ";")
    ReDim mPal(0 To UBound(parts))
    n = 0
    For i = 0 To UBound(parts)
        If IsHexColor(parts(i)) Then mPal(n) = HexToRGB(parts(i)): n = n + 1
    Next i
    If n = 0 Then mPal(0) = RGB(2, 128, 144): n = 1
    mPalCount = n
End Sub

Private Function NodeFillColor(ByVal i As Long) As Long
    If IsHexColor(oNodeColor) Then
        NodeFillColor = HexToRGB(oNodeColor)
    ElseIf LCase$(oNodeColor) = "gray" Or LCase$(oNodeColor) = "grey" Then
        NodeFillColor = RGB(70, 80, 84)
    Else
        NodeFillColor = mNodes(i).Fill
    End If
End Function

' Link colors: c1 (source side) and c2 (target side); isGrad = gradient fill to apply
Private Sub LinkColors(ByVal l As Long, ByRef c1 As Long, ByRef c2 As Long, ByRef isGrad As Boolean)
    isGrad = False
    If mLinks(l).HasColor Then
        c1 = mLinks(l).Fill: c2 = c1
    ElseIf IsHexColor(oLinkColor) Then
        c1 = HexToRGB(oLinkColor): c2 = c1
    Else
        Select Case LCase$(oLinkColor)
            Case "source": c1 = mNodes(mLinks(l).Src).Fill: c2 = c1
            Case "gray", "grey": c1 = RGB(170, 180, 184): c2 = c1
            Case "gradient"
                c1 = mNodes(mLinks(l).Src).Fill: c2 = mNodes(mLinks(l).Dst).Fill
                isGrad = (c1 <> c2)
            Case Else: c1 = mNodes(mLinks(l).Dst).Fill: c2 = c1
        End Select
    End If
End Sub

' Assigns node colors: explicit link colors first, then by flow / node / column
Private Sub AssignColors()
    Dim i As Long, l As Long, c As Long, cnt As Long, s As Long, a As Double, d As Long
    For i = 1 To mNodeCount
        mNodes(i).FillSet = False
    Next i
    For l = 1 To mLinkCount
        If mLinks(l).HasColor Then
            d = mLinks(l).Dst
            If Not mNodes(d).FillSet Then mNodes(d).Fill = mLinks(l).Fill: mNodes(d).FillSet = True
        End If
    Next l
    cnt = 0
    For c = 0 To mMaxCol
        For i = 1 To mNodeCount
            If mNodes(i).Col = c And Not mNodes(i).FillSet Then
                Select Case oColorBy
                    Case "column"
                        mNodes(i).Fill = PaletteColor(c)
                    Case "node"
                        mNodes(i).Fill = PaletteColor(cnt): cnt = cnt + 1
                    Case Else ' flow: keep the color of the main source when it sends most of its output here
                        s = DominantSource(i, a)
                        If s > 0 Then
                            If mNodes(s).FillSet And a >= 0.5 * mNodes(s).ValOut Then
                                mNodes(i).Fill = mNodes(s).Fill
                            Else
                                mNodes(i).Fill = PaletteColor(cnt): cnt = cnt + 1
                            End If
                        Else
                            mNodes(i).Fill = PaletteColor(cnt): cnt = cnt + 1
                        End If
                End Select
                mNodes(i).FillSet = True
            End If
        Next i
    Next c
End Sub

' Source node sending the largest amount to node i (0 if none); amt receives that amount
Private Function DominantSource(ByVal i As Long, ByRef amt As Double) As Long
    Dim l As Long
    amt = 0: DominantSource = 0
    For l = 1 To mLinkCount
        If mLinks(l).Dst = i And mLinks(l).Src <> i Then
            If mLinks(l).Amount > amt Then amt = mLinks(l).Amount: DominantSource = mLinks(l).Src
        End If
    Next l
End Function

' =====================================================================
'  Layout
' =====================================================================
Private Function NodeIndex(ByVal nm As String) As Long
    Dim i As Long
    For i = 1 To mNodeCount
        If StrComp(mNodes(i).Label, nm, vbTextCompare) = 0 Then NodeIndex = i: Exit Function
    Next i
    mNodeCount = mNodeCount + 1
    ReDim Preserve mNodes(1 To mNodeCount)
    mNodes(mNodeCount).Label = nm
    mNodes(mNodeCount).ColForced = -1
    NodeIndex = mNodeCount
End Function

Private Sub ComputeColumns()
    Dim i As Long, k As Long, changed As Boolean, maxCol As Long, l As Long, c As Long
    mTotalRef = 0
    For i = 1 To mNodeCount
        mNodes(i).Amount = mNodes(i).ValIn
        If mNodes(i).ValOut > mNodes(i).Amount Then mNodes(i).Amount = mNodes(i).ValOut
        mNodes(i).Col = 0
    Next i
    ' longest path from the sources (relaxation, bounded to survive cycles)
    For k = 1 To mNodeCount + 1
        changed = False
        For l = 1 To mLinkCount
            If mNodes(mLinks(l).Dst).Col < mNodes(mLinks(l).Src).Col + 1 Then
                mNodes(mLinks(l).Dst).Col = mNodes(mLinks(l).Src).Col + 1: changed = True
            End If
        Next l
        If Not changed Then Exit For
    Next k
    For i = 1 To mNodeCount
        If mNodes(i).ColForced >= 0 Then mNodes(i).Col = mNodes(i).ColForced
    Next i
    maxCol = 0
    For i = 1 To mNodeCount
        If mNodes(i).Col > maxCol Then maxCol = mNodes(i).Col
    Next i
    ' justify: sinks go to the last column; isolated sources move next to their targets
    If oAlignment <> "left" Then
        For i = 1 To mNodeCount
            If mNodes(i).ColForced < 0 Then
                If Not mNodes(i).HasOut Then
                    mNodes(i).Col = maxCol
                ElseIf Not mNodes(i).HasIn Then
                    c = maxCol
                    For l = 1 To mLinkCount
                        If mLinks(l).Src = i Then
                            If mNodes(mLinks(l).Dst).Col - 1 < c Then c = mNodes(mLinks(l).Dst).Col - 1
                        End If
                    Next l
                    If c < 0 Then c = 0
                    mNodes(i).Col = c
                End If
            End If
        Next i
    End If
    mMaxCol = 0
    For i = 1 To mNodeCount
        If mNodes(i).Col > mMaxCol Then mMaxCol = mNodes(i).Col
    Next i
    ' reference total for percentages: sum of the first column
    For i = 1 To mNodeCount
        If mNodes(i).Col = 0 Then mTotalRef = mTotalRef + mNodes(i).Amount
    Next i
    AssignColors
    ' label side: first column on the left of the node (outside), others on the right
    For i = 1 To mNodeCount
        mNodes(i).OnLeft = False
        If mNodes(i).Col = 0 And oOutsideLabels And mMaxCol > 0 Then mNodes(i).OnLeft = True
        If mNodes(i).Col = mMaxCol And mMaxCol > 0 And Not oOutsideLabels Then mNodes(i).OnLeft = True
    Next i
End Sub

' Measures the labels (creating the text boxes when shp is given), decides the box size and
' computes every position. Sets mBoxW / mBoxH to the box actually used.
Private Sub PrepareLayout(ByVal shp As Object, ByVal pLeft As Single, ByVal pTop As Single, _
                          ByVal pWidth As Single, ByVal pHeight As Single, ByVal pGrow As Boolean)
    Dim i As Long, marginL As Single, marginR As Single, sumReq As Single, needW As Single, needH As Single
    Dim nMax As Long, c As Long, n As Long, maxLH As Single, factor As Single, showLabels As Boolean, pass As Long, req() As Single
    BuildPalette
    ComputeColumns
    showLabels = (oLabels <> "none")
    mBandTitle = 0
    If Len(oTitle) > 0 Then mBandTitle = mBandTitle + oFontSize * 1.6 * 1.5
    If Len(oSubtitle) > 0 Then mBandTitle = mBandTitle + oFontSize * 1.1 * 1.4
    If mBandTitle > 0 Then mBandTitle = mBandTitle + 4
    mBandCol = 0
    If Len(oColumnTitles) > 0 Then mBandCol = oColumnTitleSize * 1.4 + 6
    For pass = 1 To 2
        MeasureLabels shp, pLeft, pTop, showLabels
        marginL = 0: marginR = 0: maxLH = 0
        If showLabels Then
            For i = 1 To mNodeCount
                If mNodes(i).OnLeft And mNodes(i).Col = 0 Then
                    If mNodes(i).LW + 6 > marginL Then marginL = mNodes(i).LW + 6
                ElseIf mNodes(i).Col = mMaxCol And Not mNodes(i).OnLeft Then
                    If mNodes(i).LW + 6 > marginR Then marginR = mNodes(i).LW + 6
                End If
                If mNodes(i).LH > maxLH Then maxLH = mNodes(i).LH
            Next i
        End If
        sumReq = RequiredSteps(req)
        needW = 2 * oMargin + marginL + marginR + sumReq + oNodeWidth
        mPadEff = oNodePadding
        If showLabels And maxLH + 3 > mPadEff Then mPadEff = maxLH + 3
        nMax = 0
        For c = 0 To mMaxCol
            n = 0
            For i = 1 To mNodeCount
                If mNodes(i).Col = c Then n = n + 1
            Next i
            If n > nMax Then nMax = n
        Next c
        needH = 2 * oMargin + mBandTitle + mBandCol + (nMax - 1) * mPadEff + nMax * 3 + 60
        If pass = 1 And Not pGrow And needW > pWidth + 0.5 And showLabels And oFontSize > 6 Then
            ' shrink the text once so that the diagram fits the given box
            factor = (pWidth - 2 * oMargin - oNodeWidth) / (needW - 2 * oMargin - oNodeWidth)
            If factor < 0.55 Then factor = 0.55
            oFontSize = oFontSize * factor
            If oFontSize < 6 Then oFontSize = 6
            oColumnTitleSize = oColumnTitleSize * factor
            If oColumnTitleSize < 6 Then oColumnTitleSize = 6
            DeleteLabels shp
        Else
            Exit For
        End If
    Next pass
    mBoxW = pWidth: mBoxH = pHeight
    If pGrow Then
        If needW > mBoxW Then mBoxW = needW
        If needH > mBoxH Then mBoxH = needH
    End If
    mAreaL = pLeft + oMargin + marginL
    mAreaT = pTop + oMargin + mBandTitle + mBandCol
    mAreaW = mBoxW - 2 * oMargin - marginL - marginR
    mAreaH = mBoxH - 2 * oMargin - mBandTitle - mBandCol
    If mAreaW < 40 Then mAreaW = 40
    If mAreaH < 40 Then mAreaH = 40
    If nMax > 1 Then
        If (nMax - 1) * mPadEff + nMax * 2 > mAreaH Then mPadEff = (mAreaH - nMax * 2) / (nMax - 1)
        If mPadEff < 2 Then mPadEff = 2
    End If
    ComputeLayout
    PlaceLabels showLabels
End Sub

' Horizontal step required after each column: node width + room for the labels drawn on the
' right of that column; short steps are raised to 70 % of the longest one so that the columns
' look evenly spaced. Returns the sum of the steps.
Private Function RequiredSteps(ByRef req() As Single) As Single
    Dim c As Long, i As Long, w As Single, total As Single, mx As Single
    ReDim req(0 To mMaxCol)
    For c = 0 To mMaxCol - 1
        req(c) = oNodeWidth + 34
        w = 0
        For i = 1 To mNodeCount
            If mNodes(i).Col = c And Not mNodes(i).OnLeft Then
                If mNodes(i).LW > w Then w = mNodes(i).LW
            End If
        Next i
        If w > 0 Then
            If oNodeWidth + w + 14 > req(c) Then req(c) = oNodeWidth + w + 14
        End If
        If req(c) > mx Then mx = req(c)
    Next c
    For c = 0 To mMaxCol - 1
        If req(c) < 0.7 * mx Then req(c) = 0.7 * mx
        total = total + req(c)
    Next c
    RequiredSteps = total
End Function

' Label sizes: real text boxes when a Shapes collection is given, estimates otherwise
Private Sub MeasureLabels(ByVal shp As Object, ByVal pLeft As Single, ByVal pTop As Single, ByVal showLabels As Boolean)
    Dim i As Long, l1 As String, l2 As String, nLines As Long, maxW As Single, w As Single
    maxW = oFontSize * 13
    For i = 1 To mNodeCount
        mNodes(i).LW = 0: mNodes(i).LH = 0: mNodes(i).LblName = ""
        If showLabels Then
            If shp Is Nothing Then
                l1 = LabelLine1(i): l2 = "": nLines = 1
                If oLabels = "name_value" And oLabelLines = 2 Then l2 = FormatAmount(mNodes(i).Amount): nLines = 2
                w = TextWidth(l1)
                If w > maxW Then
                    nLines = nLines + Int(w / maxW)
                    w = maxW
                End If
                If TextWidth(l2) > w Then w = TextWidth(l2)
                mNodes(i).LW = w + 4
                mNodes(i).LH = oFontSize * 1.25 * nLines + 4
            Else
                mNodes(i).LblName = CreateLabel(shp, i, pLeft, pTop, maxW)
            End If
        End If
    Next i
End Sub

Private Sub DeleteLabels(ByVal shp As Object)
    Dim i As Long
    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    For i = 1 To mNodeCount
        If Len(mNodes(i).LblName) > 0 Then shp.Item(mNodes(i).LblName).Delete
        mNodes(i).LblName = ""
    Next i
    On Error GoTo 0
End Sub

Private Sub ComputeLayout()
    Dim i As Long, l As Long, c As Long, maxCol As Long, pass As Long
    Dim colTotal() As Double, colCount() As Long, unitScale As Double, sc As Double
    Dim idx() As Long, j As Long, t As Long, req() As Single, sumReq As Single, avail As Single, extra As Single, ratio As Single
    maxCol = mMaxCol
    ReDim colTotal(0 To maxCol): ReDim colCount(0 To maxCol)
    For i = 1 To mNodeCount
        c = mNodes(i).Col
        colTotal(c) = colTotal(c) + mNodes(i).Amount: colCount(c) = colCount(c) + 1
    Next i
    unitScale = 0
    For c = 0 To maxCol
        If colTotal(c) > 0 Then
            sc = (mAreaH - mPadEff * (colCount(c) - 1)) / colTotal(c)
            If unitScale = 0 Or sc < unitScale Then unitScale = sc
        End If
    Next c
    If unitScale <= 0 Then unitScale = 0.001
    ' horizontal positions: each step leaves room for the labels drawn on the right of its column
    ReDim mColX(0 To maxCol)
    sumReq = RequiredSteps(req)
    avail = mAreaW - oNodeWidth
    ratio = 1: extra = 0
    If maxCol > 0 Then
        If sumReq > avail And sumReq > 0 Then ratio = avail / sumReq Else extra = (avail - sumReq) / maxCol
    End If
    mColX(0) = mAreaL
    For c = 1 To maxCol
        mColX(c) = mColX(c - 1) + req(c - 1) * ratio + extra
    Next c
    For i = 1 To mNodeCount
        mNodes(i).Rank = i
    Next i
    If oOrder <> "table" Then
        For pass = 1 To 2
            AssignY unitScale, maxCol
            For c = 1 To maxCol
                For i = 1 To mNodeCount
                    If mNodes(i).Col = c Then mNodes(i).Rank = Barycenter(i, True)
                Next i
            Next c
            AssignY unitScale, maxCol
            For c = maxCol - 1 To 0 Step -1
                For i = 1 To mNodeCount
                    If mNodes(i).Col = c Then mNodes(i).Rank = Barycenter(i, False)
                Next i
            Next c
        Next pass
    End If
    AssignY unitScale, maxCol
    For i = 1 To mNodeCount
        mNodes(i).OutOff = 0: mNodes(i).InOff = 0
    Next i
    For l = 1 To mLinkCount
        mLinks(l).H = mLinks(l).Amount * unitScale
        If mLinks(l).H < 0.5 Then mLinks(l).H = 0.5
    Next l
    ReDim idx(1 To mLinkCount)
    For l = 1 To mLinkCount
        idx(l) = l
    Next l
    For i = 1 To mLinkCount - 1
        For j = i + 1 To mLinkCount
            If LinkKeyOut(idx(j)) < LinkKeyOut(idx(i)) Then t = idx(i): idx(i) = idx(j): idx(j) = t
        Next j
    Next i
    For i = 1 To mLinkCount
        l = idx(i)
        mLinks(l).Y0 = mNodes(mLinks(l).Src).Y + mNodes(mLinks(l).Src).OutOff
        mNodes(mLinks(l).Src).OutOff = mNodes(mLinks(l).Src).OutOff + mLinks(l).H
    Next i
    For i = 1 To mLinkCount - 1
        For j = i + 1 To mLinkCount
            If LinkKeyIn(idx(j)) < LinkKeyIn(idx(i)) Then t = idx(i): idx(i) = idx(j): idx(j) = t
        Next j
    Next i
    For i = 1 To mLinkCount
        l = idx(i)
        mLinks(l).Y1 = mNodes(mLinks(l).Dst).Y + mNodes(mLinks(l).Dst).InOff
        mNodes(mLinks(l).Dst).InOff = mNodes(mLinks(l).Dst).InOff + mLinks(l).H
    Next i
End Sub

Private Function LinkKeyOut(ByVal l As Long) As Double
    LinkKeyOut = mLinks(l).Src * 1000000# + mNodes(mLinks(l).Dst).Y + mNodes(mLinks(l).Dst).H / 2
End Function

Private Function LinkKeyIn(ByVal l As Long) As Double
    LinkKeyIn = mLinks(l).Dst * 1000000# + mNodes(mLinks(l).Src).Y + mNodes(mLinks(l).Src).H / 2
End Function

Private Function Barycenter(ByVal i As Long, ByVal fromSources As Boolean) As Double
    Dim l As Long, s As Double, w As Double, o As Long
    For l = 1 To mLinkCount
        If fromSources Then
            If mLinks(l).Dst = i Then o = mLinks(l).Src: s = s + (mNodes(o).Y + mNodes(o).H / 2) * mLinks(l).Amount: w = w + mLinks(l).Amount
        Else
            If mLinks(l).Src = i Then o = mLinks(l).Dst: s = s + (mNodes(o).Y + mNodes(o).H / 2) * mLinks(l).Amount: w = w + mLinks(l).Amount
        End If
    Next l
    If w > 0 Then Barycenter = s / w Else Barycenter = mNodes(i).Y + mNodes(i).H / 2
End Function

Private Sub AssignY(ByVal unitScale As Double, ByVal maxCol As Long)
    Dim c As Long, i As Long, j As Long, t As Long, ord() As Long, n As Long, total As Single, y As Single
    For c = 0 To maxCol
        n = 0
        ReDim ord(1 To mNodeCount)
        For i = 1 To mNodeCount
            If mNodes(i).Col = c Then n = n + 1: ord(n) = i
        Next i
        If n > 0 Then
            For i = 1 To n - 1
                For j = i + 1 To n
                    If mNodes(ord(j)).Rank < mNodes(ord(i)).Rank Then t = ord(i): ord(i) = ord(j): ord(j) = t
                Next j
            Next i
            total = 0
            For i = 1 To n
                mNodes(ord(i)).H = mNodes(ord(i)).Amount * unitScale
                If mNodes(ord(i)).H < 2 Then mNodes(ord(i)).H = 2
                total = total + mNodes(ord(i)).H
            Next i
            total = total + mPadEff * (n - 1)
            y = mAreaT + (mAreaH - total) / 2
            If y < mAreaT Then y = mAreaT
            For i = 1 To n
                mNodes(ord(i)).X = mColX(c)
                mNodes(ord(i)).Y = y
                y = y + mNodes(ord(i)).H + mPadEff
            Next i
        End If
    Next c
End Sub

' Label positions: centred on the node, pushed apart when two labels of a column would overlap
Private Sub PlaceLabels(ByVal showLabels As Boolean)
    Dim c As Long, i As Long, j As Long, t As Long, ord() As Long, n As Long, bottom As Single, shift As Single, firstTop As Single
    If Not showLabels Then Exit Sub
    For i = 1 To mNodeCount
        If mNodes(i).OnLeft Then mNodes(i).LX = mNodes(i).X - 5 - mNodes(i).LW Else mNodes(i).LX = mNodes(i).X + oNodeWidth + 5
        mNodes(i).LY = mNodes(i).Y + mNodes(i).H / 2 - mNodes(i).LH / 2
    Next i
    For c = 0 To mMaxCol
        n = 0
        ReDim ord(1 To mNodeCount)
        For i = 1 To mNodeCount
            If mNodes(i).Col = c Then n = n + 1: ord(n) = i
        Next i
        If n > 1 Then
            For i = 1 To n - 1
                For j = i + 1 To n
                    If mNodes(ord(j)).LY < mNodes(ord(i)).LY Then t = ord(i): ord(i) = ord(j): ord(j) = t
                Next j
            Next i
            bottom = -999999
            For i = 1 To n
                If mNodes(ord(i)).LY < bottom + 1 Then mNodes(ord(i)).LY = bottom + 1
                bottom = mNodes(ord(i)).LY + mNodes(ord(i)).LH
            Next i
            shift = bottom - (mAreaT + mAreaH + 4)
            If shift > 0 Then
                firstTop = mNodes(ord(1)).LY - shift
                If firstTop < mAreaT - mBandCol Then shift = mNodes(ord(1)).LY - (mAreaT - mBandCol)
                If shift > 0 Then
                    For i = 1 To n
                        mNodes(ord(i)).LY = mNodes(ord(i)).LY - shift
                    Next i
                End If
            End If
        End If
    Next c
End Sub

' Estimated text width in points (Arial-like metrics), used when no text box can be measured
Private Function TextWidth(ByVal s As String) As Single
    Dim i As Long, ch As String, w As Single
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        Select Case ch
            Case "i", "j", "l", "t", "f", "I", ".", ",", ":", ";", "'", "|", "!": w = w + 0.28
            Case " ", "(", ")", "-", "r", "[", "]": w = w + 0.34
            Case "m", "w", "M", "W", "@": w = w + 0.85
            Case "A" To "Z": w = w + 0.68
            Case "0" To "9": w = w + 0.56
            Case Else: w = w + 0.52
        End Select
    Next i
    TextWidth = w * oFontSize + 2
End Function

Private Function LabelLine1(ByVal i As Long) As String
    Select Case oLabels
        Case "name": LabelLine1 = mNodes(i).Label
        Case "value": LabelLine1 = FormatAmount(mNodes(i).Amount)
        Case Else
            If oLabelLines = 1 Then LabelLine1 = mNodes(i).Label & " (" & FormatAmount(mNodes(i).Amount) & ")" Else LabelLine1 = mNodes(i).Label
    End Select
End Function

Private Function Clamp(ByVal v As Double, ByVal lo As Double, ByVal hi As Double) As Double
    If v < lo Then v = lo
    If v > hi Then v = hi
    Clamp = v
End Function

Private Function YesNo(ByVal v As String, ByVal dflt As Boolean) As Boolean
    Select Case LCase$(Trim$(v))
        Case "yes", "y", "oui", "o", "1", "true", "on": YesNo = True
        Case "no", "n", "non", "0", "false", "off": YesNo = False
        Case Else: YesNo = dflt
    End Select
End Function

' =====================================================================
'  Drawing
' =====================================================================
Private Function DrawBackground(ByVal shp As Object, ByVal x As Single, ByVal y As Single, ByVal w As Single, ByVal h As Single) As String
    Dim s As Object
    Set s = shp.AddShape(C_msoShapeRoundedRectangle, x, y, w, h)
    With s
        .Name = "SkBackground_" & mStamp
        On Error Resume Next
        .Adjustments(1) = 0.03
        On Error GoTo 0
        .Fill.Visible = C_msoTrue: .Fill.Solid: .Fill.ForeColor.RGB = oBackgroundColor: .Fill.Transparency = 0
        .Line.Visible = C_msoFalse
        On Error Resume Next
        .Shadow.Visible = C_msoFalse
        .ZOrder C_msoSendToBack
        On Error GoTo 0
    End With
    DrawBackground = s.Name
End Function

Private Function DrawLink(ByVal shp As Object, ByVal l As Long, ByVal n As Long) As String
    Dim ff As Object, s As Object, x0 As Single, x1 As Single, y0 As Single, y1 As Single, h As Single, cx0 As Single, cx1 As Single
    Dim c1 As Long, c2 As Long, isGrad As Boolean, tr As Single
    x0 = mNodes(mLinks(l).Src).X + oNodeWidth
    x1 = mNodes(mLinks(l).Dst).X
    y0 = mLinks(l).Y0: y1 = mLinks(l).Y1: h = mLinks(l).H
    Set ff = shp.BuildFreeform(C_msoEditingCorner, x0, y0)
    If oLinkShape = "straight" Then
        ff.AddNodes C_msoSegmentLine, C_msoEditingAuto, x1, y1
        ff.AddNodes C_msoSegmentLine, C_msoEditingAuto, x1, y1 + h
        ff.AddNodes C_msoSegmentLine, C_msoEditingAuto, x0, y0 + h
    Else
        cx0 = x0 + (x1 - x0) * oCurvature
        cx1 = x1 - (x1 - x0) * oCurvature
        ff.AddNodes C_msoSegmentCurve, C_msoEditingCorner, cx0, y0, cx1, y1, x1, y1
        ff.AddNodes C_msoSegmentLine, C_msoEditingAuto, x1, y1 + h
        ff.AddNodes C_msoSegmentCurve, C_msoEditingCorner, cx1, y1 + h, cx0, y0 + h, x0, y0 + h
    End If
    ff.AddNodes C_msoSegmentLine, C_msoEditingAuto, x0, y0
    Set s = ff.ConvertToShape
    LinkColors l, c1, c2, isGrad
    tr = 1 - oLinkOpacity
    With s
        .Name = "SkLink" & n & "_" & mStamp
        .Fill.Visible = C_msoTrue
        If isGrad Then
            On Error Resume Next
            .Fill.TwoColorGradient C_msoGradientVertical, 1
            .Fill.GradientAngle = 0
            .Fill.GradientStops(1).Color.RGB = c1
            .Fill.GradientStops(2).Color.RGB = c2
            .Fill.GradientStops(1).Transparency = tr
            .Fill.GradientStops(2).Transparency = tr
            If Err.Number <> 0 Then
                Err.Clear
                .Fill.Solid: .Fill.ForeColor.RGB = c1: .Fill.Transparency = tr
            End If
            On Error GoTo 0
        Else
            .Fill.Solid
            .Fill.ForeColor.RGB = c1
            .Fill.Transparency = tr
        End If
        If oLinkOutline Then
            .Line.Visible = C_msoTrue: .Line.ForeColor.RGB = RGB(255, 255, 255): .Line.Weight = 0.75
        Else
            .Line.Visible = C_msoFalse
        End If
        On Error Resume Next
        .Shadow.Visible = C_msoFalse
        On Error GoTo 0
    End With
    DrawLink = s.Name
End Function

Private Function DrawNode(ByVal shp As Object, ByVal i As Long, ByVal n As Long) As String
    Dim s As Object
    If oRoundedNodes Then
        Set s = shp.AddShape(C_msoShapeRoundedRectangle, mNodes(i).X, mNodes(i).Y, oNodeWidth, mNodes(i).H)
        On Error Resume Next
        s.Adjustments(1) = 0.35
        On Error GoTo 0
    Else
        Set s = shp.AddShape(C_msoShapeRectangle, mNodes(i).X, mNodes(i).Y, oNodeWidth, mNodes(i).H)
    End If
    With s
        .Name = "SkNode" & n & "_" & mStamp
        .Fill.Visible = C_msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = NodeFillColor(i)
        .Fill.Transparency = 0
        .Line.Visible = C_msoFalse
        On Error Resume Next
        .Shadow.Visible = C_msoFalse
        On Error GoTo 0
    End With
    DrawNode = s.Name
End Function

' Creates the label text box of a node (position is set later) and records its measured size
Private Function CreateLabel(ByVal shp As Object, ByVal i As Long, ByVal x As Single, ByVal y As Single, ByVal maxW As Single) As String
    Dim s As Object, txt As String, w As Single, h As Single, nLines As Long, l1 As String, l2 As String, wrapIt As Boolean
    l1 = LabelLine1(i): l2 = ""
    nLines = 1
    If oLabels = "name_value" And oLabelLines = 2 Then l2 = FormatAmount(mNodes(i).Amount): nLines = 2
    txt = l1
    If nLines = 2 Then txt = txt & vbCr & l2
    w = TextWidth(l1)
    wrapIt = (w > maxW)
    If wrapIt Then w = maxW
    If TextWidth(l2) > w Then w = TextWidth(l2)
    h = oFontSize * 1.3 * nLines + 4
    Set s = shp.AddTextbox(C_msoTextOrientationHorizontal, x, y, w + 4, h)
    With s
        .Name = "SkLabel" & i & "_" & mStamp
        If oLabelBackground Then
            .Fill.Visible = C_msoTrue: .Fill.Solid: .Fill.ForeColor.RGB = RGB(255, 255, 255): .Fill.Transparency = 0.25
        Else
            .Fill.Visible = C_msoFalse
        End If
        .Line.Visible = C_msoFalse
        On Error Resume Next
        .Shadow.Visible = C_msoFalse
        On Error GoTo 0
        With .TextFrame2
            .MarginLeft = 2: .MarginRight = 2: .MarginTop = 1: .MarginBottom = 1
            If wrapIt Then .WordWrap = C_msoTrue Else .WordWrap = C_msoFalse
            .AutoSize = C_msoAutoSizeShapeToFitText
            .TextRange.Text = txt
            .TextRange.Font.Name = oFont
            .TextRange.Font.Size = oFontSize
            .TextRange.Font.Fill.ForeColor.RGB = oTextColor
            .TextRange.Font.Bold = C_msoFalse
            .TextRange.Font.Italic = C_msoFalse
            If nLines = 2 Then
                .TextRange.Paragraphs(2).Font.Bold = C_msoTrue
                .TextRange.Paragraphs(2).Font.Fill.ForeColor.RGB = NodeFillColor(i)
            ElseIf oLabels = "value" Then
                .TextRange.Font.Bold = C_msoTrue
            End If
            If mNodes(i).OnLeft Then .TextRange.ParagraphFormat.Alignment = C_msoAlignRight Else .TextRange.ParagraphFormat.Alignment = C_msoAlignLeft
            .TextRange.ParagraphFormat.SpaceBefore = 0: .TextRange.ParagraphFormat.SpaceAfter = 0
        End With
        mNodes(i).LW = .Width
        mNodes(i).LH = .Height
    End With
    CreateLabel = s.Name
End Function

Private Function DrawLinkValue(ByVal shp As Object, ByVal l As Long, ByVal n As Long) As String
    Dim s As Object, x As Single, y As Single, txt As String, w As Single, h As Single, cx As Single, cy As Single
    txt = FormatAmount(mLinks(l).Amount)
    w = TextWidth(txt): h = oFontSize * 1.3 + 2
    cx = (mNodes(mLinks(l).Src).X + oNodeWidth + mNodes(mLinks(l).Dst).X) / 2
    cy = (mLinks(l).Y0 + mLinks(l).Y1) / 2 + mLinks(l).H / 2
    x = cx - w / 2: y = cy - h / 2
    Set s = shp.AddTextbox(C_msoTextOrientationHorizontal, x, y, w, h)
    With s
        .Name = "SkValue" & n & "_" & mStamp
        .Fill.Visible = C_msoFalse: .Line.Visible = C_msoFalse
        With .TextFrame2
            .WordWrap = C_msoFalse
            .AutoSize = C_msoAutoSizeShapeToFitText
            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
            .TextRange.Text = txt
            .TextRange.Font.Name = oFont
            .TextRange.Font.Size = oFontSize * 0.9
            .TextRange.Font.Fill.ForeColor.RGB = oTextColor
            .TextRange.ParagraphFormat.Alignment = C_msoAlignCenter
        End With
        .Left = cx - .Width / 2
        .Top = cy - .Height / 2
    End With
    DrawLinkValue = s.Name
End Function

Private Function DrawColumnTitle(ByVal shp As Object, ByVal c As Long, ByVal txt As String, ByVal y As Single, ByVal n As Long) As String
    Dim s As Object, cx As Single, w As Single
    cx = mColX(c) + oNodeWidth / 2
    w = TextWidth(txt) + 10
    Set s = shp.AddTextbox(C_msoTextOrientationHorizontal, cx - w / 2, y, w, oColumnTitleSize * 1.4)
    With s
        .Name = "SkColumnTitle" & n & "_" & mStamp
        .Fill.Visible = C_msoFalse: .Line.Visible = C_msoFalse
        With .TextFrame2
            .WordWrap = C_msoFalse
            .AutoSize = C_msoAutoSizeShapeToFitText
            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
            .TextRange.Text = txt
            .TextRange.Font.Name = oFont
            .TextRange.Font.Size = oColumnTitleSize
            .TextRange.Font.Bold = C_msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(110, 120, 124)
            .TextRange.ParagraphFormat.Alignment = C_msoAlignCenter
        End With
        .Left = cx - .Width / 2
    End With
    DrawColumnTitle = s.Name
End Function

Private Function DrawTitle(ByVal shp As Object, ByVal txt As String, ByVal x As Single, ByVal y As Single, _
                           ByVal w As Single, ByVal pSize As Single, ByVal isSub As Boolean, ByVal n As Long) As String
    Dim s As Object
    Set s = shp.AddTextbox(C_msoTextOrientationHorizontal, x, y, w, pSize * 1.5)
    With s
        .Name = "SkTitle" & n & "_" & mStamp
        .Fill.Visible = C_msoFalse: .Line.Visible = C_msoFalse
        With .TextFrame2
            .WordWrap = C_msoTrue
            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
            .TextRange.Text = txt
            .TextRange.Font.Name = oFont
            .TextRange.Font.Size = pSize
            If isSub Then
                .TextRange.Font.Bold = C_msoFalse: .TextRange.Font.Italic = C_msoTrue
                .TextRange.Font.Fill.ForeColor.RGB = RGB(110, 120, 124)
            Else
                .TextRange.Font.Bold = C_msoTrue: .TextRange.Font.Italic = C_msoFalse
                .TextRange.Font.Fill.ForeColor.RGB = oTextColor
            End If
            .TextRange.ParagraphFormat.Alignment = C_msoAlignLeft
        End With
    End With
    DrawTitle = s.Name
End Function

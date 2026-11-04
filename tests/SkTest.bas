Option Explicit
' Test module (not shipped): compiled and executed in LibreOffice by build/test_layout.py
Public Function SkTestRun() As String
    Dim s As String
    SkReset
    SkApplyOptionText "palette = Vivid" & vbCr & "link_color = gradient" & vbCr & "column_titles = Revenue;Gross;Operating;Result" & vbCr & "value_format = k" & vbCr & "decimals = 1" & vbCr & "unit = CHF" & vbCr & "link_opacity = 60"
    SkAddLink "Revenue", "Cost of sales", 450, "#E34948"
    SkAddLink "Revenue", "Gross profit", 550, "#1BAF7A"
    SkAddLink "Gross profit", "Operating expenses", 300, "#E34948"
    SkAddLink "Gross profit", "Operating profit", 250, "#1BAF7A"
    SkAddLink "Operating expenses", "R&D", 120, "#E34948"
    SkAddLink "Operating expenses", "Sales & marketing", 110, "#E34948"
    SkAddLink "Operating expenses", "General & admin", 70, "#E34948"
    SkAddLink "Operating profit", "Taxes", 60, "#E34948"
    SkAddLink "Operating profit", "Net profit", 190, "#1BAF7A"
    s = "links=" & SkLinkCount & vbLf & SkDebugLayout(600, 360)
    s = s & "parse=" & ParseNumber("3'000") & "/" & ParseNumber("2 650,5") & "/" & ParseNumber("1,234.5") & "/" & ParseNumber(" 42 ") & vbLf
    s = s & "hex=" & IsHexColor("#E34948") & "/" & IsHexColor("red") & "/" & HexToRGB("#FF0000") & vbLf
    SkReset
    s = s & "fmt=" & FormatAmount(3000) & "/" & FormatAmount(2.5) & vbLf
    SkTestRun = s
End Function

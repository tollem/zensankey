Option Explicit
' Test module (not shipped): compiled and executed in LibreOffice by build/test_layout.py.
' Prints the computed layouts (used by build/preview_svg.py to render previews).
Public Function SkTestRun() As String
    Dim s As String
    SkResetStyle
    SkReset
    s = "### fif default" & vbLf
    SkSetOption "column_titles", "1. Financement;2. Mise a disposition;3. Octrois;4. Devenir des prets;5. Participations;6. Resultat"
    SkSetOption "title", "Cycle de financement du FIF"
    SkSetOption "subtitle", "Montants illustratifs (kCHF)"
    LoadFif
    s = s & SkDebugLayout(640, 360, True)
    s = s & "### fif fixed box (PowerPoint)" & vbLf
    SkReset
    LoadFif
    s = s & SkDebugLayout(880, 420, False)
    s = s & "### income statement" & vbLf
    SkResetStyle
    SkReset
    SkSetOption "value_format", "k": SkSetOption "decimals", "1": SkSetOption "unit", "CHF"
    SkAddLink "Revenue", "Cost of sales", 450, "#E34948"
    SkAddLink "Revenue", "Gross profit", 550, "#1BAF7A"
    SkAddLink "Gross profit", "Operating expenses", 300, "#E34948"
    SkAddLink "Gross profit", "Operating profit", 250, "#1BAF7A"
    SkAddLink "Operating expenses", "R&D", 120, "#E34948"
    SkAddLink "Operating expenses", "Sales & marketing", 110, "#E34948"
    SkAddLink "Operating expenses", "General & admin", 70, "#E34948"
    SkAddLink "Operating profit", "Taxes", 60, "#E34948"
    SkAddLink "Operating profit", "Net profit", 190, "#1BAF7A"
    s = s & SkDebugLayout(640, 360, True)
    s = s & "### readme" & vbLf
    SkResetStyle
    SkReset
    SkSetOption "title", "Income statement": SkSetOption "subtitle", "Example: where the revenue goes (amounts in kCHF)"
    SkSetOption "column_titles", "Revenue;Gross profit;Operating;Result"
    SkAddLink "Revenue", "Cost of sales", 450, "#E34948"
    SkAddLink "Revenue", "Gross profit", 550, "#1BAF7A"
    SkAddLink "Gross profit", "Operating expenses", 300, "#E34948"
    SkAddLink "Gross profit", "Operating profit", 250, "#1BAF7A"
    SkAddLink "Operating expenses", "R&D", 120, "#E34948"
    SkAddLink "Operating expenses", "Sales & marketing", 110, "#E34948"
    SkAddLink "Operating expenses", "General & admin", 70, "#E34948"
    SkAddLink "Operating profit", "Taxes", 60, "#E34948"
    SkAddLink "Operating profit", "Net profit", 190, "#1BAF7A"
    s = s & SkDebugLayout(640, 380, True)
    s = s & "### checks" & vbLf
    s = s & "parse=" & ParseNumber("3'000") & "/" & ParseNumber("2 650,5") & "/" & ParseNumber("1,234.5") & "/" & ParseNumber(" 42 ") & vbLf
    s = s & "hex=" & IsHexColor("#E34948") & "/" & IsHexColor("red") & "/" & HexToRGB("#FF0000") & "/" & RGBToHex(HexToRGB("#1BAF7A")) & vbLf
    s = s & "ribbon=" & SkRibbonIndex("ddPalette") & "/" & SkRibbonIndex("ddOpacity") & "/" & SkRibbonIndex("ddShape") & "/" & SkRibbonKey("cbRounded") & vbLf
    SkRibbonSelect "ddShape", 1
    SkRibbonSelect "ddOpacity", 4
    SkRibbonSetPressed "cbRounded", True
    SkRibbonSetText "ebTitle", "Hello"
    s = s & "after=" & SkGetOption("curvature") & "/" & SkGetOption("link_opacity") & "/" & SkGetOption("rounded_nodes") & "/" & SkRibbonText("ebTitle") & "/" & SkRibbonIndex("ddShape") & vbLf
    SkResetStyle
    s = s & "settings=" & Len(SkSettingsText()) & vbLf
    s = s & "fmt=" & FormatAmount(3000) & "/" & FormatAmount(2.5) & vbLf
    SkTestRun = s
End Function

Private Sub LoadFif()
    SkAddLink "Etat de Geneve (ligne de financement)", "Liquidites FIF", 3000
    SkAddLink "Liquidites FIF", "Prets aux start-ups", 3000
    SkAddLink "Prets aux start-ups", "Prets rembourses", 2000
    SkAddLink "Prets aux start-ups", "Perte sur pret", 500
    SkAddLink "Prets aux start-ups", "Convertis en participations", 500
    SkAddLink "Prets rembourses", "Tresorerie recuperee par le FIF", 2000
    SkAddLink "Convertis en participations", "Exit avec moins-value", 250
    SkAddLink "Convertis en participations", "Exit au pair", 100
    SkAddLink "Convertis en participations", "Exit avec plus-value", 150
    SkAddLink "Exit avec moins-value", "Tresorerie recuperee par le FIF", 200
    SkAddLink "Exit avec moins-value", "Pertes constatees", 50
    SkAddLink "Exit au pair", "Tresorerie recuperee par le FIF", 100
    SkAddLink "Exit avec plus-value", "Tresorerie recuperee par le FIF", 150
    SkAddLink "Plus-value realisee (acquereur)", "Tresorerie recuperee par le FIF", 200
    SkAddLink "Perte sur pret", "Pertes constatees", 500
End Sub

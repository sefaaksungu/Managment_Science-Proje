Attribute VB_Name = "Module9"
Option Explicit

Sub chart1()
Attribute chart1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' chart1 Makro
'

'
    ActiveWindow.ScrollRow = 2
    ActiveWindow.ScrollRow = 3
    ActiveWindow.ScrollRow = 4
    ActiveWindow.ScrollRow = 5
    ActiveWindow.ScrollRow = 6
    ActiveWindow.ScrollRow = 7
    ActiveWindow.ScrollRow = 8
    Range("V17:V21").Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("'Amaç F. ve Kýsýtlar'!$V$17:$V$21")
    ActiveChart.Axes(xlValue).Select
    Selection.Delete
    ActiveSheet.ChartObjects("Grafik 2").Activate
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Açýlacak Fabrikalar"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Açýlacak Fabrikalar"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 19).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 19).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    ActiveChart.ChartArea.Select
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.SetElement (msoElementDataTableWithLegendKeys)
    ActiveChart.SetElement (msoElementDataTableNone)
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleNone)
    ActiveChart.ClearToMatchStyle
    ActiveChart.ChartStyle = 203
    ActiveChart.SetElement (msoElementDataTableWithLegendKeys)
    ActiveSheet.ChartObjects("Grafik 2").Activate
    ActiveChart.DataTable.Select
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisShow)
    ActiveChart.SetElement (msoElementPrimaryValueAxisShow)
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesMajor)
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveChart.SetElement (msoElementLegendLeft)
    ActiveChart.SetElement (msoElementLegendRight)
    ActiveChart.SetElement (msoElementLegendTop)
    ActiveChart.SetElement (msoElementLegendNone)
    ActiveChart.ChartArea.Select
    ActiveChart.DataTable.Select
    Range("V10").Select
    ActiveSheet.ChartObjects("Grafik 2").Activate
    ActiveChart.Parent.Cut
    Sheets("Karar Destek Sistemi").Select
    ActiveSheet.Paste
    ActiveSheet.ChartObjects("Grafik 17").Activate
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Grafik 17").IncrementLeft 282.75
    ActiveSheet.Shapes("Grafik 17").IncrementTop -240
    ActiveSheet.Shapes("Grafik 17").ScaleHeight 0.6701388889, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Grafik 17").ScaleHeight 1.1502590674, msoFalse, _
        msoScaleFromTopLeft
    ActiveChart.DataTable.Select
    ActiveChart.PlotArea.Select
    ActiveChart.DataTable.Select
    ActiveChart.FullSeriesCollection(3).Select
    ActiveChart.ChartArea.Select
    ActiveChart.DataTable.Select
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Grafik 17").ScaleHeight 0.954954955, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Grafik 17").ScaleWidth 0.8770833333, msoFalse, _
        msoScaleFromBottomRight
    ActiveSheet.Shapes("Grafik 17").ScaleWidth 0.8456054513, msoFalse, _
        msoScaleFromBottomRight
    ActiveChart.ChartTitle.Select
    With Selection.Format.TextFrame2.TextRange.Font
        .NameComplexScript = "Times New Roman"
        .NameFarEast = "Times New Roman"
        .Name = "Times New Roman"
    End With
    ActiveChart.ChartTitle.Text = "Açýlacak Fabrikalar"
    Selection.Left = 59.952
    Selection.Top = 6
    Range("M8").Select
    ActiveSheet.ChartObjects("Grafik 17").Activate
    ActiveChart.PlotArea.Select
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Grafik 17").IncrementLeft 3.75
    ActiveSheet.Shapes("Grafik 17").IncrementTop 1.5
    Range("N4").Select
    Sheets("Amaç F. ve Kýsýtlar").Select
    Range("V25:V29").Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("'Amaç F. ve Kýsýtlar'!$V$25:$V$29")
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisShow)
    ActiveChart.SetElement (msoElementPrimaryValueAxisShow)
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.SetElement (msoElementDataTableWithLegendKeys)
    ActiveChart.SetElement (msoElementLegendNone)
    ActiveChart.SetElement (msoElementLegendRight)
    ActiveChart.SetElement (msoElementLegendNone)
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Açýlacak Daðýtým Merkezleri"
    Selection.Format.TextFrame2.TextRange.Characters.Text = _
        "Açýlacak Daðýtým Merkezleri"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 27).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 8).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(9, 1).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(10, 18).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    Selection.Left = 112.53
    Selection.Top = 9
    With Selection.Format.TextFrame2.TextRange.Font
        .NameComplexScript = "Times New Roman"
        .NameFarEast = "Times New Roman"
        .Name = "Times New Roman"
    End With
    ActiveChart.ChartTitle.Text = "Açýlacak Daðýtým Merkezleri"
    Selection.Left = 121.53
    Selection.Top = 9
    ActiveChart.ChartArea.Select
    ActiveChart.Parent.Cut
    Sheets("Karar Destek Sistemi").Select
    ActiveSheet.Paste
    ActiveSheet.ChartObjects("Grafik 18").Activate
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Grafik 18").IncrementLeft -192
    ActiveSheet.Shapes("Grafik 18").IncrementTop 21
    ActiveChart.PlotArea.Select
    ActiveChart.ClearToMatchStyle
    ActiveChart.ChartStyle = 203
    ActiveChart.ChartArea.Select
    ActiveChart.ChartTitle.Select
    Application.CommandBars("Format Object").Visible = False
    With Selection.Format.TextFrame2.TextRange.Font
        .NameComplexScript = "Times New Roman"
        .NameFarEast = "Times New Roman"
        .Name = "Times New Roman"
    End With
    ActiveChart.ChartTitle.Text = "Açýlacak Daðýtým Merkezleri"
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Grafik 18").IncrementLeft -43.5
    ActiveSheet.Shapes("Grafik 18").IncrementTop -39.75
    ActiveSheet.Shapes("Grafik 18").ScaleHeight 0.7743055556, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Grafik 18").ScaleWidth 0.825, msoFalse, msoScaleFromTopLeft
    ActiveChart.Axes(xlValue).Select
    Selection.Delete
    ActiveSheet.ChartObjects("Grafik 18").Activate
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Grafik 18").ScaleWidth 1.0277777778, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Grafik 18").IncrementLeft 83.25
    ActiveSheet.Shapes("Grafik 18").IncrementTop 63
    ActiveSheet.Shapes("Grafik 18").ScaleWidth 0.9262901842, msoFalse, _
        msoScaleFromTopLeft
    ActiveSheet.Shapes("Grafik 18").IncrementLeft 243
    ActiveSheet.Shapes("Grafik 18").IncrementTop 72
    Range("M11").Select
End Sub
Sub Chart2()
Attribute Chart2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Chart2 Makro
'

'
    Sheets("Amaç F. ve Kýsýtlar").Select
    ActiveWindow.SmallScroll Down:=0
    Range("L26:O30").Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("'Amaç F. ve Kýsýtlar'!$L$26:$O$30")
    ActiveChart.SetElement (msoElementDataTableWithLegendKeys)
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendNone)
    ActiveChart.SetElement (msoElementLegendRight)
    ActiveChart.SetElement (msoElementLegendNone)
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveChart.ClearToMatchStyle
    ActiveChart.ChartStyle = 203
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.ChartTitle.Select
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    Selection.Delete
    ActiveSheet.ChartObjects("Grafik 5").Activate
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Müþteriler"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Müþteriler"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 10).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 10).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 9
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    Selection.Left = 200.075
    Selection.Top = 196.014
    ActiveChart.ChartArea.Select
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = _
        "Daðýtým merkezlerinden müþterilere gönderilen miktarlar"
    Selection.Format.TextFrame2.TextRange.Characters.Text = _
        "Daðýtým merkezlerinden müþterilere gönderilen miktarlar"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 55).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 7).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .Caps = msoAllCaps
        .NameComplexScript = "Times New Roman"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(127, 127, 127)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10.5
        .Italic = msoFalse
        .Kerning = 12
        .Name = "Times New Roman"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 1.5
        .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(8, 48).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .Caps = msoAllCaps
        .NameComplexScript = "Times New Roman"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(127, 127, 127)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10.5
        .Italic = msoFalse
        .Kerning = 12
        .Name = "Times New Roman"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 1.5
        .Strike = msoNoStrike
    End With
    Selection.Left = 49.427
    Selection.Top = -9
    Selection.Left = 49.427
    Selection.Top = 3
    Selection.Left = 49.427
    Selection.Top = 6
    ActiveChart.PlotArea.Select
    ActiveChart.ChartArea.Select
    ActiveChart.Parent.Cut
    Sheets("Karar Destek Sistemi").Select
    ActiveSheet.Paste
    ActiveSheet.ChartObjects("Grafik 19").Activate
    ActiveSheet.Shapes("Grafik 19").IncrementLeft -261
    ActiveSheet.Shapes("Grafik 19").IncrementTop -224.25
    Sheets("Amaç F. ve Kýsýtlar").Select
    ActiveWindow.SmallScroll Down:=-11
    Range("L17:P21").Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("'Amaç F. ve Kýsýtlar'!$L$17:$P$21")
    ActiveChart.SetElement (msoElementDataTableWithLegendKeys)
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisShow)
    ActiveChart.SetElement (msoElementPrimaryValueAxisShow)
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    Selection.Delete
    ActiveSheet.ChartObjects("Grafik 7").Activate
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Grafik 7").IncrementLeft -42
    ActiveSheet.Shapes("Grafik 7").IncrementTop 3
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Daðýtým Merkezleri"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Daðýtým Merkezleri"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 18).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 18).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    ActiveChart.ChartArea.Select
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    Selection.Left = 156.95
    Selection.Top = 169.919
    ActiveChart.ChartArea.Select
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisShow)
    ActiveChart.SetElement (msoElementPrimaryValueAxisShow)
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendNone)
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    Selection.Left = 151.95
    Selection.Top = 193.919
    ActiveChart.ChartArea.Select
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartArea.Select
    ActiveChart.ClearToMatchStyle
    ActiveChart.ChartStyle = 203
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = _
        "Fabrikalardan Daðýtým merkezlerine gönderilen miktar"
    Selection.Format.TextFrame2.TextRange.Characters.Text = _
        "Fabrikalardan Daðýtým merkezlerine gönderilen miktar"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 52).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 13).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .Caps = msoAllCaps
        .NameComplexScript = "Times New Roman"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(127, 127, 127)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10.5
        .Italic = msoFalse
        .Kerning = 12
        .Name = "Times New Roman"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 1.5
        .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(14, 39).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .Caps = msoAllCaps
        .NameComplexScript = "Times New Roman"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(127, 127, 127)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10.5
        .Italic = msoFalse
        .Kerning = 12
        .Name = "Times New Roman"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 1.5
        .Strike = msoNoStrike
    End With
    Selection.Left = 63.35
    Selection.Top = 0
    Selection.Left = 62.35
    Selection.Top = 9
    ActiveChart.ChartArea.Select
    ActiveChart.Parent.Cut
    Sheets("Karar Destek Sistemi").Select
    Range("K17").Select
    Sheets("Karar Destek Sistemi").Select
    ActiveSheet.Paste
    ActiveSheet.ChartObjects("Grafik 21").Activate
    ActiveSheet.Shapes("Grafik 21").IncrementLeft -117
    ActiveSheet.Shapes("Grafik 21").IncrementTop -22.5
    Sheets("Amaç F. ve Kýsýtlar").Select
    ActiveWindow.SmallScroll Down:=11
    Sheets("Karar Destek Sistemi").Select
    ActiveWindow.SmallScroll Down:=0
    Sheets("Amaç F. ve Kýsýtlar").Select
    ActiveWindow.SmallScroll Down:=-11
    Range("L4:P12").Select
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("'Amaç F. ve Kýsýtlar'!$L$4:$P$12")
    Range("T7").Select
    ActiveSheet.ChartObjects("Grafik 8").Activate
    ActiveChart.SetElement (msoElementDataTableWithLegendKeys)
    ActiveChart.SetElement (msoElementDataLabelOutSideEnd)
    ActiveChart.SetElement (msoElementPrimaryValueGridLinesNone)
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.SetElement (msoElementLegendNone)
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
    ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
    Selection.Delete
    ActiveSheet.ChartObjects("Grafik 8").Activate
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Fabrikalar"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "Fabrikalar"
    With Selection.Format.TextFrame2.TextRange.Characters(1, 10).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 10).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 10
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Strike = msoNoStrike
    End With
    Selection.Left = 175.47
    Selection.Top = 194.794
    ActiveChart.ChartArea.Select
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "TEDARÝKÇÝDEN "
    Selection.Format.TextFrame2.TextRange.Characters.Text = "TEDARÝKÇÝDEN "
    With Selection.Format.TextFrame2.TextRange.Characters(1, 13).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 12).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(13, 1).Font
        .BaselineOffset = 0
        .Bold = msoFalse
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(89, 89, 89)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 14
        .Italic = msoFalse
        .Kerning = 12
        .Name = "+mn-lt"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 0
        .Strike = msoNoStrike
    End With
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Grafik 8").ScaleHeight 1.4114581511, msoFalse, _
        msoScaleFromBottomRight
    ActiveChart.ClearToMatchStyle
    ActiveChart.ChartStyle = 203
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = _
        "TEDARÝKÇilerden fabrikalara gönderilen kompenent miktarlarý "
    Selection.Format.TextFrame2.TextRange.Characters.Text = _
        "TEDARÝKÇilerden fabrikalara gönderilen kompenent miktarlarý "
    With Selection.Format.TextFrame2.TextRange.Characters(1, 60).ParagraphFormat
        .TextDirection = msoTextDirectionLeftToRight
        .Alignment = msoAlignCenter
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(1, 15).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .Caps = msoAllCaps
        .NameComplexScript = "Times New Roman"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(127, 127, 127)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 11
        .Italic = msoFalse
        .Kerning = 12
        .Name = "Times New Roman"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 1.5
        .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(16, 44).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .Caps = msoAllCaps
        .NameComplexScript = "Times New Roman"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(127, 127, 127)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 11
        .Italic = msoFalse
        .Kerning = 12
        .Name = "Times New Roman"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 1.5
        .Strike = msoNoStrike
    End With
    With Selection.Format.TextFrame2.TextRange.Characters(60, 1).Font
        .BaselineOffset = 0
        .Bold = msoTrue
        .Caps = msoAllCaps
        .NameComplexScript = "Times New Roman"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(127, 127, 127)
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 11
        .Italic = msoFalse
        .Kerning = 12
        .Name = "Times New Roman"
        .UnderlineStyle = msoNoUnderline
        .Spacing = 1.5
        .Strike = msoNoStrike
    End With
    Selection.Left = 42.387
    Selection.Top = -6
    Selection.Left = 42.387
    Selection.Top = 5
    ActiveChart.SetElement (msoElementPrimaryCategoryAxisNone)
    ActiveChart.SetElement (msoElementPrimaryValueAxisNone)
    ActiveChart.ChartArea.Select
    ActiveSheet.Shapes("Grafik 8").ScaleHeight 1.1143912878, msoFalse, _
        msoScaleFromTopLeft
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    Selection.Left = 175.47
    Selection.Top = 320.396
    ActiveChart.ChartArea.Select
    ActiveChart.Parent.Cut
    Sheets("Karar Destek Sistemi").Select
    Range("D21").Select
    ActiveSheet.Paste
    ActiveSheet.ChartObjects("Grafik 22").Activate
    ActiveSheet.Shapes("Grafik 22").IncrementLeft -144
    ActiveSheet.Shapes("Grafik 22").IncrementTop -56.25
End Sub

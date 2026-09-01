Attribute VB_Name = "PrintCharts"
'Public Enum EExcelVersion
'    Excel2003 = 11
'    Excel2007 = 12
'    Excel2010 = 14
'End Enum

Sub ExportResults()
Dim NewWkbk As Workbook
Dim NewWS As Worksheet
Dim ScatterPlot As Integer
Dim LastOutputRow As Integer
Dim Col1, Col2, Col3, Col4 As Integer
Dim FirstUpRow As Integer

ExcelVersion = Application.Version




Application.ScreenUpdating = False
BaselineText = Range("BaselineTranslation")

'-----------Export the results to a new workbook------------------
Set NewWkbk = Workbooks.Add
    
'Colors
NewWkbk.Colors(10) = RGB(175, 123, 179) 'Baseline (purple)
NewWkbk.Colors(11) = RGB(149, 105, 179) 'Baseline upper bar
NewWkbk.Colors(12) = RGB(189, 189, 189) 'Fail gray - scatter and lower bar
NewWkbk.Colors(13) = RGB(135, 135, 135) 'fail gray - upper bar
NewWkbk.Colors(14) = RGB(67, 162, 202) 'pass scatter and lower bar
NewWkbk.Colors(15) = RGB(50, 127, 202) 'pass blue - upper bar
'NewWkbk.Colors(16) = RGB(255, 255, 255) 'White
    
    ThisWorkbook.Sheets("MResults").Activate
    Cells.Select
    Selection.Copy
            
    NewWkbk.Sheets(1).Activate
    ActiveSheet.Name = "Results"
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    ThisWorkbook.Sheets("MResults").Activate
    Rows("1:4").Copy
        
    NewWkbk.Sheets("Results").Activate
    Rows("1:4").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    
    NewWkbk.Sheets(1).Rows("3:3").Delete
    NewWkbk.Sheets(1).Range("A1").Select


    LastOutputRow = NewWkbk.Sheets(1).Cells.find(What:="*", after:=[a1], SearchDirection:=xlPrevious).row


'-----------add a tab for illuminance and luminance versus wattage scatter plots---------

For NewTab = 1 To 2
    SeriesNamePass = "Passes"
    SeriesNameFail = "Fails"
    If NewTab = 1 Then
        'Define the Illuminance ranges and labels
        GraphShtName = "Illuminance"
        GraphShtRange1 = "B:G"
        GraphShtRange2 = "H:P"
        ChartTitle = "Illuminance vs. Wattage"
        ChartTitleBar = "Illuminance Range"
        Xtitle = "Wattage"
        XtitleBar = "Illuminance"
        Ytitle = "Illuminance"
        YtitleBar = "Fixtures"
    
    ElseIf NewTab = 2 Then
        'Define the Luminance ranges and labels
        GraphShtName = "Luminance"
        GraphShtRange1 = "B:G"
        GraphShtRange2 = "Q:Y"
        ChartTitle = "Luminance vs. Wattage"
        ChartTitleBar = "Luminance Range"
        Xtitle = "Wattage"
        XtitleBar = "Luminance"
        Ytitle = "Luminance"
        YtitleBar = "Fixtures"
    End If
    
    'Copy results to second tab
    NewWkbk.Sheets.Add.Name = GraphShtName
    NewWkbk.Sheets("Results").Range(GraphShtRange1).Copy Destination:=NewWkbk.Sheets(GraphShtName).Range("A:f")
    NewWkbk.Sheets("Results").Range(GraphShtRange2).Copy Destination:=NewWkbk.Sheets(GraphShtName).Range("g:M")
    
    'Sort
    FirstUpRow = 4
    DataRange = "A" & FirstUpRow & ":O" & LastOutputRow
    NewWkbk.Sheets(GraphShtName).Range(DataRange).Sort Key1:=Range("O:O"), _
        order1:=xlDescending, Key2:=Range("D:D"), order2:=xlAscending
    
    '--------------Identify ranges for graphs
    'Check to see if no upgrade series were created
    PassFailCol = 15
    If NewWkbk.Sheets(GraphShtName).Cells(4, PassFailCol) = "" Then
        PassSeries = "Skip"
        FailSeries = "Skip"
    'If the first row fails, it means all rows failed
    ElseIf NewWkbk.Sheets(GraphShtName).Cells(FirstUpRow, PassFailCol) = 0 Then
        PassSeries = "Skip"
        FailSeries = "Make"
        PassRowFirst = 4
        PassRowLast = 4
        FailRowFirst = FirstUpRow
        FailRowLast = LastOutputRow
    'If the last row passes, it means all fixtures passed
    ElseIf NewWkbk.Sheets(GraphShtName).Cells(LastOutputRow, PassFailCol) = 1 Then
        PassSeries = "Make"
        FailSeries = "Skip"
        PassRowFirst = FirstUpRow
        PassRowLast = LastOutputRow
        FailRowFirst = 4
        FailRowLast = 4
    Else
    'Find the split between the two sections
        For r = FirstUpRow To LastOutputRow
            If NewWkbk.Sheets(GraphShtName).Cells(r, PassFailCol) = 1 Then 'do nothing
            ElseIf NewWkbk.Sheets(GraphShtName).Cells(r, PassFailCol) = 0 Then
                FirstFailRow = r
                r = LastOutputRow
            End If
        Next r
        
        PassSeries = "Make"
        FailSeries = "Make"
        PassRowFirst = FirstUpRow
        PassRowLast = FirstFailRow - 1
        FailRowFirst = FirstFailRow
        FailRowLast = LastOutputRow
    End If
        
    
    'Write values to use for the stacked bar chart
    'Headers
        'Average minus minimum
        NewWkbk.Sheets(GraphShtName).Cells(2, PassFailCol + 1) = NewWkbk.Sheets(GraphShtName).Cells(2, 7) & " - " & NewWkbk.Sheets(GraphShtName).Cells(2, 8)
        'Max minus Average
        NewWkbk.Sheets(GraphShtName).Cells(2, PassFailCol + 2) = NewWkbk.Sheets(GraphShtName).Cells(2, 9) & " - " & NewWkbk.Sheets(GraphShtName).Cells(2, 7)
        NewWkbk.Sheets(GraphShtName).Cells(2, PassFailCol + 1).WrapText = True
        NewWkbk.Sheets(GraphShtName).Cells(2, PassFailCol + 2).WrapText = True
        
    ReDim MinAvg(3 To LastOutputRow)
    ReDim AvgMax(3 To LastOutputRow)
    For row = 3 To LastOutputRow
        Average = NewWkbk.Sheets(GraphShtName).Cells(row, 7)
        minimum = NewWkbk.Sheets(GraphShtName).Cells(row, 8)
        Maximum = NewWkbk.Sheets(GraphShtName).Cells(row, 9)
        
        'check for calculation errors returned as '#value or #N/A
        If IsNumeric(Average) = False Or IsNumeric(minimum) = False Or IsNumeric(Maximum) = False Then
            NewWkbk.Sheets(GraphShtName).Cells(row, PassFailCol + 1) = "Error"
            NewWkbk.Sheets(GraphShtName).Cells(row, PassFailCol + 2) = "Error"
        Else
            'min to average range
            MinAvg(row) = Average - minimum
            AvgMax(row) = Maximum - Average
            
            NewWkbk.Sheets(GraphShtName).Cells(row, PassFailCol + 1) = MinAvg(row)
            NewWkbk.Sheets(GraphShtName).Cells(row, PassFailCol + 2) = AvgMax(row)
        End If
    Next row


    
 '-------------------------------------------------------------------------------------------
 ' add the chart first chart (scatter plot)
 '***edited for Excel 2003 compatibility
    Set myChtObj = Charts.Add
    Set myChtObj = myChtObj.Location(Where:=xlLocationAsObject, Name:=GraphShtName)
    myChtObj.Parent.Top = Range("F8").Top
    myChtObj.Parent.Left = Range("F8").Left
  'Excel 2007 method - testing 2003 version above
'    Set myChtObj = ActiveSheet.ChartObjects.Add _
'        (Left:=250, Width:=375, Top:=150, Height:=225)
    With myChtObj '.Chart

            '***Delete existing series (2003 VBA method)
            For n = .SeriesCollection.count To 1 Step -1
                .SeriesCollection(n).Delete
            Next n
            '***
            ' make an XY chart
            .ChartType = xlXYScatter

            
            'Baseline series
            With .SeriesCollection.NewSeries
                .Values = NewWkbk.Sheets(GraphShtName).Range("G3:G3")
                .Xvalues = NewWkbk.Sheets(GraphShtName).Range("D3:D3")
                .Name = BaselineText
                    .MarkerBackgroundColorIndex = 10
                    If ExcelVersion = "12.0" Or ExcelVersion = "14.0" Then
                    .format.Fill.ForeColor.RGB = RGB(175, 123, 179) 'Interior
                    .format.Line.Visible = msoFalse
                    End If
                .MarkerForegroundColorIndex = 2
                .MarkerSize = 8
                
            End With
            
            'Series for fixtures that fail Criteria
            If FailSeries = "Skip" Then
                'no fixtures failed - skip
            Else 'add series for fixtures that failed
            With .SeriesCollection.NewSeries
                .Values = NewWkbk.Sheets(GraphShtName).Range("G" & FailRowFirst & ":G" & FailRowLast)
                .Xvalues = NewWkbk.Sheets(GraphShtName).Range("D" & FailRowFirst & ":D" & FailRowLast)
                .Name = SeriesNameFail
                    .MarkerBackgroundColorIndex = 12
                    If ExcelVersion = "12.0" Or ExcelVersion = "14.0" Then
                    .format.Fill.ForeColor.RGB = RGB(189, 189, 189) 'Interior
                    .format.Line.Visible = msoFalse
                    End If
                .MarkerForegroundColorIndex = 2
                .MarkerSize = 7
                
            End With
            End If
            
            'Series for fixtures that pass criteria
            If PassSeries = "Skip" Then ' do nothing
            Else 'add series for those that passed
                With .SeriesCollection.NewSeries
                    .Values = NewWkbk.Sheets(GraphShtName).Range("G" & PassRowFirst & ":G" & PassRowLast)
                    .Xvalues = NewWkbk.Sheets(GraphShtName).Range("D" & PassRowFirst & ":D" & PassRowLast)
                    .Name = SeriesNamePass
                        .MarkerBackgroundColorIndex = 14
                        If ExcelVersion = "12.0" Or ExcelVersion = "14.0" Then
                        .format.Fill.ForeColor.RGB = RGB(67, 162, 202) 'Interior
                        .format.Line.Visible = msoFalse
                        End If
                    .MarkerForegroundColorIndex = 2
                    .MarkerSize = 7
                End With
            End If
            
            

        .HasTitle = True
        .ChartTitle.Text = ChartTitle
        'x-axis name
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = Xtitle
        'y-axis name
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = Ytitle
        .PlotArea.Interior.ColorIndex = xlNone
    End With
    
    '-------------------------------------------------------------------------
    'add the second chart -  stacked bar chart for values
    Set myChtObj = ActiveSheet.ChartObjects.Add _
        (Left:=650, Width:=375, Top:=75, Height:=150 + LastOutputRow * 18)
    With myChtObj.Chart

        ' make an XY chart
        .ChartType = xlBarStacked
        .ChartGroups(1).GapWidth = 50
            For Series = 1 To 3
                If Series = 1 Then
                    ValuesRange = "H3:H"
                    SeriesName = "Minimum"
                    SeriesColor = RGB(255, 255, 255)

                ElseIf Series = 2 Then
                    ValuesRange = "P3:P"
                    SeriesName = "Min to Avg"
                    SeriesColor = RGB(67, 162, 202)
                ElseIf Series = 3 Then
                    ValuesRange = "Q3:Q"
                    SeriesName = "Avg to Max"
                    SeriesColor = RGB(50, 127, 202)

                End If
                
                With .SeriesCollection.NewSeries
                    .Values = NewWkbk.Sheets(GraphShtName).Range(ValuesRange & LastOutputRow)
                    .Xvalues = NewWkbk.Sheets(GraphShtName).Range("B3:B" & LastOutputRow)
                    .Name = SeriesName
                        'Excel 2003 compatibility
                        If Series = 1 Then .Interior.ColorIndex = xlNone
                        
                        If ExcelVersion = "12.0" Or ExcelVersion = "14.0" Then
                        .format.Fill.ForeColor.RGB = SeriesColor
                         If Series = 1 Then .format.Fill.Visible = msoFalse
                        .format.Line.Visible = msoFalse
                        End If
                End With
            Next Series

            'Change the colors of individual points
            'baseline
                
                .SeriesCollection(2).Points(1).Interior.ColorIndex = 10
                .SeriesCollection(3).Points(1).Interior.ColorIndex = 11
                If ExcelVersion = "12.0" Or ExcelVersion = "14.0" Then
                .SeriesCollection(2).Points(1).format.Fill.ForeColor.RGB = RGB(175, 123, 179)
                .SeriesCollection(3).Points(1).format.Fill.ForeColor.RGB = RGB(149, 105, 179)
                End If
            'pass/fail check
            For i = 2 To LastOutputRow - 2 'each row
                If NewWkbk.Sheets(GraphShtName).Cells(i + 2, PassFailCol) = 0 Then
                    FillColor1 = RGB(189, 189, 189)
                    FillColor2 = RGB(135, 135, 135)
                    FillIndex1 = 12
                    FillIndex2 = 13
                ElseIf NewWkbk.Sheets(GraphShtName).Cells(i + 2, PassFailCol) = 1 Then
                    FillColor1 = RGB(67, 162, 202)
                    FillColor2 = RGB(50, 127, 202)
                    FillIndex1 = 14
                    FillIndex2 = 15
                End If
                '2003 formatting
                .SeriesCollection(2).Points(i).Interior.ColorIndex = FillIndex1
                .SeriesCollection(3).Points(i).Interior.ColorIndex = FillIndex2
                '2007 formatting
                If ExcelVersion = "12.0" Or ExcelVersion = "14.0" Then
                 .SeriesCollection(2).Points(i).format.Fill.ForeColor.RGB = FillColor1
                 .SeriesCollection(3).Points(i).format.Fill.ForeColor.RGB = FillColor2
                End If
            Next i
            
            'Change chart settings
            .PlotArea.Interior.ColorIndex = xlNone
            
            'If ExcelVersion = "12.0" Or ExcelVersion = "14.0" Then
            .HasTitle = True
            .ChartTitle.Text = ChartTitleBar
            'x-axis name
            .Axes(xlCategory, xlPrimary).HasTitle = True
            .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = XtitleBar
            'y-axis name
            .Axes(xlValue, xlPrimary).HasTitle = True
            .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = YtitleBar
            .Legend.Position = xlTop
            .Axes(xlCategory).ReversePlotOrder = True
            '.Axes(xlCategory).Crosses = xlMaximum
            'End If
            
   End With
   
'Change color coding of line text based on pass/fail
     For r = FirstUpRow To LastOutputRow
        NewWkbk.Sheets(GraphShtName).Select
         If Cells(r, PassFailCol) = 1 Then
            NewWkbk.Sheets(GraphShtName).Cells(r, PassFailCol).EntireRow.Font.Color = RGB(50, 127, 202) 'blue if pass
         ElseIf NewWkbk.Sheets(GraphShtName).Cells(r, PassFailCol) = 0 Then
            NewWkbk.Sheets(GraphShtName).Cells(r, PassFailCol).EntireRow.Font.ColorIndex = 12
                If ExcelVersion = "12.0" Or ExcelVersion = "14.0" Then
                NewWkbk.Sheets(GraphShtName).Cells(r, PassFailCol).EntireRow.Font.Color = RGB(128, 128, 128) 'gray if fail
                End If
         End If
     Next r

    
Next NewTab

'Change color coding of line text on the Results sheet based on pass/fail
For r = FirstUpRow To LastOutputRow
         If NewWkbk.Sheets("Results").Cells(r, 16) = 0 And NewWkbk.Sheets("Results").Cells(r, 25) = 0 Then 'gray if they both fail
            NewWkbk.Sheets("Results").Cells(r, 13).EntireRow.Font.ColorIndex = 12
            If ExcelVersion = "12.0" Or ExcelVersion = "14.0" Then
            NewWkbk.Sheets("Results").Cells(r, 13).EntireRow.Font.Color = RGB(128, 128, 128) 'gray if fail
            End If
         End If
     Next r
Application.ScreenUpdating = True
NewWkbk.Sheets("Results").Select

End Sub

Sub MakeScatterPlot()

End Sub




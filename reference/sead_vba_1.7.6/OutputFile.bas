Attribute VB_Name = "OutputFile"
'This macro was developed by p2w2.  http://p2w2.com/expert-in-microsoft-excel-consultants-consulting/index.php
'Please contact at CS@perceptive-analytics.com in case of any enquiry

Sub outFile(method As String)
'Application.ScreenUpdating = False
Dim pbProgBar As IProgressBar
Set pbProgBar = New FProgressBarIFace
pbProgBar.Title = "Creating file..."
pbProgBar.Text = "Creating output file..."
pbProgBar.Min = 0
pbProgBar.Max = 1
pbProgBar.Progress = 1
pbProgBar.Show
pbProgBar.Progress = 0.1


Sheets("Input").Visible = True
Sheets("Dashboard").Visible = True
Sheets("Illuminance").Visible = True
Sheets("Luminance").Visible = True
Sheets("Annual Energy").Visible = True
Sheets("Simple Payback").Visible = True
Sheets("Net Present Value").Visible = True
Sheet31.Visible = True

'Dashboard sheet - Avg Vs Wattage charts
Sheets("Illuminance").Range("AA1").Formula = "IllumWattageRangeFails"


Dim laRow As Integer

laRow = Sheets("Input").Range("B" & Rows.count).End(3).row
ThisWorkbook.Sheets("Input").Cells.Borders.LineStyle = xlNone
ThisWorkbook.Sheets("Input").Range("B4:Q" & laRow + 1).ClearContents

laRow = Sheets("Annual Energy").Range("B" & Rows.count).End(3).row
ThisWorkbook.Sheets("Annual Energy").Cells.Borders.LineStyle = xlNone
ThisWorkbook.Sheets("Annual Energy").Range("B4:G" & laRow + 1).ClearContents

laRow = Sheets("Illuminance").Range("B" & Rows.count).End(3).row
ThisWorkbook.Sheets("Illuminance").Cells.Borders.LineStyle = xlNone
ThisWorkbook.Sheets("Illuminance").Range("B4:L" & laRow + 1).ClearContents
ThisWorkbook.Sheets("Illuminance").Range("S4:S" & laRow + 1).ClearContents

laRow = Sheets("Luminance").Range("B" & Rows.count).End(3).row
ThisWorkbook.Sheets("Luminance").Cells.Borders.LineStyle = xlNone
ThisWorkbook.Sheets("Luminance").Range("B4:R" & laRow + 1).ClearContents

laRow = Sheets("Net Present Value").Range("B" & Rows.count).End(3).row
ThisWorkbook.Sheets("Net Present Value").Cells.Borders.LineStyle = xlNone
ThisWorkbook.Sheets("Net Present Value").Range("B4:AB" & laRow + 1).ClearContents

laRow = Sheets("Simple Payback").Range("B" & Rows.count).End(3).row
ThisWorkbook.Sheets("Simple Payback").Cells.Borders.LineStyle = xlNone
ThisWorkbook.Sheets("Simple Payback").Range("B4:AB" & laRow + 1).ClearContents

'ROI sheet
laRow = Sheets("ROI").Range("B" & Rows.count).End(3).row
ThisWorkbook.Sheets("ROI").Cells.Borders.LineStyle = xlNone
ThisWorkbook.Sheets("ROI").Range("B4:AB" & laRow + 1).ClearContents

i = 4

    Dim Rng As Range
 
    'Inputs sheet
    'Header data
    mToInputCol1 = 2        'Column on the source sheet to start the range
    mToInputCount1 = 3      'Number of columns to add to the first column in the source and target ranges
    InputCol1 = 2           'Column on the target sheet to start the range
    
    'Inputs range
    mToInputCol2 = 72
    mToInputCount2 = 9
    InputCol2 = 6
    
    'Multiplier range
    mToInputCol3 = 83
    mToInputCount3 = 0
    InputCol3 = 16
    
    laRow = Sheets("MResults").Range("B" & Rows.count).End(3).row
    'Header data
    With Sheets("MResults")
        Set r1 = .Range(.Cells(i, mToInputCol1), .Cells(laRow, mToInputCol1 + mToInputCount1))
    End With
    With Sheets("Input")
        Set r2 = .Range(.Cells(i, InputCol1), .Cells(laRow, InputCol1 + mToInputCount1))
    End With
    r2.Value = r1.Value
    
    With Sheets("MResults")
        Set r1 = .Range(.Cells(i, mToInputCol2), .Cells(laRow, mToInputCol2 + mToInputCount2))
    End With
    With Sheets("Input")
        Set r2 = .Range(.Cells(i, InputCol2), .Cells(laRow, InputCol2 + mToInputCount2))
    End With
    r2.Value = r1.Value
    
    'Multiplier range (to skip Number of Points in Grid)
    Set r1 = Sheets("MResults").Range(Sheets("MResults").Cells(i, mToInputCol3), Sheets("MResults").Cells(laRow, mToInputCol3 + mToInputCount3))
    Set r2 = Sheets("Input").Range(Sheets("Input").Cells(i, InputCol3), Sheets("Input").Cells(laRow, InputCol3 + mToInputCount3))
    r2.Value = r1.Value
    
    Set Rng = Sheets("Input").Range("B4:P" & laRow)

    With Rng.Borders
        .LineStyle = xlContinuous
        .Color = RGB(191, 191, 191)
        .Weight = xlThin
    End With
 
    Cells(i, 2).Activate

    'Annual Energy sheet
    pbProgBar.Progress = 0.2
    Sheets("MResults").Select
    Range(Cells(i, 2), Cells(laRow, 5)).Copy
    Sheets("Annual Energy").Select
    Range(Cells(i, 2), Cells(laRow, 5)).PasteSpecial xlPasteValues
    
    Sheets("MResults").Select
    Range(Cells(i, 6), Cells(laRow, 7)).Copy
    Sheets("Annual Energy").Select
    Range(Cells(i, 6), Cells(laRow, 7)).PasteSpecial xlPasteValues
    
    Range("B5:G" & laRow).Select
    ActiveWorkbook.Worksheets("Annual Energy").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Annual Energy").Sort.SortFields.Add Key:=Range( _
        "F5:F" & laRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Annual Energy").Sort
        .SetRange Range("B5:G" & laRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Set Rng = Range("B4:G" & laRow)
    With Rng.Borders
        .LineStyle = xlContinuous
        .Color = RGB(191, 191, 191)
        .Weight = xlThin
    End With

    Cells(i, 2).Activate
 
 i = 4
 ' Illuminance sheet
    pbProgBar.Progress = 0.3
    Sheets("MResults").Select
    Range(Cells(i, 2), Cells(laRow, 5)).Copy
    Sheets("Illuminance").Select
    Range(Cells(i, 2), Cells(laRow, 5)).PasteSpecial xlPasteValues
    
    Sheets("MResults").Select
    Range(Cells(i, 8), Cells(laRow, 13)).Copy
    Sheets("Illuminance").Select
    Range(Cells(i, 6), Cells(laRow, 11)).PasteSpecial xlPasteValues
    
    Sheets("MResults").Select
    Range(Cells(i, 16), Cells(laRow, 16)).Copy
    Sheets("Illuminance").Select
    Range(Cells(i, 12), Cells(laRow, 12)).PasteSpecial xlPasteValues
    
    Sheets("MResults").Select
    Range(Cells(i, 6), Cells(laRow, 6)).Copy
    Sheets("Illuminance").Select
    Range(Cells(i, 19), Cells(laRow, 19)).PasteSpecial xlPasteValues
    
    Range("B5:S" & laRow).Select
    ActiveWorkbook.Worksheets("Illuminance").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Illuminance").Sort.SortFields.Add Key:=Range( _
        "N5:N" & laRow), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Illuminance").Sort.SortFields.Add Key:=Range( _
        "S5:S" & laRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Illuminance").Sort
        .SetRange Range("B5:S" & laRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
 
    Set Rng = Range("B4:L" & laRow)

    With Rng.Borders
        .LineStyle = xlContinuous
        .Color = RGB(191, 191, 191)
        .Weight = xlThin
    End With

 Cells(i, 2).Activate
 
 'Luminance sheet
 pbProgBar.Progress = 0.4
    Sheets("MResults").Select
    Range(Cells(i, 2), Cells(laRow, 5)).Copy
    Sheets("Luminance").Select
    Range(Cells(i, 2), Cells(laRow, 5)).PasteSpecial xlPasteValues
    
    Sheets("MResults").Select
    Range(Cells(i, 17), Cells(laRow, 25)).Copy
    Sheets("Luminance").Select
    Range(Cells(i, 6), Cells(laRow, 14)).PasteSpecial xlPasteValues
    
    Sheets("MResults").Select
    Range(Cells(i, 6), Cells(laRow, 6)).Copy
    Sheets("Luminance").Select
    Range(Cells(i, 18), Cells(laRow, 18)).PasteSpecial xlPasteValues
 
 
    
    Range("B5:R" & laRow).Select
    ActiveWorkbook.Worksheets("Luminance").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Luminance").Sort.SortFields.Add Key:=Range( _
        "N5:N" & laRow), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("Luminance").Sort.SortFields.Add Key:=Range( _
        "R5:R" & laRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Luminance").Sort
        .SetRange Range("B5:R" & laRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    Set Rng = Range("B4:N" & laRow)

    With Rng.Borders
        .LineStyle = xlContinuous
        .Color = RGB(191, 191, 191)
        .Weight = xlThin
    End With
    
    If Sheets("FixtureData").Range("iescieGraphChoice").Value = "IES" Then
        Sheets("Luminance").Range("I3").Value = Sheets("Translation").Range("tMsg1").Value
        Sheets("Luminance").Range("J3").Value = Sheets("Translation").Range("tMsg2").Value
        Sheets("Luminance").Range("L3").Value = Sheets("Translation").Range("tMsg5").Value
        Sheets("Luminance").Range("M3").Value = Sheets("Translation").Range("tMsg6").Value
        
    ElseIf Sheets("FixtureData").Range("iescieGraphChoice").Value = "CIE" Then
        Sheets("Luminance").Range("I3").Value = Sheets("Translation").Range("tMsg3").Value
        Sheets("Luminance").Range("J3").Value = Sheets("Translation").Range("tMsg4").Value
        Sheets("Luminance").Range("L3").Value = Sheets("Translation").Range("tMsg8").Value
        Sheets("Luminance").Range("M3").Value = Sheets("Translation").Range("tMsg9").Value
    End If
Cells(i, 2).Activate
 
 'Simple Payback sheet
    pbProgBar.Progress = 0.5
    Sheets("MResults").Select
    Range(Cells(i, 2), Cells(laRow, 5)).Copy
    Sheets("Simple Payback").Select
    Range(Cells(i, 2), Cells(laRow, 5)).PasteSpecial xlPasteValues
    
    Sheets("MResults").Select
    Range(Cells(i, 26), Cells(laRow, 30)).Copy
    Sheets("Simple Payback").Select
    Range(Cells(i, 6), Cells(laRow, 7)).PasteSpecial xlPasteValues
    
    
    
    Set Rng = Range("B4:J" & laRow)

    With Rng.Borders
        .LineStyle = xlContinuous
        .Color = RGB(191, 191, 191)
        .Weight = xlThin
    End With
  

 Cells(i, 2).Activate
 
 'Net Present Value sheet
    pbProgBar.Progress = 0.6
    Sheets("MResults").Select
    Range(Cells(i, 2), Cells(laRow, 5)).Copy
    Sheets("Net Present Value").Select
    Range(Cells(i, 2), Cells(laRow, 5)).PasteSpecial xlPasteValues
    
    Sheets("MResults").Select
    Range(Cells(i, 31), Cells(laRow, 51)).Copy
    Sheets("Net Present Value").Select
    Range(Cells(i, 6), Cells(laRow, 26)).PasteSpecial xlPasteValues
    
    Sheets("MResults").Select
    Range(Cells(i, 6), Cells(laRow, 6)).Copy
    Sheets("Net Present Value").Select
    Range(Cells(i, 28), Cells(laRow, 28)).PasteSpecial xlPasteValues
    
    Range("B5:AB" & laRow).Select
    ActiveWorkbook.Worksheets("Net Present Value").Sort.SortFields.Clear

    ActiveWorkbook.Worksheets("Net Present Value").Sort.SortFields.Add Key:=Range( _
        "AB5:AB" & laRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Net Present Value").Sort
        .SetRange Range("B5:AB" & laRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
       Set Rng = Range("B4:Z" & laRow)

    With Rng.Borders
        .LineStyle = xlContinuous
        .Color = RGB(191, 191, 191)
        .Weight = xlThin
    End With

Cells(i, 2).Activate
'-----------------------------------------
'Internal Rate of Return / Return on Investment
    pbProgBar.Progress = 0.8
    Sheets("MResults").Select
    Range(Cells(i, 2), Cells(laRow, 5)).Copy
    Sheet31.Select
    Range(Cells(i, 2), Cells(laRow, 5)).PasteSpecial xlPasteValues
    
    Sheets("MResults").Select
    Range(Cells(i, 52), Cells(laRow, 71)).Copy
    Sheet31.Select
    Range(Cells(i, 6), Cells(laRow, 25)).PasteSpecial xlPasteValues
    
    Sheets("MResults").Select
    Range(Cells(i, 6), Cells(laRow, 6)).Copy
    Sheet31.Select
    Range(Cells(i, 27), Cells(laRow, 27)).PasteSpecial xlPasteValues
    
    'Sorting
    Range("B5:AA" & laRow).Select
    ActiveWorkbook.Worksheets("ROI").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("ROI").Sort.SortFields.Add Key:=Range( _
        "AA5:AA" & laRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("ROI").Sort
        .SetRange Range("B5:AA" & laRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Formatting
    Set Rng = Range("B4:Y" & laRow)
    With Rng.Borders
        .LineStyle = xlContinuous
        .Color = RGB(191, 191, 191)
        .Weight = xlThin
    End With

Cells(i, 2).Activate
 '-------------------------------------------
'create output file
Dim wb As Workbook
'Set wb = Workbooks.Add
pbProgBar.Progress = 0.9
'Sheets(Array("Dashboard", "Input", "Annual Energy", "Illuminance", "Luminance", "Simple Payback", "Net Present Value")).Select
'Sheets("Dashboard").Activate
ThisWorkbook.Sheets(Array("Dashboard", "Input", "Annual Energy", "Illuminance", "Luminance", "Simple Payback", "Net Present Value", "Translation", "ROI")).Copy
Set wb = ActiveWorkbook
wb.Sheets("Translation").Visible = xlSheetHidden
wb.Sheets("Dashboard").Activate
wb.SaveAs ThisWorkbook.Path & "\" & method & "Results" & format(Now(), "mm_dd_yy HH_MM_SS") & ".xlsx "
wb.Saved = True

'wb.Close
pbProgBar.Progress = 1
With ThisWorkbook

.Sheets("Input").Visible = False
.Sheets("Dashboard").Visible = False
.Sheets("Illuminance").Visible = False
.Sheets("Luminance").Visible = False
.Sheets("Annual Energy").Visible = False
.Sheets("Simple Payback").Visible = False
.Sheets("Net Present Value").Visible = False
.Sheets("ROI").Visible = False

End With

pbProgBar.Hide

Dim outMsg As String
outMsg = Sheets("Translation").Range("OutputGenMsg").Value

MsgBox outMsg & ThisWorkbook.Path
  
 
'Application.ScreenUpdating = True
 
End Sub



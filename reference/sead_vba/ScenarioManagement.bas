Attribute VB_Name = "ScenarioManagement"
Sub SaveSingleScenario()

SavePrompt = Sheet25.Range("SavePrompt")
SaveTitle = Sheet25.Range("SaveTitle")
SaveDefault = Sheet25.Range("SaveDefault")

ScenarioName = Application.InputBox(prompt:=SavePrompt, _
          Title:=SaveTitle, Default:=SaveDefault)

If ScenarioName = False Then
    Exit Sub
Else


    lastRow = Sheet10.Cells.find(What:="*", after:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    lastCol = Sheet10.Cells.find(What:="*", after:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).column
    
    'change to upgrade and save only upgrade data
    Sheet3.Range("Base_Upgrade_Choice").Value = "Upgrade"
    RefreshIllCalcs
    
    Sheet10.Cells(lastRow + 1, 2) = ScenarioName
    
            For c = 3 To lastCol
                Sheet10.Cells(lastRow + 1, c) = Sheet10.Cells(3, c)
            Next c
    
    SortMultiResults (7)
    UpdateGraphAlignmentData

End If
End Sub

Sub UpdateGraphAlignmentData()
Application.Calculation = xlCalculationManual

lastRow = Sheet10.Cells.find(What:="*", after:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
count = 0
For j = 4 To lastRow
    Sheet10.Cells(j, 1) = 0.5 + count
    count = count + 1
Next j
Application.Calculation = xlCalculationAutomatic
End Sub

Sub SortMultiResults(SortCol As Integer)
Dim screenUpdateState As Variant

screenUpdateState = Application.ScreenUpdating
Application.ScreenUpdating = False


lastRow = Sheet10.Cells.find(What:="*", after:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
lastCol = Sheet10.Cells.find(What:="*", after:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).column

'sort row 5 to last row, all used columns (plus a few extras)
Sheet10.Range("A5:CC" & lastRow).Sort _
    Key1:=Sheet10.Cells(5, SortCol), _
    Header:=xlNo

'-------------code below works for Excel2007 but not 2003---------------
'Sheet10.Sort.SortFields.Clear
'Sheet10.Sort.SortFields.Add Key:=Range(Cells(4, SortCol), Cells(LastRow, SortCol)), _
'        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'    'Sort the results - starting with row 5 ('Baseline' result always stays in row 4)
'    With Sheet10.Sort
'        .SetRange Range(Cells(5, 1), Cells(LastRow, LastCol))
'        .Header = xlNo
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
Application.ScreenUpdating = screenUpdateState

End Sub

Sub DeleteResults()

DeletePrompt = Sheet25.Range("DeletePrompt")
DeleteTitle = Sheet25.Range("DeleteTitle")

'Select Case MsgBox(DeletePrompt, vbOKCancel, DeleteTitle)
'Case vbCancel
'    Exit Sub
'Case vbOK

    lastRow = Sheet10.Cells.find(What:="*", after:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    lastCol = Sheet10.Cells.find(What:="*", after:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).column
    
    If lastRow = 3 Then
        Exit Sub
    ElseIf lastRow > 3 Then
        Sheet10.Activate
        Sheet10.Range(Cells(4, 1), Cells(lastRow, lastCol)).ClearContents
    End If
'End Select

End Sub


Sub SortbyEnergy()
    SortMultiResults (7)
End Sub

Sub SortbyNPV()
    SortMultiResults (13)
End Sub
Sub SortbyAvgIll()
    SortMultiResults (16)
End Sub
Sub SortbyAvgLum()
    SortMultiResults (23)
End Sub


Sub UnlockSheet()

    ActiveSheet.Unprotect
    
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        False



End Sub


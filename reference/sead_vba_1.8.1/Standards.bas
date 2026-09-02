Attribute VB_Name = "Standards"
Option Explicit

Sub AddStandard()

Dim row As Integer
Dim firstRow As Integer
Dim lastRow As Integer

Dim Header As Boolean
Dim Section As String
Dim AddStdMessage As String

AddStdMessage = Sheet25.Range("AddStdMessage")

'Find the Header Section,
Section = "Header"
Call IDStartEndRow(Section, firstRow, lastRow)

'Reset FirstRow to be first row to hide
firstRow = lastRow + 1
lastRow = Sheet7.Cells.find(What:="*", After:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
Rows(firstRow & ":" & lastRow).EntireRow.Hidden = True

MsgBox (AddStdMessage)



End Sub

Sub SaveStandard()

Dim StandardName As String
Dim firstRow As Integer
Dim lastRow As Integer
Dim StandardListRow As Integer
Dim nUsedRows As Variant
Dim Section As String
Dim row As Integer

Dim StdNamePrompt As String
Dim StdNameTitle As String
Dim StdNameDefault As String, StdConfirm As String


StdNamePrompt = Sheet25.Range("StdNamePrompt")
StdNameTitle = Sheet25.Range("StdNametitle")
StdNameDefault = Sheet25.Range("StdNameDefault")
StdConfirm = Sheet25.Range("StdConfirm")

StandardName = Application.InputBox(prompt:=StdNamePrompt, _
          Title:=StdNameTitle, Default:=StdNameDefault)

'Find the Header Section,
Section = "Header"
Call IDStartEndRow(Section, firstRow, lastRow)
'unhide all rows
firstRow = lastRow + 1
lastRow = Sheet7.Cells.find(What:="*", After:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
Rows(firstRow & ":" & lastRow).EntireRow.Hidden = False

'Find first and last row of the new section
lastRow = Sheet7.Cells.find(What:="*", After:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
firstRow = Range("A65536").End(xlUp).row + 1


For row = firstRow To lastRow
    Sheet7.Cells(row, 1) = StandardName
Next row

StandardListRow = Sheet7.Range("AB65536").End(xlUp).row + 1
Sheet7.Range("AB" & StandardListRow) = StandardName

Range("AB4") = Range("Standards_List").Rows.count

'hide all non-selected standards
SelectStandard

MsgBox (StdConfirm)


End Sub

Sub SelectStandard()

Dim StandardName As String, Section As String

Dim firstRow As Integer
Dim lastRow As Integer

'unprotect sheet------------------------
'ActiveSheet.Unprotect

    StandardName = Sheet7.Range("AB5")
    
    'Find the Header Section,
    Section = "Header"
    Call IDStartEndRow(Section, firstRow, lastRow)
    
    'Hide all Standard info rows
    firstRow = lastRow + 1
    lastRow = Sheet7.Cells.find(What:="*", After:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    Rows(firstRow & ":" & lastRow).EntireRow.Hidden = True
    
    'Unhide the rows w/ the standard selected
    Call IDStartEndRow(StandardName, firstRow, lastRow)
    Rows(firstRow & ":" & lastRow).EntireRow.Hidden = False

'ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:=False


End Sub


Sub IDStartEndRow(ByRef Section As String, ByRef StartRow As Integer, ByRef EndRow As Integer)

Dim lastRow As Integer
Dim row As Integer
Dim AboveFound As Boolean
Dim Above, Below As Boolean

'--------------------Important - must unhide all rows first!!!

lastRow = Sheet7.Cells.find(What:="*", After:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row

    'initialize variables to find the start and end of the section
    row = 1
    Above = True
    AboveFound = False
    Below = False
    
    Do While Below = False
        If Cells(row, 1) = Section And AboveFound = False Then 'add a with???? insert dot before cell
            StartRow = row
            AboveFound = True
            row = row + 1
        ElseIf Cells(row, 1) <> Section And AboveFound = True Then
            EndRow = row - 1
            Below = True
        Else
            row = row + 1
        End If
    Loop


End Sub

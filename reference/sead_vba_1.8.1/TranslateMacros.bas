Attribute VB_Name = "TranslateMacros"

Option Explicit
Public count As Integer
 
Function ListNamedRanges(rngcell)
Dim xlnName As Excel.Name
Dim rngNamedRange As Range
'Dim rngcell As Range
Dim strOutput As String
'Set rngcell = Selection(1, 1)

For Each xlnName In ActiveWorkbook.Names
    On Error Resume Next
    Set rngNamedRange = xlnName.RefersToRange
    If rngNamedRange.Parent.Name = ActiveSheet.Name Then
        If Union(rngNamedRange, rngcell).Address = rngNamedRange.Address Then
            strOutput = strOutput & vbCrLf & xlnName.Name
        End If
    End If
Next
ListNamedRanges = strOutput
End Function



Sub find()

Dim rngRangeToTrace As Range
Dim c As Range, i As Range
Application.ScreenUpdating = False

Set rngRangeToTrace = Sheet25.Range("B2:B833")
count = 2

For Each c In rngRangeToTrace
    wksTest.Cells(count, 1) = c.Address(, , , True)
    Call FindExternalDependents(c)
    count = count + 1
Next c

Application.ScreenUpdating = True

End Sub

Public Sub FindExternalDependents(ByVal rngPrecedent As Range)
    ' Find all of the cells that are dependent on the precedent cell,
   ' but are on a different sheet to the precedent cell..
   Dim dependentCell As Range

   On Error Resume Next          ' We'll be checking for errors as we go
   rngPrecedent.ShowDependents   ' Show the dependency arrows
    
    Dim arrowNumber As Integer
    arrowNumber = 1
    
    wksTest.Cells(count, 1 + arrowNumber) = ListNamedRanges(rngPrecedent)
    
    Do                           ' We'll break out of this loop when no more dependencies are found.
       rngPrecedent.NavigateArrow False, arrowNumber, 1
        If Err.Number <> 0 Or rngPrecedent.Address = Selection.Address Then
            GoTo NoMoreArrows
        Else
            wksTest.Cells(count, 2 + arrowNumber) = Selection.Address(External:=True)
            Debug.Print "Dependent found at " + Selection.Address(External:=True)
            CheckForExternalLinks rngPrecedent, arrowNumber
            
            ' Check the next arrow.
           arrowNumber = arrowNumber + 1
        End If
    Loop While True
    
NoMoreArrows:
    Exit Sub
End Sub

Private Sub CheckForExternalLinks(ByVal rngPrecedent As Range, ByVal arrowNumber As Integer)
    ' One arrow is the external links arrow.  Follow all of its links.
   Dim linkNumber As Integer
    linkNumber = 2
    
    Do
        rngPrecedent.NavigateArrow False, arrowNumber, linkNumber
        
        If Err.Number <> 0 Then
            GoTo NoMoreArrows
        Else
            wksTest.Cells(count, 2 + arrowNumber + linkNumber) = Selection.Address(External:=True)
            Debug.Print "Dependent found at " + Selection.Address(External:=True)
            
            ' Check the next link
           linkNumber = linkNumber + 1
        End If
    Loop While True

NoMoreArrows:
    Exit Sub
End Sub







Sub FindText()
Dim sSource As String
Dim sTranslate As String
Dim TranslateArray() As Variant

sSource = "InfoBoxes"
sTranslate = "Translation"
r1 = 1
r2 = 471

Application.Calculation = xlCalculationManual

'Read in data from the Translate sheet
c1 = 3

ReDim TranslateArray(r2 - r1)

For i = 0 To (r2 - r1)
    TranslateArray(i) = Sheets(sTranslate).Cells(i + r1, c1)
Next i


'Read data from the Source sheet
LastRowSource = Sheets(sSource).Cells.find(What:="*", After:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
LastColSource = Sheets(sSource).Cells.find(What:="*", After:=[a1], SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).column

For row = 1 To LastRowSource
    For column = 1 To LastColSource
            
    If Sheets(sSource).Cells(row, column).HasFormula Then
    ElseIf Sheets(sSource).Cells(row, column).Value <> Empty Then
    
        Value = Sheets(sSource).Cells(row, column)
        
        For i = 0 To (r2 - r1)
            If TranslateArray(i) = Value Then
                trow = r1 + i
                Sheets(sSource).Cells(row, column) = "=Translation!b" & trow
            End If
        Next i
    End If
    
    Next column
Next row


Application.Calculation = xlCalculationAutomatic

End Sub



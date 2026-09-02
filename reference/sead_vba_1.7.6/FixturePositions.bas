Attribute VB_Name = "FixturePositions"
'This macro was developed by p2w2.  http://p2w2.com/expert-in-microsoft-excel-consultants-consulting/index.php
'Please contact at CS@perceptive-analytics.com in case of any enquiry

Sub drawFixtures(dataSheet As String, lanewidth, MedianLength, FixtureHeight, NumberOfLanes, polespacing, polesetback, ArmLength, calculationmethod As String, PoleConfiguration As String, gridlength)

' to draw fixtures on road geometry page

Dim outputX() As Variant
Dim outputY() As Variant
outputX = FixturePosition(NumberOfLanes, PoleConfiguration, MedianLength, polespacing, lanewidth, polesetback, ArmLength, gridlength)(0)
outputY = FixturePosition(NumberOfLanes, PoleConfiguration, MedianLength, polespacing, lanewidth, polesetback, ArmLength, gridlength)(1)


For i = 0 To UBound(outputX)

Sheets(dataSheet).Cells(i + 2, 84) = outputX(i)
Sheets(dataSheet).Cells(i + 2, 85) = outputY(i)

Next
End Sub
Function FixturePosition(NumberOfLanes, PoleConfiguration As String, MedianLength, polespacing, lanewidth, polesetback, ArmLength, gridlength)

' to get positions of fixtures for different pole configurations
'Output of this function is an array with the x and y positions of each pole
'For arrangements with poles on both sides of the road, fixtures in the array alternate sides of the road

Dim FPArrayX(), FPArrayY()

If PoleConfiguration = "Single-side" Then
    numberoffixtures = (gridlength / polespacing)
Else
    numberoffixtures = (gridlength / polespacing) * 2
End If

'Array with the X and Y coordinates of each pole
ReDim FPArrayX(CInt(numberoffixtures) + 1)
ReDim FPArrayY(CInt(numberoffixtures) + 1)

'single pole position
If PoleConfiguration = "Single-side" Then
    For i = 0 To UBound(FPArrayX)
    FPArrayX(i) = i * polespacing
    FPArrayY(i) = 0 - polesetback + ArmLength
    Next
'opposite pole position
ElseIf PoleConfiguration = "Opposite" Then
    For i = 0 To UBound(FPArrayX)
    If i Mod 2 = 0 Then
        FPArrayX(i) = i * polespacing / 2
        FPArrayY(i) = 0 - polesetback + ArmLength
    Else
        FPArrayX(i) = (i - 1) * polespacing / 2
        FPArrayY(i) = NumberOfLanes * lanewidth + MedianLength + polesetback - ArmLength
    End If
Next
'median mounted pole position
ElseIf PoleConfiguration = "Median mounted" Then
    For i = 0 To UBound(FPArrayX)
        If i Mod 2 = 0 Then
            FPArrayX(i) = i * polespacing / 2
            FPArrayY(i) = (NumberOfLanes * lanewidth + MedianLength) / 2 - polesetback + ArmLength   'note, polesetback should probably be removed from here. Error validation is used in version 1.7.1 to prevent non-zero pole setbacks.
        Else
            FPArrayX(i) = (i - 1) * polespacing / 2
            FPArrayY(i) = (NumberOfLanes * lanewidth + MedianLength) / 2 + polesetback - ArmLength
        End If
    Next

'staggred pole position
ElseIf PoleConfiguration = "Staggered" Then
    For i = 0 To UBound(FPArrayX)
        If i Mod 2 = 0 Then
            FPArrayX(i) = i * polespacing / 2
            FPArrayY(i) = 0 - polesetback + ArmLength
        Else
            FPArrayX(i) = i * polespacing / 2
            FPArrayY(i) = NumberOfLanes * lanewidth + MedianLength + polesetback - ArmLength
        End If
    Next

End If

Dim ArrayXY(2)
ArrayXY(0) = FPArrayX
ArrayXY(1) = FPArrayY

FixturePosition = ArrayXY

End Function


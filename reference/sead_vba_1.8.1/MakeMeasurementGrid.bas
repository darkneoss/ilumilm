Attribute VB_Name = "MakeMeasurementGrid"
'This macro was developed by p2w2.  http://p2w2.com/expert-in-microsoft-excel-consultants-consulting/index.php
'Please contact at CS@perceptive-analytics.com in case of any enquiry

Function makeGrid(NumberOfLanes, calculationmethod As String, numberOfGridPoints, PoleConfiguration As String, MedianLength, polespacing, lanewidth)
Dim NumberOfGirdsPerLane, gridlength, medianYvalue
Dim GringSpacing
Dim flagformedian As Boolean

'Y is across the road. X is along the road.
' to determine number of y values in XY plane
If calculationmethod = "IES" Then
NumberOfGridsPerLane = 2
Else
NumberOfGridsPerLane = 3
End If

'calculating number of y values **FLAG take out logic?**
If NumberOfLanes Mod 2 = 0 Then
NumberOfYvalues = NumberOfGridsPerLane * (NumberOfLanes)
Else
NumberOfYvalues = (NumberOfGridsPerLane * (NumberOfLanes))
End If

'Define arrays for x values and y values in a XY plane
Dim Xvalues()
Dim Yvalues()
ReDim Xvalues(numberOfGridPoints)
ReDim Yvalues(NumberOfYvalues - 1)

GridSpacing = GridSpace(calculationmethod, polespacing)

' Assigning X values
If calculationmethod = "CIE" Then
Xvalues(0) = GridSpacing / 2
Else 'IES
Xvalues(0) = GridSpacing / 2
End If

For i = 1 To numberOfGridPoints
Xvalues(i) = Xvalues(i - 1) + GridSpacing
Next

'detemining median coordinates
If NumberOfLanes Mod 2 = 0 Then
    medianYvalue = (NumberOfLanes / 2) * lanewidth
    flagformedian = True
Else
    medianYvalue = 0
    flagformedian = False
End If

'Assigning Y values
' if IES method
If calculationmethod = "IES" Then
    Yvalues(0) = lanewidth / 4
    For i = 1 To NumberOfYvalues - 1
        Yvalues(i) = Yvalues(i - 1) + (lanewidth / 2)
        'Adding median length for lanes on otherside
        If Yvalues(i) >= medianYvalue And flagformedian Then
            Yvalues(i) = Yvalues(i) + MedianLength
            flagformedian = False
        End If
    Next

' if CIE method
Else
    Yvalues(0) = lanewidth / 6
    flagformedian = True            'FLAG this might be an error that prevents CIE method from having odd number of lanes
    For i = 1 To NumberOfYvalues - 1
        Yvalues(i) = Yvalues(i - 1) + (lanewidth / 3)
        'Adding median length for lanes on otherside
        If Yvalues(i) >= medianYvalue And flagformedian Then
            Yvalues(i) = Yvalues(i) + MedianLength
            flagformedian = False
        End If
    Next
End If

Dim ArrayXY(2)

ArrayXY(0) = Xvalues
ArrayXY(1) = Yvalues

makeGrid = ArrayXY
    
End Function



Function GridSpace(calculationmethod As String, polespacing)
If calculationmethod = "IES" Then
If (polespacing / 10) > 5 Then
GridSpace = 5
Else
GridSpace = (polespacing / 10)
End If
Else
If polespacing > 30 Then
GridSpace = 3
Else
If polespacing Mod 3 = 0 Then
GridSpace = 3
Else
GridSpace = polespacing / Int(polespacing / 3)
End If
End If
End If
End Function
Function TotalGridLength(calculationmethod As String, FixtureHeight, polespacing)
If calculationmethod = "IES" Then
TotalGridLength = 4 * polespacing
Else
TotalGridLength = 17 * FixtureHeight + polespacing ' 5h to 17h covered
End If
End Function
Sub plotRoadGeometry(chartName As String, dataSheet As String, lanewidth, MedianLength, FixtureHeight, NumberOfLanes, polespacing, calculationmethod As String, PoleConfiguration As String, polesetback)
Dim ngp As Integer
ngp = TotalGridLength(calculationmethod, FixtureHeight, polespacing) / GridSpace(calculationmethod, polespacing)
Dim outputX()
Dim outputY()
outputX = makeGrid(NumberOfLanes, calculationmethod, ngp, PoleConfiguration, MedianLength, polespacing, lanewidth)(0)
outputY = makeGrid(NumberOfLanes, calculationmethod, ngp, PoleConfiguration, MedianLength, polespacing, lanewidth)(1)


'draw x axis
For i = 0 To UBound(outputX)
    Sheets(dataSheet).Cells(i + 2, 1) = outputX(i)
Next

If calculationmethod = "IES" Then
    istart = WorksheetFunction.Match(polespacing, outputX, True) + 1
    iend = WorksheetFunction.Match(2 * polespacing, outputX, True)
ElseIf calculationmethod = "CIE" Then
    'start at what fixture
    startfixture = Int(5 * FixtureHeight / polespacing)
    startfixture = startfixture + 1
    istart = WorksheetFunction.Match(polespacing * startfixture, outputX, True) + 1
    iend = WorksheetFunction.Match(polespacing * (startfixture + 1), outputX, True)
    'iStart = WorksheetFunction.Match(5 * FixtureHeight, outputX, True)
    'iEnd = WorksheetFunction.Match(5 * FixtureHeight + polespacing, outputX, True)
End If

medianFlag = False
'Draw median
If NumberOfLanes Mod 2 = 0 Then             'Only if there are an even number of lanes
medianFlag = True
For i = 0 To UBound(outputX)
    Sheets(dataSheet).Cells(i + 2, 2) = (NumberOfLanes / 2) * lanewidth
    Sheets(dataSheet).Cells(i + 2, 3) = Sheets(dataSheet).Cells(i + 2, j + 2) + MedianLength
Next
End If

'draw road edges
For i = 0 To UBound(outputX)
Sheets(dataSheet).Cells(i + 2, 4) = 0
Sheets(dataSheet).Cells(i + 2, 5) = NumberOfLanes * lanewidth + MedianLength
Next

'draw lane edges
medianYvalue = (NumberOfLanes / 2) * lanewidth
oneLessForMedian = 0
If medianFlag = True Then oneLessForMedian = 1
For j = 1 To NumberOfLanes - 1 - oneLessForMedian
For i = 0 To UBound(outputX)
    Sheets(dataSheet).Cells(i + 2, 5 + j) = (j) * lanewidth
    If Sheets(dataSheet).Cells(i + 2, 5 + j) >= medianYvalue And medianFlag = True Then
        Sheets(dataSheet).Cells(i + 2, 5 + j) = Sheets(dataSheet).Cells(i + 2, 5 + j) + MedianLength + lanewidth
    End If
Next i
Next j

'draw Gridlines
For i = istart To iend
For j = 0 To UBound(outputY)
    Sheets(dataSheet).Cells(i + 1, j + 24) = outputY(j)
Next
Next


Sheets("Road Geometry").Activate
Sheets("Road Geometry").Unprotect
'adjusting chart y axis for look of road
ActiveSheet.ChartObjects(chartName).Activate
ActiveChart.Axes(xlValue).Select
ActiveChart.Axes(xlValue).MaximumScale = NumberOfLanes * lanewidth + MedianLength + polesetback + 1
ActiveChart.Axes(xlValue).MinimumScale = -polesetback - 1
ActiveChart.Axes(xlCategory).MaximumScale = outputX(UBound(outputX))
Sheets("Road Geometry").Protect
End Sub

Sub baselineUpdate()
bCieIesChoiceForm.Label1.Caption = Sheets("Translation").Range("tCieIesChoice").Value
bCieIesChoiceForm.Show
End Sub
Sub upgradeUpdate()
uCieIesChoiceForm.Label1.Caption = Sheets("Translation").Range("tCieIesChoice").Value
uCieIesChoiceForm.Show
End Sub
Sub baselinePlot(choice As String)

'---------------------------------------------------------------------------------------
'start by performing some input validation
'Perform some input validation
Dim rMaster(1 To 16) As Range
Dim iCheck As Long
Dim missingFlag As Boolean

Application.Calculation = xlCalculationManual

missingFlag = False

Set rMaster(1) = wksRoadGeometry.Range("bNumLanes")
Set rMaster(2) = wksRoadGeometry.Range("bLaneWidth")
Set rMaster(3) = wksRoadGeometry.Range("bMedianWidth")
Set rMaster(4) = wksRoadGeometry.Range("bMountingHeight")
Set rMaster(5) = wksRoadGeometry.Range("bPoleSpacing")
Set rMaster(6) = wksRoadGeometry.Range("bPoleSetback")
Set rMaster(7) = wksRoadGeometry.Range("bArmLength")
Set rMaster(8) = wksRoadGeometry.Range("bFixtureArrangement")

Set rMaster(9) = wksRoadGeometry.Range("uNumLanes")
Set rMaster(10) = wksRoadGeometry.Range("uLaneWidth")
Set rMaster(11) = wksRoadGeometry.Range("uMedianWidth")
Set rMaster(12) = wksRoadGeometry.Range("uMountingHeight")
Set rMaster(13) = wksRoadGeometry.Range("uPoleSpacing")
Set rMaster(14) = wksRoadGeometry.Range("uPoleSetback")
Set rMaster(15) = wksRoadGeometry.Range("uArmLength")
Set rMaster(16) = wksRoadGeometry.Range("uFixtureArrangement")
                
For iCheck = LBound(rMaster) To UBound(rMaster)
    If IsEmpty(rMaster(iCheck).Value) Then missingFlag = True
Next iCheck

If missingFlag = True Then
    prompt = Sheet25.Range("tMissingRoadGeometry")
    MsgBox (prompt)
    Exit Sub
End If
'---------------------------------------------------------------------------------------




Application.ScreenUpdating = False
'clear contents on
Sheets("Chart Data Baseline").Activate
Rows("2:10000").Select
Selection.ClearContents

Dim lanewidth, MedianLength, FixtureHeight, NumberOfLanes, polespacing
lanewidth = Sheets("Road Geometry").Range("bLaneWidth").Value
MedianLength = Sheets("Road Geometry").Range("bMedianWidth").Value
FixtureHeight = Sheets("Road Geometry").Range("bMountingHeight").Value
NumberOfLanes = Sheets("Road Geometry").Range("bNumLanes").Value
polespacing = Sheets("Road Geometry").Range("bPoleSpacing").Value

Dim polesetback, ArmLength
polesetback = Sheets("Road Geometry").Range("bPoleSetback").Value
ArmLength = Sheets("Road Geometry").Range("bArmLength").Value


Dim calcMethod As String, poleconfig As String, dataSheet As String, chartName As String
calcMethod = choice
poleconfig = Sheets("Road Geometry").Range("bFixtureArrangement").Value
dataSheet = "Chart Data Baseline"
chartName = "Baseline"

' puting grid spacing in road geometry
Sheets("Road Geometry").Range("bGridSpacing").Value = GridSpace(calcMethod, polespacing)

gridlength = TotalGridLength(calcMethod, FixtureHeight, polespacing)

Debug.Print Now()
Call plotRoadGeometry(chartName, dataSheet, lanewidth, MedianLength, FixtureHeight, NumberOfLanes, polespacing, calcMethod, poleconfig, polesetback)
Debug.Print Now()
Call drawFixtures(dataSheet, lanewidth, MedianLength, FixtureHeight, NumberOfLanes, polespacing, polesetback, ArmLength, calcMethod, poleconfig, gridlength)
Debug.Print Now()

Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub
Sub upgradePlot(choice As String)

'---------------------------------------------------------------------------------------
'start by performing some input validation
'Perform some input validation
Dim rMaster(1 To 16) As Range
Dim iCheck As Long
Dim missingFlag As Boolean

Application.Calculation = xlCalculationManual

missingFlag = False

Set rMaster(1) = wksRoadGeometry.Range("bNumLanes")
Set rMaster(2) = wksRoadGeometry.Range("bLaneWidth")
Set rMaster(3) = wksRoadGeometry.Range("bMedianWidth")
Set rMaster(4) = wksRoadGeometry.Range("bMountingHeight")
Set rMaster(5) = wksRoadGeometry.Range("bPoleSpacing")
Set rMaster(6) = wksRoadGeometry.Range("bPoleSetback")
Set rMaster(7) = wksRoadGeometry.Range("bArmLength")
Set rMaster(8) = wksRoadGeometry.Range("bFixtureArrangement")

Set rMaster(9) = wksRoadGeometry.Range("uNumLanes")
Set rMaster(10) = wksRoadGeometry.Range("uLaneWidth")
Set rMaster(11) = wksRoadGeometry.Range("uMedianWidth")
Set rMaster(12) = wksRoadGeometry.Range("uMountingHeight")
Set rMaster(13) = wksRoadGeometry.Range("uPoleSpacing")
Set rMaster(14) = wksRoadGeometry.Range("uPoleSetback")
Set rMaster(15) = wksRoadGeometry.Range("uArmLength")
Set rMaster(16) = wksRoadGeometry.Range("uFixtureArrangement")
                
For iCheck = LBound(rMaster) To UBound(rMaster)
    If IsEmpty(rMaster(iCheck).Value) Then missingFlag = True
Next iCheck

If missingFlag = True Then
    prompt = Sheet25.Range("tMissingRoadGeometry")
    MsgBox (prompt)
    Exit Sub
End If
'---------------------------------------------------------------------------------------




Application.ScreenUpdating = False
'clear contents on
Sheets("Chart Data Upgrade").Activate
Rows("2:10000").Select
'Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents

Dim lanewidth, MedianLength, FixtureHeight, NumberOfLanes, polespacing
lanewidth = Sheets("Road Geometry").Range("uLaneWidth").Value
MedianLength = Sheets("Road Geometry").Range("uMedianWidth").Value
FixtureHeight = Sheets("Road Geometry").Range("uMountingHeight").Value
NumberOfLanes = Sheets("Road Geometry").Range("uNumLanes").Value
polespacing = Sheets("Road Geometry").Range("uPoleSpacing").Value

Dim polesetback, ArmLength
polesetback = Sheets("Road Geometry").Range("uPoleSetback").Value
ArmLength = Sheets("Road Geometry").Range("uArmLength").Value

Dim calcMethod As String, poleconfig As String, dataSheet As String, chartName As String
'calcMethod = Sheets("FixtureData").Range("iescieGraphChoice").Value
calcMethod = choice
poleconfig = Sheets("Road Geometry").Range("uFixtureArrangement").Value
dataSheet = "Chart Data Upgrade"
chartName = "Upgrade"

' puting grid spacing in road geometry
Sheets("Road Geometry").Range("uGridSpacing").Value = GridSpace(calcMethod, polespacing)

gridlength = TotalGridLength(calcMethod, FixtureHeight, polespacing)

Call plotRoadGeometry(chartName, dataSheet, lanewidth, MedianLength, FixtureHeight, NumberOfLanes, polespacing, calcMethod, poleconfig, polesetback)
Call drawFixtures(dataSheet, lanewidth, MedianLength, FixtureHeight, NumberOfLanes, polespacing, polesetback, ArmLength, calcMethod, poleconfig, gridlength)
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic

End Sub

Sub goToLuminance()
Application.GoTo ActiveSheet.Range("AZ1"), True
End Sub
Sub goToIlluminance()
Application.GoTo ActiveSheet.Range("A1"), True
End Sub
Sub drawGrid()
' to plot points of measurement grid. Not being used anywhere

Dim ngp As Integer
Dim lanewidth As Integer
Dim MedianLength As Integer
Dim FixtureHeight As Integer
Dim NumberOfLanes As Integer
Dim polespacing As Integer
Dim calculationmethod As String
Dim poleconfig As String
NumberOfLanes = 20
lanewidth = 3
poleconfig = "opposite"
calculationmethod = "CIE"
MedianLength = 2
polespacing = 5
FixtureHeight = 30

ngp = TotalGridLength(calculationmethod, FixtureHeight, polespacing) / GridSpace(calculationmethod, polespacing)
Dim outputX() As Variant
Dim outputY() As Variant
outputX = makeGrid(NumberOfLanes, calculationmethod, ngp, poleconfig, MedianLength, polespacing, lanewidth)(0)
outputY = makeGrid(NumberOfLanes, calculationmethod, ngp, poleconfig, MedianLength, polespacing, lanewidth)(1)


'Draw measurement grid
For i = 0 To UBound(outputX)
For j = 0 To UBound(outputY)
    Sheets("Chart Data").Cells(i + 1, 1) = outputX(i)
    Sheets("Chart Data").Cells(i + 1, j + 2) = outputY(j)
Next
Next

'Draw median
For i = 0 To UBound(outputX)
    Sheets("Chart Data").Cells(i + 1, j + 2) = (NumberOfLanes / 2) * lanewidth
    Sheets("Chart Data").Cells(i + 1, j + 3) = Sheets("Chart Data").Cells(i + 1, j + 2) + MedianLength
Next
End Sub


Attribute VB_Name = "IlluminanceAndLuminance"
'This macro was developed by p2w2.  http://p2w2.com/expert-in-microsoft-excel-consultants-consulting/index.php
'Please contact at CS@perceptive-analytics.com in case of any enquiry
Global a


Sub finalMatrices()

' Calculate final luminance and illuminace at each grid point from all the fixtures
Dim lanewidth, MedianLength, FixtureHeight, NumberOfLanes, polespacing
Dim polesetback, ArmLength
Dim calcMethod As String, poleconfig As String


'Clearing all illuminance and luminance sheets
Sheets("Illuminance Calcs").Select
Rows("12:12").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents
Range("A1").Select
Sheets("Luminance Calcs").Select
Rows("12:12").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.ClearContents
Range("A1").Select
Sheets("Luminance Calcs CIE").Select
Rows("12:5000").Select
Selection.ClearContents
Range("A1").Select

'Taking road geometry values for Baseline vs. Upgrade
If Sheets("FixtureData").Range("Base_Upgrade_Choice").Value = "Baseline" Then
    ExitFlag = False
    NumberOfLanes = Sheets("Road Geometry").Range("bNumLanes").Value
    lanewidth = Sheets("Road Geometry").Range("bLaneWidth").Value
    MedianLength = Sheets("Road Geometry").Range("bMedianWidth").Value
    FixtureHeight = Sheets("Road Geometry").Range("bMountingHeight").Value
    polespacing = Sheets("Road Geometry").Range("bPoleSpacing").Value
    polesetback = Sheets("Road Geometry").Range("bPoleSetback").Value
    ArmLength = Sheets("Road Geometry").Range("bArmLength").Value
    If VarType(Sheets("Road Geometry").Range("bFixtureArrangement").Value) = vbError Then
        ExitFlag = True
    Else
        poleconfig = Sheets("Road Geometry").Range("bFixtureArrangement").Value
    End If
ElseIf Sheets("FixtureData").Range("A6").Value = "Upgrade" Then
    ExitFlag = False
    NumberOfLanes = Sheets("Road Geometry").Range("uNumLanes").Value
    lanewidth = Sheets("Road Geometry").Range("uLaneWidth").Value
    MedianLength = Sheets("Road Geometry").Range("uMedianWidth").Value
    FixtureHeight = Sheets("Road Geometry").Range("uMountingHeight").Value
    polespacing = Sheets("Road Geometry").Range("uPoleSpacing").Value
    polesetback = Sheets("Road Geometry").Range("uPoleSetback").Value
    ArmLength = Sheets("Road Geometry").Range("uArmLength").Value
    poleconfig = Sheets("Road Geometry").Range("uFixtureArrangement").Value
End If
If ExitFlag = True Then
    ErrorMessage = Sheet25.Range("poleconfigError").Value
    MsgBox (ErrorMessage)
    Exit Sub
End If

'Determine calculation method
calcMethod = Sheets("FixtureData").Range("iescieGraphChoice").Value

Dim LLF
LLF = Sheets("FixtureData").Range("H6").Value 'light loss factor

Dim outputXY()
Dim gammaArray()
Dim phi()

'number of grid points?
Dim ngp As Integer
ngp = TotalGridLength(calcMethod, FixtureHeight, polespacing) / GridSpace(calcMethod, polespacing)

'Making the grid and getting its postions into an array
'This grid has the XY coordinates of the entire road array (regardless of wheter values are calculated at those points). Logic in the angle and illuminance/luminance calcs decides which of these to use.
outputXY = makeGrid(NumberOfLanes, calcMethod, ngp, poleconfig, MedianLength, polespacing, lanewidth)

'Output the X and Y coordinates of each individual fixture.
'X is along the road. Y is across the road.
Dim fixtureX()
Dim fixtureY()
gridlength = TotalGridLength(calcMethod, FixtureHeight, polespacing)
fixtureX = FixturePosition(NumberOfLanes, poleconfig, MedianLength, polespacing, lanewidth, polesetback, ArmLength, gridlength)(0)
fixtureY = FixturePosition(NumberOfLanes, poleconfig, MedianLength, polespacing, lanewidth, polesetback, ArmLength, gridlength)(1)
'**FLAG performance speedup - FixturePosition function recalculates each time it is called

'Tilt------------------------
'Old Way
'tiltDegrees = 10
'tiltRadians = tiltDegrees * WorksheetFunction.Pi / 180
'Dim rayLength As Double, extraY As Double
'rayLength = FixtureHeight / Cos(tiltRadians)    'ray is the imaginary line extending perpendicularly out of the fixture. If the tilt is zero, this is the same as fixture height.
'extraY = FixtureHeight * Tan(tiltRadians)       'Additional distance along cross-road axis due to tilt. X is along the road. Y is across the road.

'new way
tiltOnX = 0 / 180 * WorksheetFunction.Pi        'the up down tilt
tiltOnY = 0 / 180 * WorksheetFunction.Pi        'towards or away from observer, i.e. twisting the arm
tiltOnZ = 0 / 180 * WorksheetFunction.Pi        'twisting the pole
'Tilt--------------------------

' Every array prefixed L is used for Illuminance calculations and prefixed with R is used for Luminance calculations
Dim larray()
Dim LarrayMatrix()
ReDim LarrayMatrix(UBound(fixtureX))

Dim Rarray()
Dim RarrayMatrix()
ReDim RarrayMatrix(UBound(fixtureX))

Dim tempArray1()
Dim illuminanceFixture()
ReDim illuminanceFixture(UBound(fixtureX))

Dim temparray2()
Dim luminanceFixture()
ReDim luminanceFixture(UBound(fixtureX))

Dim LthisArray()
Dim LsumArray()
Dim LnextArray()
Dim RthisArray()
Dim RsumArray()
Dim RnextArray()
Dim rownum
rownum = 13

' If IES method
If calcMethod = "IES" Then
    '******************************************************************************************************
    '*******************************if IES method**********************************************************
    '******************************************************************************************************
    ' Calculate the angles, luminous intensity, reflectance and finally luminance/illuminance at each grid point for the contributions of each fixture. Save each array of contribution of all grid points for each fixture into an array of arrays.
    'Note, outputXY is the coordinates of each grid point (included unused gridpoints).
    'Each phi, gamma, beta array is redim'd to just the x coordinates that are between the two primary fixtures (but their location from the first calculated fixture is the same position in the array as in the outputXY array)
    '
    'for instance, if some poles are before and after the primary poles and the first pole is at (9) the two arrays would look like this:
    'X directionin outputXY: (1)(2) (3) (4) (5) (6) (7) (8) (9) (10) (11) (12) (13) (14) (15) (16)
    'x direction in others :                                (9) (10) (11) (12)
    
    
    For k = LBound(fixtureX) To UBound(fixtureX)
        'Angle calculations
        phi = anglePhi(fixtureX(k), fixtureY(k), outputXY, polespacing, FixtureHeight, 0, 0, 0, calcMethod)
        phiArrayForITable = anglePhiWithTilt(fixtureX(k), fixtureY(k), outputXY, polespacing, FixtureHeight, tiltOnX, tiltOnY, tiltOnZ, calcMethod)
        gammaArray = angleGamma(fixtureX(k), fixtureY(k), outputXY, polespacing, FixtureHeight, 0, 0, 0, calcMethod)
        gammaArrayForITable = angleGammaWithTilt(fixtureX(k), fixtureY(k), outputXY, polespacing, FixtureHeight, tiltOnX, tiltOnY, tiltOnZ, calcMethod)
        betaArray = angleBeta(phi(), calcMethod, fixtureX(k), fixtureY(k), outputXY, polespacing, lanewidth, FixtureHeight, 0)
        
        'Luminous intensity calculations using quadratic interpolation
        larray = LintensityMatrix(ngp, poleconfig, fixtureX(k), fixtureY(k), outputXY, polespacing, FixtureHeight, calcMethod, phiArrayForITable, gammaArrayForITable) 'FLAG
        LarrayMatrix(k) = larray
        
        'Sheets("test1").Activate
        'For i = LBound(larray(), 1) To UBound(larray(), 1)
        'For j = 0 To UBound(larray(), 2)
        'Cells(i + 101, j + 1) = larray(i, j)
        'Next
        'Next
    
        'Road reflectance using quadratic interpolation
        Rarray = RMatrix(gridlength, poleconfig, fixtureX(k), fixtureY(k), outputXY(), polespacing, FixtureHeight, calcMethod, betaArray, gammaArray)
        RarrayMatrix(k) = Rarray
        
        ' Illuminance at every grid point by fixture k
        tempArray1 = Illum(larray, gammaArray, LLF, FixtureHeight)
        illuminanceFixture(k) = tempArray1
    
        ' Luminance at every grid point by fixture k
        temparray2 = Lum(larray, gammaArray, Rarray, LLF, FixtureHeight)
        luminanceFixture(k) = temparray2
    Next
    
    '*************************** got luminance and illuminance matrices at every grid point by every fixture till now *********
    ReDim LsumArray(LBound(gammaArray(), 1) To UBound(gammaArray(), 1), LBound(gammaArray(), 2) To UBound(gammaArray(), 2))
    ReDim RsumArray(LBound(gammaArray(), 1) To UBound(gammaArray(), 1), LBound(gammaArray(), 2) To UBound(gammaArray(), 2))
    
    'clearing the previous array before calculations
    For p = LBound(gammaArray, 1) To UBound(gammaArray, 1)
        For q = LBound(gammaArray, 2) To UBound(gammaArray, 2)
            LsumArray(p, q) = 0
            RsumArray(p, q) = 0
        Next
    Next
    
    'Running arrays for all fixtures to sum luminance and illuminacne
    'Get each fixture array (..thisArray) and add each of it's values to the sumArray
    For i = LBound(illuminanceFixture) To UBound(illuminanceFixture)
        LthisArray = illuminanceFixture(i)
        RthisArray = luminanceFixture(i)
        
        For j = LBound(LsumArray, 1) To UBound(LsumArray, 1)
            For k = LBound(LsumArray, 2) To UBound(LsumArray, 2)
                    'Sum of illuminance from every fixture
                    LsumArray(j, k) = LsumArray(j, k) + LthisArray(j, k)
                
                    'Sum of luminance from every fixture
                    RsumArray(j, k) = RsumArray(j, k) + RthisArray(j, k)
            Next
        Next
    Next
    
    ' getting the values into luminance and iluminance sheets
    Sheets("Illuminance Calcs").[B13].Resize(UBound(LsumArray, 1) - LBound(LsumArray, 1) + 1, UBound(LsumArray, 2) - LBound(LsumArray, 2) + 1).Value = LsumArray
    Sheets("Luminance Calcs").[B13].Resize(UBound(RsumArray, 1) - LBound(RsumArray, 1) + 1, UBound(RsumArray, 2) - LBound(RsumArray, 2) + 1).Value = RsumArray
    
    'add the labels
    Dim t As Integer
    Dim s As Integer
    s = 2
    For t = 1 To NumberOfLanes
        Sheets("Illuminance Calcs").Cells(12, s).Value = "Lane " & t & " - 1/4 lane"
        Sheets("Illuminance Calcs").Cells(12, s + 1).Value = "Lane " & t & " - 3/4 lane"
        Sheets("Luminance Calcs").Cells(12, s).Value = "Lane " & t & " - 1/4 lane"
        Sheets("Luminance Calcs").Cells(12, s + 1).Value = "Lane " & t & " - 3/4 lane"
        s = s + 2
    Next t

'******************************************************************************************************
'*******************************if CIE method**********************************************************
'******************************************************************************************************
ElseIf calcMethod = "CIE" Then
    '-----------
    'Luminance
    '-----------
    'running for all observer locations
    For m = 1 To NumberOfLanes
        yo = (2 * m - 1) * lanewidth / 2                    'yo = y of observer, in the center of each lane in turn
        If yo > NumberOfLanes * lanewidth / 2 Then
        yo = yo + MedianLength
        End If
        
        ' Running for all fixtures
        For k = LBound(fixtureX) To UBound(fixtureX)
        
        ' running for all observer locations.
        'Different luminance for different observer location is put into a sheet called Luminance Calcs CIE
        ' Least average values and uniformity values are taken from this sheet into main luminance calcs sheet
        
        'getting angle matrices
        phi = anglePhi(fixtureX(k), fixtureY(k), outputXY, polespacing, FixtureHeight, 0, 0, 0, calcMethod)
        phiArrayForITable = anglePhiWithTilt(fixtureX(k), fixtureY(k), outputXY, polespacing, FixtureHeight, tiltOnX, tiltOnY, tiltOnZ, calcMethod)
        gammaArray = angleGamma(fixtureX(k), fixtureY(k), outputXY, polespacing, FixtureHeight, 0, 0, 0, calcMethod)
        gammaArrayForITable = angleGammaWithTilt(fixtureX(k), fixtureY(k), outputXY, polespacing, FixtureHeight, tiltOnX, tiltOnY, tiltOnZ, calcMethod)
        betaArray = angleBeta(phi(), calcMethod, fixtureX(k), fixtureY(k), outputXY, polespacing, lanewidth, FixtureHeight, yo)
        
        'getting luminous intensity using quadratic interpolation
        larray = LintensityMatrix(ngp, poleconfig, fixtureX(k), fixtureY(k), outputXY, polespacing, FixtureHeight, calcMethod, phiArrayForITable, gammaArrayForITable)
        LarrayMatrix(k) = larray
        
        'Road reflectance using quadratic interpolation
        Rarray = RMatrix(gridlength, poleconfig, fixtureX(k), fixtureY(k), outputXY(), polespacing, FixtureHeight, calcMethod, betaArray, gammaArray)
        RarrayMatrix(k) = Rarray
        
        'Calculating luminance
        temparray2 = Lum(larray, gammaArray, Rarray, LLF, FixtureHeight)
        luminanceFixture(k) = temparray2
        
        Next k  'next fixture
        
        ' sum of luminnance from all fixtures
        ReDim RsumArray(LBound(gammaArray(), 1) To UBound(gammaArray(), 1), LBound(gammaArray(), 2) To UBound(gammaArray(), 2))
        
        ' clearing sum arrays
        For p = LBound(gammaArray, 1) To UBound(gammaArray, 1)
            For q = LBound(gammaArray, 2) To UBound(gammaArray, 2)
                  RsumArray(p, q) = 0
            Next q
        Next p
        
        
        For i = LBound(luminanceFixture) To UBound(luminanceFixture)
        RthisArray = luminanceFixture(i)
            For j = LBound(RsumArray, 1) To UBound(RsumArray, 1)
                For k = LBound(RsumArray, 2) To UBound(RsumArray, 2)
                    ' Summing up luminace at each grid point from one observer location
                    RsumArray(j, k) = RsumArray(j, k) + RthisArray(j, k)
                Next k
            Next j
        Next i
        
        'getting luminance for one oberserver location into excel sheet
        Sheets("Luminance Calcs CIE").Range("F" & rownum).Resize(UBound(RsumArray, 1) - LBound(RsumArray, 1) + 1, UBound(RsumArray, 2) - LBound(RsumArray, 2) + 1).Value = RsumArray
        
        Dim c As Integer
        c = UBound(RsumArray, 2) + 1
        Dim sc As Integer
        
        Dim colNo As Integer
        colNo = 53
        Dim arr As Variant
        arr = RsumArray
        'Find max/min of each lane at midpoint
        ReDim Preserve arr(LBound(RsumArray, 1) To UBound(RsumArray, 1), 1)
        For sc = 2 To c - 1 Step 3
            arr = Application.WorksheetFunction.Index(RsumArray, 0, sc)
            Sheets("Luminance Calcs CIE").Cells(rownum, colNo).Value = Application.WorksheetFunction.Min(arr)
            Sheets("Luminance Calcs CIE").Cells(rownum + 1, colNo).Value = Application.WorksheetFunction.Max(arr)
            If Application.WorksheetFunction.Min(arr) = 0 Then
                Sheets("Luminance Calcs CIE").Cells(rownum + 2, colNo).Value = ""
            Else
                Sheets("Luminance Calcs CIE").Cells(rownum + 2, colNo).Value = Application.WorksheetFunction.Min(arr) / Application.WorksheetFunction.Max(arr)  'changed from max/min ***
            End If
            colNo = colNo + 1
        Next sc
        
        'Find maximimum of the min/max for all lanes
        Dim overallmax
        Dim arr1 As Range
        Sheets("Luminance Calcs CIE").Activate
        Set arr1 = Range(Cells(rownum + 2, 53), Cells(rownum + 2, colNo - 1))
        overallmax = Application.WorksheetFunction.Max(arr1)
        
        Sheets("Luminance Calcs CIE").Range("e" & rownum).Value = overallmax
        
        'Other calcs
        Sheets("Luminance Calcs CIE").Range("a" & rownum).Value = Application.WorksheetFunction.Average(RsumArray)
        Sheets("Luminance Calcs CIE").Range("b" & rownum).Value = Application.WorksheetFunction.Min(RsumArray)
        Sheets("Luminance Calcs CIE").Range("c" & rownum).Value = Application.WorksheetFunction.Max(RsumArray)
        If Application.WorksheetFunction.Min(RsumArray) = 0 Then
            Sheets("Luminance Calcs CIE").Range("d" & rownum).Value = ""
        Else
            Sheets("Luminance Calcs CIE").Range("d" & rownum).Value = Application.WorksheetFunction.Min(RsumArray) / Application.WorksheetFunction.Average(RsumArray)
        End If
        
        'If Application.WorksheetFunction.Max(RsumArray) = 0 Then
        'Sheets("Luminance Calcs CIE").Range("e" & rownum).Value = ""
        'Else
        'Sheets("Luminance Calcs CIE").Range("e" & rownum).Value = Application.WorksheetFunction.Min(RsumArray) / Application.WorksheetFunction.Max(RsumArray)
        'End If
        
        rownum = rownum + UBound(RsumArray, 1) - LBound(RsumArray, 1) + 1
    Next
    
    Dim t1 As Integer
    Dim s1 As Integer
    s1 = 6
    For t1 = 1 To NumberOfLanes
        Sheets("Luminance Calcs CIE").Cells(12, s1).Value = "L " & t1 & " - 1/6"
        Sheets("Luminance Calcs CIE").Cells(12, s1 + 1).Value = "L " & t1 & " - 3/6"
        Sheets("Luminance Calcs CIE").Cells(12, s1 + 2).Value = "L " & t1 & " - 5/6"
        s1 = s1 + 3
    Next t1
    
    '-----------
    'Illuminance
    '-----------
    'Grid changed to take only 3 longitudnal points across road instead of lane.
    LLF = Sheets("FixtureData").Range("H6").Value
    
    If MedianLength = 0 Then
        Dim outputY1(2)
        outputY1(0) = (lanewidth * NumberOfLanes + MedianLength) / 6
        outputY1(1) = (lanewidth * NumberOfLanes + MedianLength) / 2
        outputY1(2) = (lanewidth * NumberOfLanes + MedianLength) * 5 / 6
        outputXY(1) = outputY1
    Else
        Dim outputY2(5)
        outputY2(0) = (lanewidth * NumberOfLanes / 2) / 6
        outputY2(1) = (lanewidth * NumberOfLanes / 2) / 2
        outputY2(2) = (lanewidth * NumberOfLanes / 2) * 5 / 6
        outputY2(3) = (lanewidth * NumberOfLanes / 2) / 6 + MedianLength + (lanewidth * NumberOfLanes / 2)
        outputY2(4) = (lanewidth * NumberOfLanes / 2) / 2 + MedianLength + (lanewidth * NumberOfLanes / 2)
        outputY2(5) = (lanewidth * NumberOfLanes / 2) * 5 / 6 + MedianLength + (lanewidth * NumberOfLanes / 2)
        outputXY(1) = outputY2
    End If
    
    ' same routine again, but a different grid.
    For k = LBound(fixtureX) To UBound(fixtureX)
        phi = anglePhi(fixtureX(k), fixtureY(k), outputXY, polespacing, FixtureHeight, tiltOnX, tiltOnY, tiltOnZ, calcMethod)
        gammaArray = angleGamma(fixtureX(k), fixtureY(k), outputXY, polespacing, FixtureHeight, tiltOnX, tiltOnY, tiltOnZ, calcMethod)
        betaArray = angleBeta(phi(), calcMethod, fixtureX(k), fixtureY(k), outputXY, polespacing, lanewidth, FixtureHeight, 0)
        
        'removing all the luminaries outside 5H distance
        If outputXY(0)(LBound(gammaArray)) - fixtureX(k) > 5 * FixtureHeight Or fixtureX(k) - outputXY(0)(UBound(gammaArray)) > 5 * FixtureHeight Then
        LLF = 0
        Else
        LLF = Sheets("FixtureData").Range("H6").Value
        End If
        
        larray = LintensityMatrix(ngp, poleconfig, fixtureX(k), fixtureY(k), outputXY, polespacing, FixtureHeight, calcMethod, phi, gammaArray)
        LarrayMatrix(k) = larray
        
        tempArray1 = Illum(larray, gammaArray, LLF, FixtureHeight)
        illuminanceFixture(k) = tempArray1
    Next
    
    '*************************** got luminance and illuminance matrices at every grid point by every fixture till now *********
    ' sum of luminnance from all fixtures
    ReDim LsumArray(LBound(gammaArray(), 1) To UBound(gammaArray(), 1), LBound(gammaArray(), 2) To UBound(gammaArray(), 2))
    For p = LBound(gammaArray, 1) To UBound(gammaArray, 1)
        For q = LBound(gammaArray, 2) To UBound(gammaArray, 2)
            LsumArray(p, q) = 0
            'RsumArray(p, q) = 0
        Next q
    Next p
    
    For i = LBound(illuminanceFixture) To UBound(illuminanceFixture)
    LthisArray = illuminanceFixture(i)
        For j = LBound(LsumArray, 1) To UBound(LsumArray, 1)
            For k = LBound(LsumArray, 2) To UBound(LsumArray, 2)
                    LsumArray(j, k) = LsumArray(j, k) + LthisArray(j, k)
            Next
        Next
    Next
    
    Sheets("Illuminance Calcs").[B13].Resize(UBound(LsumArray, 1) - LBound(LsumArray, 1) + 1, UBound(LsumArray, 2) - LBound(LsumArray, 2) + 1).Value = LsumArray
    'Sheets("Luminance Calcs").[B13].Resize(UBound(RsumArray, 1) - LBound(RsumArray, 1) + 1, UBound(RsumArray, 2) - LBound(RsumArray, 2) + 1).Value = RsumArray
        Sheets("Illuminance Calcs").Cells(12, 2).Value = "R - 1/6"
        Sheets("Illuminance Calcs").Cells(12, 3).Value = "R - 3/6"
        Sheets("Illuminance Calcs").Cells(12, 4).Value = "R - 5/6"
        
    If MedianLength > 0 Then
        Sheets("Illuminance Calcs").Cells(12, 5).Value = "R - 1/6"
        Sheets("Illuminance Calcs").Cells(12, 6).Value = "R - 3/6"
        Sheets("Illuminance Calcs").Cells(12, 7).Value = "R - 5/6"
    End If
    
End If  'CIE vs. IES
End Sub

Sub PrintArray(Data(), Cl As Range)
    Cl.Resize(UBound(Data, 1), UBound(Data, 2) + 1) = Data
End Sub


Function getDimension(var As Variant) As Integer
On Error GoTo Err:
    Dim i As Integer
    Dim tmp As Integer
    i = 0
    Do While True:
        i = i + 1
        tmp = UBound(var, i)
    Loop
Err:
    getDimension = i - 1
End Function

Function Illum(larray(), gammaArray(), LLF, FixtureHeight) As Variant
Dim tarray()
ReDim tarray(LBound(gammaArray(), 1) To UBound(gammaArray(), 1), LBound(gammaArray(), 2) To UBound(gammaArray(), 2))

For i = LBound(gammaArray(), 1) To UBound(gammaArray(), 1)
    For j = LBound(gammaArray(), 2) To UBound(gammaArray(), 2)

        tarray(i, j) = larray(i, j) * Cos(gammaArray(i, j) * WorksheetFunction.Pi / 180) ^ 3 * LLF / FixtureHeight ^ 2
        'MsgBox tarray(i, j)
    Next
Next

Illum = tarray

End Function
Function Lum(larray(), gammaArray(), Rarray(), LLF, FixtureHeight) As Variant
Dim tarray()
ReDim tarray(LBound(gammaArray(), 1) To UBound(gammaArray(), 1), LBound(gammaArray(), 2) To UBound(gammaArray(), 2))

For i = LBound(gammaArray(), 1) To UBound(gammaArray(), 1)
    For j = LBound(gammaArray(), 2) To UBound(gammaArray(), 2)

        tarray(i, j) = larray(i, j) * Rarray(i, j) * LLF / ((FixtureHeight ^ 2) * 10000)
    Next
Next

Lum = tarray

End Function



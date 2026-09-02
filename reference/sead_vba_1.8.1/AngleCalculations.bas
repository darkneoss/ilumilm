Attribute VB_Name = "AngleCalculations"
'This macro was developed by p2w2.  http://p2w2.com/expert-in-microsoft-excel-consultants-consulting/index.php
'Please contact at CS@perceptive-analytics.com in case of any enquiry

Function angleGammaWithTilt(fixtureX, fixtureY, backwardsFlag As Boolean, tiltOnX, tiltOnY, tiltOnZ, gridXY, calculationmethod, intBaselineUpgradeChoice, geometryValues()) As Variant
'gamma is the vertical angle coming out of the fixture, between the line directly to the ground and the line extending from fixture to grid point
'When the fixture is tilted, the line directly to the ground is instead the line extending straight out of the center of the fixture.

'Variables required:
'fixtureX is the x coordinate of the fixture currently being calculated
'fixtureY is the y coordinate of the same
'backwardsFlag determines which direction the fixture faces(based on which side of the road it is on)
'tiltOnX, Y, and Z are the tilt in radians of the fixture, for the current calculation
'gridXY() is an array containing the x/y coordinates of all grid points from the start to end of the road. istart and iend are used to select only the actual calculated grid bounds to be used here.
'calculationmethod is either CIE or IES
'intBaselineUpgradeChoice is used for choosing which road geometry data to use. 1 = Baseline, 2 = Upgrade
'geometryValues() is an array containing the road geometry data. Column 0 is header names (for reference), column 1 is baseline, 2 is upgrade. Rows are:
    'geometryValues(1, 0) = "NumLanes"
    'geometryValues(2, 0) = "LaneWidth"
    'geometryValues(3, 0) = "MedianWidth"
    'geometryValues(4, 0) = "MountingHeight"
    'geometryValues(5, 0) = "PoleSpacing"
    'geometryValues(6, 0) = "PoleSetback"
    'geometryValues(7, 0) = "ArmLength"
    'geometryValues(8, 0) = "FixtureArrangement"

'Extract the road geometry variables we need for this calculation
NumberOfLanes = geometryValues(1, intBaselineUpgradeChoice)
lanewidth = geometryValues(2, intBaselineUpgradeChoice)
MedianLength = geometryValues(3, intBaselineUpgradeChoice)
FixtureHeight = geometryValues(4, intBaselineUpgradeChoice)
polespacing = geometryValues(5, intBaselineUpgradeChoice)
poleconfig = geometryValues(8, intBaselineUpgradeChoice)

'outputX and outputY are ALL measurement grid point coordinates. These are not pairs, since it's a rectangle we can just list each direction once.
'X is along the road. Y is across the road.
Dim outputX(), outputY()
outputX = gridXY(0)
outputY = gridXY(1)

'------------------------
'Calculate grid start and end in the X direction (along road).
    'Different methods are needed because of the rules about how many fixtures are included in the calculation in the two methods
    'The match method is a sloppy way to do this - it depends on the outputX array to have a certain number of poles before the
    'grid starts, but doesn't save that data explicitly, just assumes it will be. Works, but would be nice to refactor eventually.
    'FLAG move this and the version in Phi up the call stack, and do it cleaner.
'------------------------
If calculationmethod = "IES" Then
    istart = WorksheetFunction.Match(polespacing, outputX, True)
    iend = WorksheetFunction.Match(2 * polespacing, outputX, True) - 1
ElseIf calculationmethod = "CIE" Then
'start at what fixture
    startfixture = Int(5 * FixtureHeight / polespacing)
    startfixture = startfixture + 1
    istart = WorksheetFunction.Match(polespacing * startfixture, outputX, True) + 1
    iend = WorksheetFunction.Match(polespacing * (startfixture + 1), outputX, True)
End If


'------------------------------------------------------------------------------------
'For each grid point, calculate the angle and load it into an array called gammaArray()
'------------------------------------------------------------------------------------
numberOfX = iend - istart
numberOfY = UBound(outputY)
Dim gammaArray()
ReDim gammaArray(istart To iend, numberOfY)
'for debug purposes - arrays to hold the intermediate calculated values, ReDim'd to the grid size. Printed to sheet
Dim xArray(), yArray(), xPrimeArray(), yPrimeArray(), hPrimeArray()
ReDim xArray(istart To iend, numberOfY): ReDim yArray(istart To iend, numberOfY): ReDim xPrimeArray(istart To iend, numberOfY): ReDim yPrimeArray(istart To iend, numberOfY): ReDim hPrimeArray(istart To iend, numberOfY)
'm = outputX(1) FLAG delete me if no calc errors - appears unused

For i = istart To iend          'each grid point in the x direction
    For j = 0 To numberOfY      'each grid point in the y direction
         'Variables in formulas written in standard vs. variable names used here
         'Just for comparing to PDF, v,w,o are not used in the VBA
        '    v = tiltOnZ
        '    w = tiltOnY
        '    o = tiltOnX
        Dim x As Double, y As Double
        
        'Get distances for calculating the angle
        x = (outputX(i) - fixtureX)
        y = (outputY(j) - fixtureY)
        
        'Depending which way the fixture faces, the sign of 'y' may need to be reversed because the fixture will be rotated 180 degrees.
        If backwardsFlag = False Then y = y
        If backwardsFlag = True Then y = -y

         'Calculate angle including tilt
        xPrime = x * (Cos(tiltOnZ) * Cos(tiltOnY) - Sin(tiltOnZ) * Sin(tiltOnX) * Sin(tiltOnY)) + _
                 y * (Sin(tiltOnZ) * Cos(tiltOnY) + Cos(tiltOnZ) * Sin(tiltOnX) * Sin(tiltOnY)) + _
                 FixtureHeight * Cos(tiltOnX) * Sin(tiltOnY)
        yPrime = -x * Sin(tiltOnZ) * Cos(tiltOnX) + _
                    y * Cos(tiltOnZ) * Cos(tiltOnX) - _
                    FixtureHeight * Sin(tiltOnX)
        HPrime = -x * (Sin(tiltOnZ) * Sin(tiltOnX) * Cos(tiltOnY) + Cos(tiltOnZ) * Sin(tiltOnY)) - _
                 y * (Sin(tiltOnZ) * Sin(tiltOnY) - Cos(tiltOnZ) * Sin(tiltOnX) * Cos(tiltOnY)) + _
                 FixtureHeight * Cos(tiltOnX) * Cos(tiltOnY)
                    
        'Basic gamma calc yields angle between -90 and 90; needs to be converted to be between 0 and 180
        If HPrime = 0 Then
            gammaTemp = 90
        Else
            gammaTemp = Atn(((xPrime ^ 2 + yPrime ^ 2) ^ 0.5) / HPrime) * 180 / WorksheetFunction.Pi
        End If
        If HPrime < 0 Then gammaTemp = gammaTemp + 180
        gammaArray(i, j) = gammaTemp
        
        'saving the intermediate variables to array for debug purposes only - can export later if needed.
        xArray(i, j) = x
        yArray(i, j) = y
        xPrimeArray(i, j) = xPrime
        yPrimeArray(i, j) = yPrime
        hPrimeArray(i, j) = HPrime
    Next
Next

angleGammaWithTilt = gammaArray

'Temp code for debugging purposes
'If fixtureX = 35 Then 'Only worry about writing the two opposing fixtures
'    Dim rrow As Integer
'    rrow = 1
'    Dim aOutputGamma()
'    ReDim aOutput(300, 300) As Variant
'
'    Call printIntermediateVariables(rrow, aOutputGamma, xArray)
'    Call printIntermediateVariables(rrow, aOutputGamma, yArray)
'    Call printIntermediateVariables(rrow, aOutputGamma, xPrimeArray)
'    Call printIntermediateVariables(rrow, aOutputGamma, yPrimeArray)
'    Call printIntermediateVariables(rrow, aOutputGamma, hPrimeArray)
'
'    If fixtureY = 0.5 Then yColumn = 22
'    If fixtureY = 16.5 Then yColumn = 32
'    Set rTarget = wksScratch.Cells(100, yColumn)
'    rTarget.Resize(UBound(aOutput, 1), UBound(aOutput, 2)) = aOutput
'End If
End Function

Function Distance(x1, y1, x2, y2)
Distance = Sqr((x1 - x2) * (x1 - x2) + (y1 - y2) * (y1 - y2))
End Function


Function anglePhiWithTilt(fixtureX, fixtureY, backwardsFlag As Boolean, tiltOnX, tiltOnY, tiltOnZ, gridXY(), calculationmethod, intBaselineUpgradeChoice, geometryValues()) As Variant
'This function calculates the angle phi between the measurement grid point and the fixture itself.
'With the tilt, this angle phi can be used to calculate the light coming out of the fixture, i.e. the data in the IES files.
'If tilt is set to 0,0,0, the angle matches the actual angle between the fixture and grid for use in the Luminance calculation.
'The results of the function are assigned to an array that has the form phi(xStart to xEnd,yStart to yEnd)

'Variables required:
'fixtureX is the x coordinate of the fixture currently being calculated
'fixtureY is the y coordinate of the same
'gridXY() is an array containing the x/y coordinates of all grid points from the start to end of the road. istart and iend are used to select only the actual calculated grid bounds to be used here.
'tiltOnX, Y, and Z are the tilt in radians of the fixture, for the current calculation
'calculationmethod is either CIE or IES (string)
'intBaselineUpgradeChoice is used for choosing which road geometry data to use. 1 = Baseline, 2 = Upgrade
'geometryValues() is an array containing the road geometry data. Column 0 is header names (for reference), column 1 is baseline, 2 is upgrade. Rows are:
    'geometryValues(1, 0) = "NumLanes"
    'geometryValues(2, 0) = "LaneWidth"
    'geometryValues(3, 0) = "MedianWidth"
    'geometryValues(4, 0) = "MountingHeight"
    'geometryValues(5, 0) = "PoleSpacing"
    'geometryValues(6, 0) = "PoleSetback"
    'geometryValues(7, 0) = "ArmLength"
    'geometryValues(8, 0) = "FixtureArrangement"


'Extract the road geometry variables we need for this calculation
NumberOfLanes = geometryValues(1, intBaselineUpgradeChoice)
lanewidth = geometryValues(2, intBaselineUpgradeChoice)
MedianLength = geometryValues(3, intBaselineUpgradeChoice)
FixtureHeight = geometryValues(4, intBaselineUpgradeChoice)
polespacing = geometryValues(5, intBaselineUpgradeChoice)
poleconfig = geometryValues(8, intBaselineUpgradeChoice)

'outputX and outputY are ALL measurement grid point coordinates. These are not pairs, since it's a rectangle we can just list each direction once.
Dim outputX(), outputY()
outputX = gridXY(0)
outputY = gridXY(1)

'------------------------------------
'Calculate grid start and end in the X direction (along road).
        'Different methods are needed because of the rules about how many fixtures are included in the calculation in the two methods
        'The match method is a sloppy way to do this - it depends on the outputX array to have a certain number of poles before the
        'grid starts, but doesn't save that data explicitly, just assumes it will be. Works, but would be nice to refactor eventually.
'------------------------------------
If calculationmethod = "IES" Then
    'In the IES method, the measurement grid starts one polespacing length from the first pole.
    'istart is the lower bound to use for grid points in the X direction; iend is the upper bound.
    istart = WorksheetFunction.Match(polespacing, outputX, True)                'Match syntax: lookup value, lookup array, exact match. "True" = 1 = "less than"
    iend = WorksheetFunction.Match(2 * polespacing, outputX, True) - 1
ElseIf calculationmethod = "CIE" Then
    'start at what fixture
    startfixture = Int(5 * FixtureHeight / polespacing)
    startfixture = startfixture + 1
    istart = WorksheetFunction.Match(polespacing * startfixture, outputX, True) + 1
    iend = WorksheetFunction.Match(polespacing * (startfixture + 1), outputX, True)
    '**FLAG** when testing CIE make sure this is correct
End If



'------------------------------------------------------------------------------------
'For each grid point, calculate the angle and load it into an array called phiArray()
'------------------------------------------------------------------------------------
numberOfX = iend - istart
numberOfY = UBound(outputY)
Dim phiArray()
ReDim phiArray(istart To iend, numberOfY)
For i = istart To iend                  'each grid point in the x direction
    For j = 0 To numberOfY              'each grid point in the y direction
        'Variables in formulas written in standard vs. variable names used here
        'v,w,o are just how they are written in the pdf - not used anywhere here
        '    v = tiltOnZ
        '    w = tiltOnY
        '    o = tiltOnX
        Dim x As Double, y As Double
        Dim nearSide As Boolean
        
        'Get distances for calculating the angle
        x = (outputX(i) - fixtureX)
        y = (outputY(j) - fixtureY)
        
        'Depending which way the fixture faces, the sign of 'y' may need to be reversed because the fixture will be rotated 180 degrees.
        If backwardsFlag = False Then y = y
        If backwardsFlag = True Then y = -y
        
        'Calculate angle including tilt
        xPrime = x * (Cos(tiltOnZ) * Cos(tiltOnY) - Sin(tiltOnZ) * Sin(tiltOnX) * Sin(tiltOnY)) + _
                 y * (Sin(tiltOnZ) * Cos(tiltOnY) + Cos(tiltOnZ) * Sin(tiltOnX) * Sin(tiltOnY)) + _
                 FixtureHeight * Cos(tiltOnX) * Sin(tiltOnY)
        yPrime = -x * Sin(tiltOnZ) * Cos(tiltOnX) + _
                    y * Cos(tiltOnZ) * Cos(tiltOnX) - _
                    FixtureHeight * Sin(tiltOnX)
        HPrime = -x * (Sin(tiltOnZ) * Sin(tiltOnX) * Cos(tiltOnY) + Cos(tiltOnZ) * Sin(tiltOnY)) - _
                y * (Sin(tiltOnZ) * Sin(tiltOnY) - Cos(tiltOnZ) * Sin(tiltOnX) * Cos(tiltOnY)) + _
                FixtureHeight * Cos(tiltOnX) * Cos(tiltOnY)
        
        If yPrime <> 0 Then
            phiTemp = Atn(Abs(xPrime) / Abs(yPrime)) * 180 / WorksheetFunction.Pi
        Else
            phiTemp = 90
        End If
        
        'convert the angle if it is located behind the fixture
        If yPrime < 0 Then
            phiTemp = 180 - phiTemp
        Else
            phiTemp = phiTemp
        End If
        
        phiArray(i, j) = phiTemp
    Next j
Next i
'------------------------------------------------------------------------------------

anglePhiWithTilt = phiArray

End Function

Function angleBeta(anglePhi(), calculationmethod As String, fixtureX, fixtureY, gridXY(), polespacing, lanewidth, FixtureHeight, yo) As Variant
'yo is the y of the observer, placed in the center of each lane in turn.
Dim xo
Dim outputX()
Dim outputY()
outputX = gridXY(0)     'the location of the grid points
outputY = gridXY(1)     'the location of the grid points

'grid start and end
If calculationmethod = "IES" Then
    istart = WorksheetFunction.Match(polespacing, outputX, True)
    iend = WorksheetFunction.Match(2 * polespacing, outputX, True) - 1
ElseIf calculationmethod = "CIE" Then
    'start at what fixture
    startfixture = Int(5 * FixtureHeight / polespacing)
    startfixture = startfixture + 1
    istart = WorksheetFunction.Match(polespacing * startfixture, outputX, True) + 1
    iend = WorksheetFunction.Match(polespacing * (startfixture + 1), outputX, True)
    'iStart = WorksheetFunction.Match(5 * FixtureHeight, outputX, True)
    'iEnd = WorksheetFunction.Match(5 * FixtureHeight + polespacing, outputX, True)
End If

numberOfX = iend - istart
numberOfY = UBound(outputY)

Dim betaArray()
ReDim betaArray(istart To iend, numberOfY)
'This is for IES method
If calculationmethod = "IES" Then
For i = istart To iend
For j = 0 To numberOfY
    If fixtureX - outputX(i) > 0 Then
        If anglePhi(i, j) >= 90 Then
            betaArray(i, j) = anglePhi(i, j) - 90                                   'agree
        Else
            betaArray(i, j) = 90 - anglePhi(i, j) '90 degrees=1.57079633radians     'agree
        End If
    Else
        If anglePhi(i, j) >= 90 Then
            betaArray(i, j) = 270 - anglePhi(i, j)                                  'this is between 90 and 180 - is this how Beta was designed?
        Else
            betaArray(i, j) = 90 + anglePhi(i, j) '90 degrees=1.57079633radians     'agree
        End If
    End If
    Next
    Next
    '  if CIE method
ElseIf calculationmethod = "CIE" Then
    For i = istart To iend
    For j = 0 To numberOfY
        xo = outputX(i) - 60
        m1 = (outputY(j) - yo) / (outputX(i) - xo)
        If (outputX(i) - fixtureX) = 0 Then
            m2 = 10000000 ' a very high value in place of infinity
        Else
            m2 = (outputY(j) - fixtureY) / (outputX(i) - fixtureX)
        End If
        
        If m1 * m2 = -1 Then
            betaArray(i, j) = 90
        Else
            If m2 >= 10000000 Then
                betaArray(i, j) = 180
            Else
                If fixtureX - outputX(i) > 0 Then
                    betaArray(i, j) = Atn(Abs(m1 - m2 / 1 + m1 * m2)) * (180 / WorksheetFunction.Pi)
                Else
                    betaArray(i, j) = 180 - Atn(Abs(m1 - m2 / 1 + m1 * m2)) * (180 / WorksheetFunction.Pi)
                End If
            End If
        End If
Next
Next
Else
End If
angleBeta = betaArray
End Function
Sub drawtest()

' to plot points of measurement grid. Not being used anywhere in final version.
'This was used to check angle calcualtion of gamma, phi and beta at every grid point.

Dim ngp As Integer
Dim lanewidth As Integer, MedianLength As Integer, FixtureHeight As Integer, NumberOfLanes As Integer, polespacing As Integer
lanewidth = Sheets("Road Geometry").Range("bLaneWidth").Value
MedianLength = Sheets("Road Geometry").Range("bMedianWidth").Value
FixtureHeight = Sheets("Road Geometry").Range("bMountingHeight").Value
NumberOfLanes = Sheets("Road Geometry").Range("bNumLanes").Value
polespacing = Sheets("Road Geometry").Range("bPoleSpacing").Value

Dim polesetback As Integer, ArmLength As Integer
polesetback = Sheets("Road Geometry").Range("bPoleSetback").Value
ArmLength = Sheets("Road Geometry").Range("bArmLength").Value

Dim calcMethod As String, poleconfig As String, dataSheet As String, chartName As String
calcMethod = Sheets("FixtureData").Range("iescieGraphChoice").Value
poleconfig = Sheets("Road Geometry").Range("bFixtureArrangement").Value
dataSheet = "Chart Data Baseline"
chartName = "Baseline"

ngp = TotalGridLength(calcMethod, FixtureHeight, polespacing) / GridSpace(calcMethod, polespacing)
Dim outputXY()
Dim gammaArray()
Dim phi()

outputXY = makeGrid(NumberOfLanes, calcMethod, ngp, poleconfig, MedianLength, polespacing, lanewidth)
'outputY = makeGrid(NumberOfLanes, CalculationMethod, ngp, poleConfig, MedianLength, polespacing, lanewidth)(1)
gammaArray = anglePhi(75, -1, outputXY, polespacing, FixtureHeight, calcMethod)
'gammaArray = angleBeta(phi(), calcMethod, 0, 15, outputXY, polespacing, lanewidth, FixtureHeight)
'gammaArray = angleGamma(75, -1, outputXY, polespacing, FixtureHeight, calcMethod)
Sheets("test1").Activate
For i = LBound(gammaArray(), 1) To UBound(gammaArray(), 1)
For j = 0 To UBound(gammaArray(), 2)
Cells(i + 101, j + 1) = gammaArray(i, j)
Next
Next

End Sub








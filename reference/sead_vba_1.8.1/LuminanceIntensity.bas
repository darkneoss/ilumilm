Attribute VB_Name = "LuminanceIntensity"
'This macro was developed by p2w2.  http://p2w2.com/expert-in-microsoft-excel-consultants-consulting/index.php
'Please contact at CS@perceptive-analytics.com in case of any enquiry

' Quadratic interpolation to get luminous intensity
Function LintensityMatrix(gridlength, poleconfig As String, fixtureX, fixtureY, gridXY(), polespacing, FixtureHeight, calculationmethod As String, gridPhi, gridGamma)

Dim outputX()
Dim outputY()
outputX = gridXY(0)
outputY = gridXY(1)

'FLAG this does not appear to be used anywhere; commenting out, delete if no errors
'Dim numberoffixtures As Integer
'If PoleConfiguration = "Single-side" Then
'    numberoffixtures = (gridlength / polespacing)
'Else
'    numberoffixtures = (gridlength / polespacing) * 2
'End If


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

Dim numberOfX, numberOfY As Integer

numberOfX = iend - istart
numberOfY = UBound(outputY)
Dim LIarray()
ReDim LIarray(istart To iend, numberOfY)
m = outputX(1)

Dim tablePhi2()
Dim tableGamma2()
Dim tableArray()

Dim firstRow, lastRow As Integer
Dim lastCol As Integer
'Locate the relevant fixture light output on the FixtureData tab
firstRow = Sheets("FixtureData").Range("B16").Value + 9             'Row with angles (vert angles)
lastRow = Sheets("FixtureData").Range("B17").Value
lastCol = Sheets("FixtureData").Cells(firstRow, Sheets("FixtureData").Columns.count).End(xlToLeft).column
Sheets("FixtureData").Select

'FLAG this should be moved up one call stack level, so that it is done only once per fixture (not for every fixture instance in the range)
'Transfer the angle data to appropriate cells. table...2 is dimensioned using spreadsheet dimensions, and then values are transferred to table..1 which are 1-d
tablePhi2() = Range(Cells(firstRow + 1, 1), Cells(lastRow, 1))        'values in Column A
tableGamma2() = Range(Cells(firstRow, 2), Cells(firstRow, lastCol))       'Values in first row of intensity table
Dim gammaN, phiN As Integer
gammaN = UBound(tableGamma2, 2)
phiN = UBound(tablePhi2, 1)

Dim tablePhi1()
Dim tableGamma1()
ReDim tableGamma1(gammaN)
ReDim tablePhi1(phiN)
For m = LBound(tableGamma2(), 2) To UBound(tableGamma2(), 2)
    tableGamma1(m) = tableGamma2(1, m)
Next

For n = LBound(tablePhi2(), 1) To UBound(tablePhi2(), 1)
    tablePhi1(n) = tablePhi2(n, 1)
Next

tableArray() = Range(Cells(firstRow + 1, 2), Cells(lastRow, lastCol))


For i = istart To iend
For j = 0 To numberOfY
    LIarray(i, j) = LintensityCalc(gridPhi(i, j), gridGamma(i, j), tablePhi1(), tableGamma1(), tableArray())
Next
Next

LintensityMatrix = LIarray

End Function

Function LintensityCalc(gridPhi, gridGamma, tablePhi(), tableGamma(), tableArray()) As Variant
' gives I(phi,gamma) at single grindpoint from single fixture
'gridPhi - the angle of Phi for the specific grid point being calculated
'gridGamma - same for gamma
'tableGamma - the table of all gamma values in the fixture data tab
'tablePhi - the table o
'************* Getting the Phi values and calculating K1, K2, K3 ******************
Dim Phi0, Phi1, Phi2
Dim Phi0pos, Phi1pos, Phi2pos

With Application.WorksheetFunction
'find the closest Phi from the table to the gridPhi - greater than gridPhi

    Phi1pos = Application.Match(gridPhi, tablePhi, 1)
    Phi1 = tablePhi(Phi1pos)

    'find the next closest Phi - less than gridPhi
    If Phi1pos = 1 Then
        Phi0pos = Phi1pos + 1
    ElseIf Phi1pos = UBound(tablePhi()) Then
        Phi0pos = Phi1pos - 1
    Else
        Phi0pos = Phi1pos - 1
    End If
    Phi0 = tablePhi(Phi0pos)
    
    'calculating the constants K1, K2, K3 based on Phi values
    'Phi0=C(m+1), Phi1=C(m+2), Phi2=Cm

Dim K1, K2

    If Phi1 = Phi0 Then
        K1 = 0
    Else
        K1 = (gridPhi - Phi1) / (Phi0 - Phi1)
    End If
    K2 = 1 - K1

End With

'************* Getting the Gamma values and calculating k1, k2, k3 ******************
Dim Gamma0, Gamma1, Gamma2
Dim Gamma0pos As Integer, Gamma1pos As Integer, Gamma2pos As Integer

With Application.WorksheetFunction
'find the closest Gamma from the table to the gridGamma - greater than gridGamma
    
    Gamma1pos = Application.Match(gridGamma, tableGamma, 1)
    Gamma1 = tableGamma(Gamma1pos)
    
    'find the next closest Gamma- less than gridGamma
    If Gamma1pos = 1 Then
        Gamma0pos = Gamma1pos + 1
    ElseIf Gamma1pos = UBound(tableGamma()) Then
        Gamma0pos = Gamma1pos - 1
    Else
        Gamma0pos = Gamma1pos - 1
    End If
        Gamma0 = tableGamma(Gamma0pos)

'calculating the constants k1, k2, k3 based on Gamma values
'Gamma0=Gamma(j+1), Gamma1=Gamma(j), Gamma2=Gamma(j+2)
Dim kGamma1, kGamma2, kGamma3
    If Gamma1 = Gamma0 Then
        kGamma1 = 0
    Else
        kGamma1 = (gridGamma - Gamma1) / (Gamma0 - Gamma1)
    End If
    kGamma2 = 1 - (kGamma1)

End With

'********************calculation of Luminance intensity **************************
'step 1: calculation of Luminance Intensity for gridPhi at three Gamma values - Gamma0, Gamma1, Gamma2

'IgridPhiGamma0=I(C,Gamma(j+1)),IgridPhiGamma1=I(C,Gamma(j), IgridPhiGamma2=I(C,Gamma(j+2)
'IPhi0Gamma0=I(C(m+1)Gamma(j+1)), IPhi0Gamma1=I(C(m+1)Gamma(j)), IPhi0Gamma2=I(C(m+1)Gamma(j+2))
'IPhi1Gamma0=I(C(m+2)Gamma(j+1)), IPhi1Gamma1=I(C(m+2)Gamma(j)), IPhi1Gamma2=I(C(m+2)Gamma(j+2))
'IPhi2Gamma0=I(C(m)Gamma(j+1)), IPhi2Gamma1=I(C(m)Gamma(j)), IPhi2Gamma2=I(C(m)Gamma(j+2))

Dim IgridPhiGamma0, IgridPhiGamma1
Dim IPhi0Gamma0, IPhi0Gamma1
Dim IPhi1Gamma0, IPhi1Gamma1

IPhi0Gamma0 = tableArray(Phi0pos, Gamma0pos)
IPhi0Gamma1 = tableArray(Phi0pos, Gamma1pos)
IPhi1Gamma0 = tableArray(Phi1pos, Gamma0pos)
IPhi1Gamma1 = tableArray(Phi1pos, Gamma1pos)

IgridPhiGamma0 = K1 * IPhi0Gamma0 + K2 * IPhi1Gamma0
IgridPhiGamma1 = K1 * IPhi0Gamma1 + K2 * IPhi1Gamma1


'step 2: calculation of final value of Luminance Intensity
LintensityCalc = kGamma1 * IgridPhiGamma0 + kGamma2 * IgridPhiGamma1

End Function




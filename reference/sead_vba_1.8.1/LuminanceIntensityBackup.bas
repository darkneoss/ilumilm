Attribute VB_Name = "LuminanceIntensityBackup"
'This macro was developed by p2w2.  http://p2w2.com/expert-in-microsoft-excel-consultants-consulting/index.php
'Please contact at CS@perceptive-analytics.com in case of any enquiry

' Quadratic interpolation to get luminous intensity
Function LintensityMatrix1(gridlength, poleconfig As String, fixtureX, fixtureY, gridXY(), polespacing, FixtureHeight, calculationmethod As String, gridPhi, gridGamma)

Dim outputX()
Dim outputY()
outputX = gridXY(0)
outputY = gridXY(1)

Dim numberoffixtures As Integer


If PoleConfiguration = "Single-side" Then
numberoffixtures = (gridlength / polespacing)
Else
numberoffixtures = (gridlength / polespacing) * 2
End If


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
firstRow = Sheets("FixtureData").Range("B16").Value + 9
lastRow = Sheets("FixtureData").Range("B17").Value
lastCol = Sheets("FixtureData").Cells(firstRow, Sheets("FixtureData").Columns.count).End(xlToLeft).column
Sheets("FixtureData").Select
tableGamma2() = Range(Cells(firstRow + 1, 1), Cells(lastRow, 1))
tablePhi2() = Range(Cells(firstRow, 2), Cells(firstRow, lastCol))
Dim gammaN, phiN As Integer
phiN = UBound(tablePhi2, 2)
gammaN = UBound(tableGamma2, 1)

Dim tablePhi1()
Dim tableGamma1()
ReDim tablePhi1(phiN)
ReDim tableGamma1(gammaN)
For m = LBound(tablePhi2(), 2) To UBound(tablePhi2(), 2)
    tablePhi1(m) = tablePhi2(1, m)
Next

For n = LBound(tableGamma2(), 1) To UBound(tableGamma2(), 1)
    tableGamma1(n) = tableGamma2(n, 1)
Next

tableArray() = Range(Cells(firstRow + 1, 2), Cells(lastRow, lastCol))


For i = istart To iend
For j = 0 To numberOfY
' tablephi1 and tablegamma1 got swapped. Instead of changing all the above logic, just changed the arguments being passed
LIarray(i, j) = LintensityCalc(gridPhi(i, j), gridGamma(i, j), tableGamma1(), tablePhi1(), tableArray())

Next
Next

LintensityMatrix = LIarray

End Function

Function LintensityCalc1(gridPhi, gridGamma, tablePhi(), tableGamma(), tableArray()) As Variant
' gives I(phi,gamma) at single grindpoint from single fixture

'************* Getting the Phi values and calculating K1, K2, K3 ******************

Dim Phi0, Phi1, Phi2
Dim Phi0pos As Integer, Phi1pos As Integer, Phi2pos As Integer


With Application.WorksheetFunction
'find the closest Phi from the table to the gridPhi - greater than gridPhi

Phi0pos = Application.Match(gridPhi, tablePhi, 1)
Phi0 = tablePhi(Phi0pos)



'find the next closest Phi - less than gridPhi

If Phi0pos = 1 Then
    Phi1pos = Phi0pos + 1
ElseIf Phi0pos = UBound(tablePhi()) Then
    Phi1pos = Phi0pos - 1
Else
    Phi1pos = Phi0pos - 1
End If
Phi1 = tablePhi(Phi1pos)


'find the third point

If gridPhi <= (Phi0 + Phi1) / 2 Then
    If Phi1pos = 1 Then
        Phi2pos = Phi0pos + 1
    ElseIf Phi0pos = UBound(tablePhi) Then
        Phi2pos = Phi1pos - 1
    Else
        Phi2pos = Phi1pos - 1
    End If
Else
    If Phi1pos = 1 Then
        Phi2pos = Phi0pos + 1
    ElseIf Phi0pos = UBound(tablePhi) Then
        Phi2pos = Phi1pos - 1
    Else
        Phi2pos = Phi0pos + 1
    End If
End If

Phi2 = tablePhi(Phi2pos)

'rearranging phi's in ascending order
Dim tempphi
tempphi = Phi0pos
Phi0pos = Phi1pos
Phi1pos = tempphi

If Phi2pos > Phi1pos Then
Else
tempphi = Phi0pos
Phi0pos = Phi2pos
Phi2pos = Phi1pos
Phi1pos = tempphi
End If

Phi0 = tablePhi(Phi0pos)
Phi1 = tablePhi(Phi1pos)
Phi2 = tablePhi(Phi2pos)


'calculating the constants K1, K2, K3 based on Phi values
'Phi0=C(m+1), Phi1=C(m+2), Phi2=Cm

Dim K1, K2, K3

If Phi0 = Phi2 Or Phi1 = Phi0 Then
    K1 = 0
Else
    K1 = ((gridPhi - Phi1) * (gridPhi - Phi2)) / ((Phi0 - Phi1) * (Phi0 - Phi2))
End If

If Phi1 = Phi2 Or Phi0 = Phi1 Then
    K2 = 0
Else
    K2 = ((gridPhi - Phi0) * (gridPhi - Phi2)) / ((Phi1 - Phi0) * (Phi1 - Phi2))
End If

K3 = 1 - (K1 + K2)

End With

'************* Getting the Gamma values and calculating k1, k2, k3 ******************

Dim Gamma0, Gamma1, Gamma2
Dim Gamma0pos As Integer, Gamma1pos As Integer, Gamma2pos As Integer



With Application.WorksheetFunction
'find the closest Gamma from the table to the gridGamma - greater than gridGamma


Gamma0pos = Application.Match(gridGamma, tableGamma, 1)
Gamma0 = tableGamma(Gamma0pos)



'find the next closest Gamma- less than gridGamma

If Gamma0pos = 1 Then
    Gamma1pos = Gamma0pos + 1
ElseIf Gamma0pos = UBound(tableGamma()) Then
    Gamma1pos = Gamma0pos - 1
Else
    Gamma1pos = Gamma0pos - 1
End If
Gamma1 = tableGamma(Gamma1pos)



'find the third point

If gridGamma <= (Gamma0 + Gamma1) / 2 Then
    If Gamma1pos = 1 Then
        Gamma2pos = Gamma0pos + 1
    ElseIf Gamma1pos = UBound(tableGamma) Then
        Gamma2pos = Gamma0pos - 1
    Else
        Gamma2pos = Gamma1pos - 1
    End If
Else
    If Gamma1pos = 1 Then
        Gamma2pos = Gamma0pos + 1
    ElseIf Gamma1pos = UBound(tableGamma) Then
        Gamma2pos = Gamma0pos - 1
    Else
        Gamma2pos = Gamma0pos + 1
    End If
End If

Gamma2 = tableGamma(Gamma2pos)


'rearranging gamma
Dim tempgamma
tempgamma = Gamma0pos
Gamma0pos = Gamma1pos
Gamma1pos = tempgamma

If Gamma2pos > Gamma1pos Then
Else
tempgamma = Gamma0pos
Gamma0pos = Gamma2pos
Gamma2pos = Gamma1pos
Gamma1pos = tempgamma
End If

Gamma0 = tableGamma(Gamma0pos)
Gamma1 = tableGamma(Gamma1pos)
Gamma2 = tableGamma(Gamma2pos)


'calculating the constants k1, k2, k3 based on Gamma values
'Gamma0=Gamma(j+1), Gamma1=Gamma(j), Gamma2=Gamma(j+2)

Dim kGamma1, kGamma2, kGamma3

If Gamma1 = Gamma0 Or Gamma0 = Gamma2 Then
    kGamma1 = 0
Else
    kGamma1 = ((gridGamma - Gamma1) * (gridGamma - Gamma2)) / ((Gamma0 - Gamma1) * (Gamma0 - Gamma2))
End If

If Gamma0 = Gamma1 Or Gamma1 = Gamma2 Then
    kGamma2 = 0
Else
    kGamma2 = ((gridGamma - Gamma0) * (gridGamma - Gamma2)) / ((Gamma1 - Gamma0) * (Gamma1 - Gamma2))
End If

kGamma3 = 1 - (kGamma1 + kGamma2)

End With

'********************calculation of Luminance intensity **************************

'step 1: calculation of Luminance Intensity for gridPhi at three Gamma values - Gamma0, Gamma1, Gamma2

'IgridPhiGamma0=I(C,Gamma(j+1)),IgridPhiGamma1=I(C,Gamma(j), IgridPhiGamma2=I(C,Gamma(j+2)
'IPhi0Gamma0=I(C(m+1)Gamma(j+1)), IPhi0Gamma1=I(C(m+1)Gamma(j)), IPhi0Gamma2=I(C(m+1)Gamma(j+2))
'IPhi1Gamma0=I(C(m+2)Gamma(j+1)), IPhi1Gamma1=I(C(m+2)Gamma(j)), IPhi1Gamma2=I(C(m+2)Gamma(j+2))
'IPhi2Gamma0=I(C(m)Gamma(j+1)), IPhi2Gamma1=I(C(m)Gamma(j)), IPhi2Gamma2=I(C(m)Gamma(j+2))

Dim IgridPhiGamma0, IgridPhiGamma1, IgridPhiGamma2
Dim IPhi0Gamma0, IPhi0Gamma1, IPhi0Gamma2
Dim IPhi1Gamma0, IPhi1Gamma1, IPhi1Gamma2
Dim IPhi2Gamma0, IPhi2Gamma1, IPhi2Gamma2

IPhi0Gamma0 = tableArray(Phi0pos, Gamma0pos)
IPhi0Gamma1 = tableArray(Phi0pos, Gamma1pos)
IPhi0Gamma2 = tableArray(Phi0pos, Gamma2pos)
IPhi1Gamma0 = tableArray(Phi1pos, Gamma0pos)
IPhi1Gamma1 = tableArray(Phi1pos, Gamma1pos)
IPhi1Gamma2 = tableArray(Phi1pos, Gamma2pos)
IPhi2Gamma0 = tableArray(Phi2pos, Gamma0pos)
IPhi2Gamma1 = tableArray(Phi2pos, Gamma1pos)
IPhi2Gamma2 = tableArray(Phi2pos, Gamma2pos)

IgridPhiGamma0 = K1 * IPhi0Gamma0 + K2 * IPhi1Gamma0 + K3 * IPhi2Gamma0
IgridPhiGamma1 = K1 * IPhi0Gamma1 + K2 * IPhi1Gamma1 + K3 * IPhi2Gamma1
IgridPhiGamma2 = K1 * IPhi0Gamma2 + K2 * IPhi1Gamma2 + K3 * IPhi2Gamma2

'step 2: calculation of final value of Luminance Intensity


LintensityCalc = kGamma1 * IgridPhiGamma0 + kGamma2 * IgridPhiGamma1 + kGamma3 * IgridPhiGamma2



End Function




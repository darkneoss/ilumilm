Attribute VB_Name = "Rtablecalcs"
'This macro was developed by p2w2.  http://p2w2.com/expert-in-microsoft-excel-consultants-consulting/index.php
'Please contact at CS@perceptive-analytics.com in case of any enquiry

'Quadratic interpolation for r table
Function RMatrix(gridlength, poleconfig As String, fixtureX, fixtureY, gridXY(), polespacing, FixtureHeight, calculationmethod As String, gridBeta, gridGamma) As Variant

Dim outputX()
Dim outputY()
outputX = gridXY(0)
outputY = gridXY(1)

Dim numberoffixtures As Integer

If poleconfig = "Opposite" Or poleconfig = "Median mounted" Then
    numberoffixtures = (gridlength / polespacing) * 2
ElseIf poleconfig = "Single" Or poleconfig = "Staggered" Then
    numberoffixtures = (gridlength / polespacing)
End If


'grid start and end
If calculationmethod = "IES" Then
iStart = WorksheetFunction.Match(polespacing, outputX, True)
iEnd = WorksheetFunction.Match(2 * polespacing, outputX, True) - 1
ElseIf calculationmethod = "CIE" Then
'start at what fixture
startfixture = Int(5 * FixtureHeight / polespacing)
startfixture = startfixture + 1
iStart = WorksheetFunction.Match(polespacing * startfixture, outputX, True) + 1
iEnd = WorksheetFunction.Match(polespacing * (startfixture + 1), outputX, True)
'iStart = WorksheetFunction.Match(5 * FixtureHeight, outputX, True)
'iEnd = WorksheetFunction.Match(5 * FixtureHeight + polespacing, outputX, True)
End If

Dim numberOfX, numberOfY As Integer

numberOfX = iEnd - iStart
numberOfY = UBound(outputY)
Dim Rarray()
ReDim Rarray(iStart To iEnd, numberOfY)
m = outputX(1)

Dim tableBeta2()
Dim tableTanGamma2()
Dim tableRArray()

Sheets("Rtables").Select
tableBeta2() = Range("D5:W5")
tableTanGamma2() = Range("C6:C34")



Dim gammaN, betaN As Integer
betaN = UBound(tableBeta2, 2)
gammaN = UBound(tableTanGamma2, 1)

Dim tableBeta1()
Dim tableTanGamma1()
ReDim tableBeta1(1 To betaN)
ReDim tableTanGamma1(1 To gammaN)

For m = LBound(tableBeta2(), 2) To UBound(tableBeta2(), 2)
    tableBeta1(m) = tableBeta2(1, m)
Next

For n = LBound(tableTanGamma2(), 1) To UBound(tableTanGamma2(), 1)
    tableTanGamma1(n) = tableTanGamma2(n, 1)
Next

tableRArray() = Range("D6:W34")


For i = iStart To iEnd
For j = 0 To numberOfY

Rarray(i, j) = RCalc(gridBeta(i, j), gridGamma(i, j), tableBeta1(), tableTanGamma1(), tableRArray())

Next
Next

RMatrix = Rarray

End Function

Function RCalc(gridBeta, gridGamma, tableBeta(), tableTanGamma(), tableRArray()) As Variant
' gives R(beta,gamma) at single grindpoint from single fixture

'************* Getting the Beta values and calculating K1, K2, K3 ******************

Dim Beta0, Beta1, Beta2
Dim Beta0pos, Beta1pos, Beta2pos As Integer
Dim K1, K2, K3
Dim kGamma1, kGamma2, kGamma3

With Application.WorksheetFunction
'find the closest Beta from the table to the gridBeta - greater than gridBeta

Beta1pos = Application.Match(gridBeta, tableBeta, 1)
If Abs(gridBeta - tableBeta(Beta1pos)) < 0.0000001 Then
Beta0pos = Beta1pos
Beta2pos = Beta1pos
GoTo 111:
End If
Beta1pos = Beta1pos + 1
If gridBeta = 180 Then
Beta1pos = Beta1pos - 1
End If
Beta1 = tableBeta(Beta1pos)



'find the next closest Beta - less than gridBeta


Beta0pos = Beta1pos - 1



'find the third point
If Beta1pos >= UBound(tableBeta()) Then

Beta0 = tableBeta(Beta0pos)
Beta1 = tableBeta(Beta1pos)
Beta2pos = Beta1pos
Beta2 = tableBeta(Beta2pos)
If Beta1 = Beta0 Then
    K1 = 0
Else
    K1 = (gridBeta - Beta1) / (Beta0 - Beta1)
End If

K2 = 1 - (K1)

Else

Beta2pos = Beta1pos + 1

111:

Beta0 = tableBeta(Beta0pos)
Beta1 = tableBeta(Beta1pos)
Beta2 = tableBeta(Beta2pos)



'calculating the constants K1, K2, K3 based on Beta values
'Beta0=C(m+1), Beta1=C(m+2), Beta2=Cm



If Beta0 = Beta2 Or Beta1 = Beta0 Then
    K1 = 0
Else
    K1 = ((gridBeta - Beta1) * (gridBeta - Beta2)) / ((Beta0 - Beta1) * (Beta0 - Beta2))
End If

If Beta1 = Beta2 Or Beta0 = Beta1 Then
    K2 = 0
Else
    K2 = ((gridBeta - Beta0) * (gridBeta - Beta2)) / ((Beta1 - Beta0) * (Beta1 - Beta2))
End If

K3 = 1 - (K1 + K2)

End If

End With

'************* Getting the Gamma values and calculating k1, k2, k3 ******************

Dim Gamma0, Gamma1, Gamma2
Dim tanGamma0pos, tanGamma1pos, tanGamma2pos As Integer

Dim gridTanGamma
gridTanGamma = Tan(gridGamma * WorksheetFunction.Pi / 180)

With Application.WorksheetFunction
'find the closest Gamma from the table to the gridGamma - greater than gridGamma

'x = UBound(tableTanGamma)
'y = LBound(tableTanGamma)
tanGamma1pos = Application.Match(gridTanGamma, tableTanGamma, 1)
If Abs(gridTanGamma - tableTanGamma(tanGamma1pos)) < 0.0000001 Then
tanGamma0pos = tanGamma1pos
tanGamma2pos = tanGamma1pos
GoTo 222:
End If
tanGamma1pos = tanGamma1pos + 1
If gridTanGamma >= 12 Then
tanGamma1pos = tanGamma1pos - 1
End If
Gamma1 = tableTanGamma(tanGamma1pos)
'find the next closest Gamma- less than gridGamma


tanGamma0pos = tanGamma1pos - 1


'find the third point
If tanGamma1pos = UBound(tableTanGamma) Then

Gamma0 = tableTanGamma(tanGamma0pos)
Gamma1 = tableTanGamma(tanGamma1pos)
tanGamma2pos = tanGamma1pos
Gamma2 = tableTanGamma(tanGamma2pos)
If Gamma0 = Gamma1 Then
    kGamma1 = 0
Else
    kGamma1 = (gridTanGamma - Gamma1) / (Gamma0 - Gamma1)
End If
kGamma2 = 1 - (kGamma1)

Else

tanGamma2pos = tanGamma1pos + 1

222:

Gamma0 = tableTanGamma(tanGamma0pos)
Gamma1 = tableTanGamma(tanGamma1pos)
Gamma2 = tableTanGamma(tanGamma2pos)


'calculating the constants k1, k2, k3 based on Gamma values
'Gamma0=Gamma(j+1), Gamma1=Gamma(j), Gamma2=Gamma(j+2)



If Gamma0 = Gamma1 Or Gamma0 = Gamma2 Then
    kGamma1 = 0
Else
    kGamma1 = ((gridTanGamma - Gamma1) * (gridTanGamma - Gamma2)) / ((Gamma0 - Gamma1) * (Gamma0 - Gamma2))
End If

If Gamma0 = Gamma1 Or Gamma1 = Gamma2 Then
    kGamma2 = 0
Else
    kGamma2 = ((gridTanGamma - Gamma0) * (gridTanGamma - Gamma2)) / ((Gamma1 - Gamma0) * (Gamma1 - Gamma2))
End If

kGamma3 = 1 - (kGamma1 + kGamma2)

End If

End With

'********************calculation of Luminance intensity **************************

'step 1: calculation of Luminance Intensity for gridBeta at three Gamma values - Gamma0, Gamma1, Gamma2

'IgridBetaGamma0=I(C,Gamma(j+1)),IgridBetaGamma1=I(C,Gamma(j), IgridBetaGamma2=I(C,Gamma(j+2)
'IBeta0Gamma0=I(C(m+1)Gamma(j+1)), IBeta0Gamma1=I(C(m+1)Gamma(j)), IBeta0Gamma2=I(C(m+1)Gamma(j+2))
'IBeta1Gamma0=I(C(m+2)Gamma(j+1)), IBeta1Gamma1=I(C(m+2)Gamma(j)), IBeta1Gamma2=I(C(m+2)Gamma(j+2))
'IBeta2Gamma0=I(C(m)Gamma(j+1)), IBeta2Gamma1=I(C(m)Gamma(j)), IBeta2Gamma2=I(C(m)Gamma(j+2))

Dim IgridBetaGamma0, IgridBetaGamma1, IgridBetaGamma2
Dim IBeta0Gamma0, IBeta0Gamma1, IBeta0Gamma2
Dim IBeta1Gamma0, IBeta1Gamma1, IBeta1Gamma2
Dim IBeta2Gamma0, IBeta2Gamma1, IBeta2Gamma2

If K3 = 0 And kGamma3 = 0 Then
'some code
IBeta0Gamma0 = tableRArray(tanGamma0pos, Beta0pos)
IBeta1Gamma0 = tableRArray(tanGamma0pos, Beta1pos)
IBeta0Gamma1 = tableRArray(tanGamma1pos, Beta0pos)
IBeta1Gamma1 = tableRArray(tanGamma1pos, Beta1pos)

IgridBetaGamma0 = K1 * IBeta0Gamma0 + K2 * IBeta1Gamma0
IgridBetaGamma1 = K1 * IBeta0Gamma1 + K2 * IBeta1Gamma1
RCalc = kGamma1 * IgridBetaGamma0 + kGamma2 * IgridBetaGamma1
 If RCalc < 0 Then
RCalc = 0
 End If
 
Else

IBeta0Gamma0 = tableRArray(tanGamma0pos, Beta0pos)
IBeta0Gamma1 = tableRArray(tanGamma1pos, Beta0pos)
IBeta0Gamma2 = tableRArray(tanGamma2pos, Beta0pos)
IBeta1Gamma0 = tableRArray(tanGamma0pos, Beta1pos)
IBeta1Gamma1 = tableRArray(tanGamma1pos, Beta1pos)
IBeta1Gamma2 = tableRArray(tanGamma2pos, Beta1pos)
IBeta2Gamma0 = tableRArray(tanGamma0pos, Beta2pos)
IBeta2Gamma1 = tableRArray(tanGamma1pos, Beta2pos)
IBeta2Gamma2 = tableRArray(tanGamma2pos, Beta2pos)


IgridBetaGamma0 = K1 * IBeta0Gamma0 + K2 * IBeta1Gamma0 + K3 * IBeta2Gamma0
IgridBetaGamma1 = K1 * IBeta0Gamma1 + K2 * IBeta1Gamma1 + K3 * IBeta2Gamma1
IgridBetaGamma2 = K1 * IBeta0Gamma2 + K2 * IBeta1Gamma2 + K3 * IBeta2Gamma2

'step 2: calculation of final value of Luminance Intensity


RCalc = kGamma1 * IgridBetaGamma0 + kGamma2 * IgridBetaGamma1 + kGamma3 * IgridBetaGamma2
 If RCalc < 0 Then
RCalc = 0
 End If
End If


End Function







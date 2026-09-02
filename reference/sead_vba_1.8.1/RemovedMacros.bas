Attribute VB_Name = "RemovedMacros"
'
'Sub UpdateSingleFixture()
'
'
'Dim column As Integer, Rcolumn As Integer, RIrow1 As Integer, RLrow1 As Integer
'Dim SBcolumn As Integer, SUcolumn As Integer, SIrow1 As Integer, SLrow1 As Integer
'Dim row As Integer
'Dim gcolumn As Integer, GRIrow1 As Integer, GRLrow1 As Integer, GBIrow1 As Integer, GUIrow1 As Integer, GBLrow1 As Integer, GULrow1 As Integer
'Dim Grid As Variant
'
'Application.ScreenUpdating = False
'
'
'Rcolumn = 3
'RIrow1 = 143
'RLrow1 = 302
'
'SBcolumn = 4
'SUcolumn = 5
'SIrow1 = 6
'SLrow1 = 13
'
''Grid points
'gcolumn = 4
'GRIrow1 = 104
'GRLrow1 = 125
'
'GBIrow1 = 9
'GUIrow1 = 30
'GBLrow1 = 58
'GULrow1 = 79
'
''=======Calculate Baseline Results===========================================================
'    Application.Calculation = xlCalculationAutomatic
'    Sheet3.Range("Base_Upgrade_Choice").Value = "Baseline"
'    RefreshIllCalcs
'    Application.Calculation = xlCalculationManual
'
'    'write results to the SResults sheet
'    column = SBcolumn
'
'
'        'Sheet 12 is Sresults - no longer used
'        'Illuminance
'        For row = 0 To 3
'            Sheet12.Cells(SIrow1 + row, column) = Sheet3.Cells(RIrow1 + row, Rcolumn)
'        Next row
'        'Luminance
'        For row = 0 To 3
'            Sheet12.Cells(SLrow1 + row, column) = Sheet3.Cells(RLrow1 + row, Rcolumn)
'        Next row
'
'    'Write results to SResultsGrid sheet
'    column = 3
'
'        With wksLuminanceOutputCIE
'            'illuminance
'            Grid = .Range(.Cells(GRIrow1, column), .Cells(GRIrow1 + 18, column + 20))
'            .Range(.Cells(GBIrow1, column), .Cells(GBIrow1 + 18, column + 20)) = Grid
'
'            'Luminance
'            Grid = .Range(.Cells(GRLrow1, column), .Cells(GRLrow1 + 18, column + 20))
'            .Range(.Cells(GBLrow1, column), .Cells(GBLrow1 + 18, column + 20)) = Grid
'
'        End With
'    Application.Calculation = xlCalculationAutomatic
'
'
''========Calculate Upgrade Results===========================================================
'    Application.Calculation = xlCalculationAutomatic
'    Sheet3.Range("Base_Upgrade_Choice").Value = "Upgrade"
'    RefreshIllCalcs
'    Application.Calculation = xlCalculationManual
'
'    'write results to the SResults sheet
'    column = SUcolumn
'
'        'Sheet 12 is Sresults - no longer used
'        'Illuminance
'        For row = 0 To 3
'            Sheet12.Cells(SIrow1 + row, column) = Sheet3.Cells(RIrow1 + row, Rcolumn)
'        Next row
'        'luminance
'        For row = 0 To 3
'            Sheet12.Cells(SLrow1 + row, column) = Sheet3.Cells(RLrow1 + row, Rcolumn)
'        Next row
'
'    'Write results to the SResultsGrid sheet
'        column = 3
'        With wksLuminanceOutputCIE
'            'illuminance
'            Grid = .Range(.Cells(GRIrow1, column), .Cells(GRIrow1 + 18, column + 20))
'            .Range(.Cells(GUIrow1, column), .Cells(GUIrow1 + 18, column + 20)) = Grid
'
'            'Luminance
'            Grid = .Range(.Cells(GRLrow1, column), .Cells(GRLrow1 + 18, column + 20))
'            .Range(.Cells(GULrow1, column), .Cells(GULrow1 + 18, column + 20)) = Grid
'
'        End With
'
'Application.ScreenUpdating = True
'
'Application.Calculation = xlCalculationAutomatic
'
'End Sub
'
'Sub Parametric()
'
'
'Dim i, c, j, k, r As Integer
'Dim num_fixtures As Integer
'Dim lastRow, lastCol As Integer
'Dim count As Integer
'Dim Spacing As Integer
'
'
'
''show the multifixture page
'Sheet11.Select
'
'Application.Calculation = xlCalculationManual
'
'
'Max_Spacing = 100
'Spacing_Increments = 5
'
'num_fixtures = Application.CountA(Sheet6.Range("Fixturechoices"))
'lastRow = Sheet11.Cells.find(What:="*", After:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
'lastCol = Sheet11.Cells.find(What:="*", After:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).column
'
'ReDim EnergyPerKm(Max_Spacing / Spacing_Increments, num_fixtures)
'ReDim PassFail(Max_Spacing / Spacing_Increments, num_fixtures)
'ReDim LowestEnergySpacing(num_fixtures, 2)
'
'For Spacing = 1 To Max_Spacing / Spacing_Increments
'
'    Range("UPoleSpacing") = Spacing * 5
'
'    For i = 1 To num_fixtures
'        Range("Ufixturechoice") = i
'        Application.Calculation = xlCalculationAutomatic
'        Application.Calculation = xlCalculationManual
'
'            Sheet11.Cells(lastRow + i, 1) = i
'            For c = 2 To lastCol
'                Sheet11.Cells(lastRow + i, c) = Sheet11.Cells(3, c)
'            Next c
'
'            'Col 7 is kWh/kilometer/yr, 16 is pass fail
'            EnergyPerKm(Spacing, i) = Sheet11.Cells(3, 7)
'            PassFail(Spacing, i) = Sheet11.Cells(3, 16)
'
'            If LowestEnergySpacing(i) = "" Then
'                LowestEnergySpacing(i, 1) = Spacing
'                LowestEnergySpacing(i) = EnergyPerKm(Spacing, i)
'            End If
'            If PassFail(Spacing, i) = 1 And EnergyPerKm(Spacing, i) < LowestEnergySpacing(Spacing) Then
'                LowestEnergySpacing(i, 1) = Spacing
'                LowestEnergySpacing(i, 2) = EnergyPerKm(Spacing, i)
'            End If
'
'    Next i
'Next Spacing
'
'
'
'
'
''sort the results in ascending order by kWh/km/yr
'Sheet11.Sort.SortFields.Clear
'Sheet11.Sort.SortFields.Add Key:=Range(Cells(4, 6), Cells(lastRow + i, 6)), _
'        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
'    With Sheet11.Sort
'        .SetRange Range(Cells(4, 1), Cells(lastRow + i, lastCol))
'        .Header = xlNo
'        .MatchCase = False
'        .Orientation = xlTopToBottom
'        .SortMethod = xlPinYin
'        .Apply
'    End With
'
'
'
'
'Application.Calculation = xlCalculationAutomatic
'End Sub
'
'
'Function anglePhi(fixtureX, fixtureY, gridXY(), polespacing, FixtureHeight, angleOnX, angleOnY, angleOnZ, calculationmethod) As Variant
''arrays to carry all x and y values
'Dim outputX()
'Dim outputY()
'outputX = gridXY(0)
'outputY = gridXY(1)
'
''inputs
'If Sheets("FixtureData").Range("A6").Value = "Baseline" Then
'    lanewidth = Sheets("Road Geometry").Range("bLaneWidth").Value
'    MedianLength = Sheets("Road Geometry").Range("bMedianWidth").Value
'    NumberOfLanes = Sheets("Road Geometry").Range("bNumLanes").Value
'    poleconfig = Sheets("Road Geometry").Range("bFixtureArrangement").Value
'ElseIf Sheets("FixtureData").Range("A6").Value = "Upgrade" Then
'    lanewidth = Sheets("Road Geometry").Range("uLaneWidth").Value
'    MedianLength = Sheets("Road Geometry").Range("uMedianWidth").Value
'    NumberOfLanes = Sheets("Road Geometry").Range("uNumLanes").Value
'    poleconfig = Sheets("Road Geometry").Range("uFixtureArrangement").Value
'End If
'
''grid start and end
'If calculationmethod = "IES" Then
'    istart = WorksheetFunction.Match(polespacing, outputX, True)
'    iend = WorksheetFunction.Match(2 * polespacing, outputX, True) - 1
'ElseIf calculationmethod = "CIE" Then
'    'start at what fixture
'    startfixture = Int(5 * FixtureHeight / polespacing)
'    startfixture = startfixture + 1
'    istart = WorksheetFunction.Match(polespacing * startfixture, outputX, True) + 1
'    iend = WorksheetFunction.Match(polespacing * (startfixture + 1), outputX, True)
'
'    Debug.Print "iStart in anglePhi is " & istart
'    Debug.Print "iEnd in anglePhi is " & iend
'    'iStart = WorksheetFunction.Match(5 * FixtureHeight, outputX, True)
'    'iEnd = WorksheetFunction.Match(5 * FixtureHeight + polespacing, outputX, True)
'    '**FLAG** when testing CIE make sure this is correct
'End If
'
'
''debug
'If fixtureX = 35 Then
'    stopcall = True
'End If
'
'
'numberOfX = iend - istart
'numberOfY = UBound(outputY)
'Dim phiArray()
'ReDim phiArray(istart To iend, numberOfY)
'm = outputX(1)
'For i = istart To iend
'For j = 0 To numberOfY
''distance between grid point and fixture point
''dist = Distance(fixtureX, fixtureY, outputX(i), outputY(j))
'
'    ' if pole configuration is median mounted
'    If poleconfig = "Median mounted" Then
'        distY = fixtureY
'        If fixtureY > (lanewidth * NumberOfLanes + MedianLength) / 2 Then           'pole is located on far side of road
'                If distY - outputY(j) > 0 Then
'                   phiArray(i, j) = 180 - Atn(Abs(fixtureX - outputX(i)) / Abs(distY - outputY(j))) * 180 / WorksheetFunction.Pi
'                ElseIf distY - outputY(j) < 0 Then
'                   phiArray(i, j) = (Atn(Abs(fixtureX - outputX(i)) / Abs(distY - outputY(j))) * 180 / WorksheetFunction.Pi)
'                Else
'                    phiArray(i, j) = 90
'                End If
'        ElseIf fixtureY < lanewidth * NumberOfLanes + MedianLength / 2 Then         'pole is located on near side of road
'                If distY - outputY(j) > 0 Then
'                   phiArray(i, j) = Atn(Abs(fixtureX - outputX(i)) / Abs(distY - outputY(j))) * 180 / WorksheetFunction.Pi
'                ElseIf distY - outputY(j) < 0 Then
'                   phiArray(i, j) = 180 - (Atn(Abs(fixtureX - outputX(i)) / Abs(distY - outputY(j))) * 180 / WorksheetFunction.Pi)
'                Else
'                    phiArray(i, j) = 90
'                End If
'        End If
'    Else
'    'For all configurations except median mounted
'            distY = fixtureY                       'grid measurement including tilt
'            x = (outputX(i) - fixtureX)
'            If fixtureY < (lanewidth * NumberOfLanes / 2) Then  'if the fixture is located on the near side of the road
'                y = (outputY(j) - fixtureY)
'            Else
'                y = (fixtureY - outputY(j))
'            End If
'
'            If y > 0 Then
'                phiTemp = Atn(Abs(x) / Abs(y)) * 180 / WorksheetFunction.Pi 'FLAG changed to distY
'            'FLAG check which version the section below was originally commented out in. This is needed for accurate calculations (comment inserted in 1.7.6)
'            ElseIf y < 0 Then
'               phiTemp = 180 - (Atn(Abs(x) / Abs(y)) * 180 / WorksheetFunction.Pi)
'            Else
'                phiTemp = 90
'            End If
'            phiArray(i, j) = phiTemp
'    End If
'
'Next j
'Next i
'anglePhi = phiArray
'
'End Function
'
'
'Function angleGamma(fixtureX, fixtureY, gridXY, polespacing, FixtureHeight, angleOnX, angleOnY, angleOnZ, calculationmethod) As Variant
'Dim phi() As Integer
'
''arrays to carry all x and y values
'Dim outputX()
'Dim outputY()
'outputX = gridXY(0)
'outputY = gridXY(1)
'
''grid start and end
'If calculationmethod = "IES" Then
'    istart = WorksheetFunction.Match(polespacing, outputX, True)
'    iend = WorksheetFunction.Match(2 * polespacing, outputX, True) - 1
'ElseIf calculationmethod = "CIE" Then
''start at what fixture
'    startfixture = Int(5 * FixtureHeight / polespacing)
'    startfixture = startfixture + 1
'    istart = WorksheetFunction.Match(polespacing * startfixture, outputX, True) + 1
'    iend = WorksheetFunction.Match(polespacing * (startfixture + 1), outputX, True)
'    Debug.Print "iStart in angleGamma is " & istart
'    Debug.Print "iEnd in angleGamma is " & iend
'    'iStart = WorksheetFunction.Match(5 * FixtureHeight, outputX, True)
'    'iEnd = WorksheetFunction.Match(5 * FixtureHeight + polespacing, outputX, True)
'End If
'
''X is along the road. Y is across the road.
'numberOfX = iend - istart
'numberOfY = UBound(outputY)
'
'Dim gammaArray()
'ReDim gammaArray(istart To iend, numberOfY)
'm = outputX(1)
'For i = istart To iend
'For j = 0 To numberOfY
'    distY = fixtureY                     'grid measurement including tilt
'    dist = Distance(fixtureX, distY, outputX(i), outputY(j))
'    'gamma at each grid point
'    If dist <> 0 Then
'        gammaArray(i, j) = (Atn(dist / FixtureHeight)) * 180 / WorksheetFunction.Pi
'    Else
'        gammaArray(i, j) = 0
'    End If
'Next
'Next
'angleGamma = gammaArray
'End Function
'
'
'
'
'Sub ReadSingleISOfile()
''This macro works as of version 1.6.0; however, it does not allow for optional input to override defaults
''Single upload option removed to simplify use of the tool, training, updates, etc.
''Read in the text values from the Translation tab to ensure proper translation
'isoprompt = Sheet25.Range("isoprompt")
'ISOFixtureprompt = Sheet25.Range("isofixtureprompt")
'ISOFixturetitle = Sheet25.Range("isofixturetitle")
'ISOFixtureDefault = Sheet25.Range("isofixturedefault")
'ISOTypeDefault = Sheet25.Range("isotypedefault")
'ISOTypeTitle = Sheet25.Range("isotypetitle")
'
'ISOFixtureError = Sheet25.Range("ISOFixtureError")
'ISOAngleError = Sheet25.Range("isoangleerror")
'
'
'TooManyColumns = False 'test to see if the Candela data will fit - set to false initialy
'OtherError = False 'again, set other error to false initially
'
'Select Case MsgBox(isoprompt, _
'            vbOKCancel)
'Case vbCancel
'    Exit Sub
'Case vbOK
'    GetOpenFilename
'    If filepath = False Then Exit Sub
'    ISOpath = filepath
'
'    FixtureName = Application.InputBox(prompt:=ISOFixtureprompt, _
'              Title:=ISOFixturetitle, Default:=ISOFixtureDefault)
'    If FixtureName = False Then Exit Sub
'
'    'Change - should be better input box, drop down choices, form?***************************************************************
'    FixtureType = Application.InputBox(prompt:=ISOTypeDefault, _
'              Title:=ISOTypeTitle, Default:="HPS, LED or MH")
'
'    If FixtureType = False Then Exit Sub
'    If FixtureType <> "LED" And FixtureType <> "HPS" And FixtureType <> "MH" Then
'        MsgBox (ISOFixtureError)
'        Exit Sub
'    End If
'
'    'Initialize to defaults
'    EquipmentCost = Range(FixtureType & "cost")
'    InstallationCost = Range(FixtureType & "instcost")
'    Rebate = 0
'    MaintenanceCost = Range(FixtureType & "maintcost")
'    MaintenanceInflation = Range(FixtureType & "maintinflate")
'    LLD = Range(FixtureType & "LLD")
'    LDD = Range(FixtureType & "DD")
'    BF = Range(FixtureType & "BF")
'
'    ReadISOfile
'
'    'Error alerts
'    If TooManyColumns = True Then
'        MsgBox (ISOAngleError)
'    ElseIf OtherError = True Then
'        '********************************************************need to change to translated error message************************************************
'        MsgBox (FixtureName & " could not be uploaded due to an error. Please check the file and ensure it is in the proper format, or choose another file. This file will be skipped.")
'    End If
'
'End Select
'End Sub

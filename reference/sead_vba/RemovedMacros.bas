Attribute VB_Name = "RemovedMacros"

Sub UpdateSingleFixture()


Dim column As Integer, Rcolumn As Integer, RIrow1 As Integer, RLrow1 As Integer
Dim SBcolumn As Integer, SUcolumn As Integer, SIrow1 As Integer, SLrow1 As Integer
Dim row As Integer
Dim gcolumn As Integer, GRIrow1 As Integer, GRLrow1 As Integer, GBIrow1 As Integer, GUIrow1 As Integer, GBLrow1 As Integer, GULrow1 As Integer
Dim Grid As Variant

Application.ScreenUpdating = False


Rcolumn = 3
RIrow1 = 143
RLrow1 = 302

SBcolumn = 4
SUcolumn = 5
SIrow1 = 6
SLrow1 = 13

'Grid points
gcolumn = 4
GRIrow1 = 104
GRLrow1 = 125

GBIrow1 = 9
GUIrow1 = 30
GBLrow1 = 58
GULrow1 = 79

'=======Calculate Baseline Results===========================================================
    Application.Calculation = xlCalculationAutomatic
    Sheet3.Range("Base_Upgrade_Choice").Value = "Baseline"
    RefreshIllCalcs
    Application.Calculation = xlCalculationManual
    
    'write results to the SResults sheet
    column = SBcolumn
        
        
        'Sheet 12 is Sresults - no longer used
        'Illuminance
        For row = 0 To 3
            Sheet12.Cells(SIrow1 + row, column) = Sheet3.Cells(RIrow1 + row, Rcolumn)
        Next row
        'Luminance
        For row = 0 To 3
            Sheet12.Cells(SLrow1 + row, column) = Sheet3.Cells(RLrow1 + row, Rcolumn)
        Next row
        
    'Write results to SResultsGrid sheet
    column = 3
    
        With Sheet22
            'illuminance
            Grid = .Range(.Cells(GRIrow1, column), .Cells(GRIrow1 + 18, column + 20))
            .Range(.Cells(GBIrow1, column), .Cells(GBIrow1 + 18, column + 20)) = Grid
            
            'Luminance
            Grid = .Range(.Cells(GRLrow1, column), .Cells(GRLrow1 + 18, column + 20))
            .Range(.Cells(GBLrow1, column), .Cells(GBLrow1 + 18, column + 20)) = Grid
            
        End With
    Application.Calculation = xlCalculationAutomatic
         
        
'========Calculate Upgrade Results===========================================================
    Application.Calculation = xlCalculationAutomatic
    Sheet3.Range("Base_Upgrade_Choice").Value = "Upgrade"
    RefreshIllCalcs
    Application.Calculation = xlCalculationManual
    
    'write results to the SResults sheet
    column = SUcolumn
        
        'Sheet 12 is Sresults - no longer used
        'Illuminance
        For row = 0 To 3
            Sheet12.Cells(SIrow1 + row, column) = Sheet3.Cells(RIrow1 + row, Rcolumn)
        Next row
        'luminance
        For row = 0 To 3
            Sheet12.Cells(SLrow1 + row, column) = Sheet3.Cells(RLrow1 + row, Rcolumn)
        Next row

    'Write results to the SResultsGrid sheet
        column = 3
        With Sheet22
            'illuminance
            Grid = .Range(.Cells(GRIrow1, column), .Cells(GRIrow1 + 18, column + 20))
            .Range(.Cells(GUIrow1, column), .Cells(GUIrow1 + 18, column + 20)) = Grid
            
            'Luminance
            Grid = .Range(.Cells(GRLrow1, column), .Cells(GRLrow1 + 18, column + 20))
            .Range(.Cells(GULrow1, column), .Cells(GULrow1 + 18, column + 20)) = Grid
            
        End With

Application.ScreenUpdating = True

Application.Calculation = xlCalculationAutomatic

End Sub

Sub Parametric()


Dim i, c, j, k, r As Integer
Dim num_fixtures As Integer
Dim lastRow, lastCol As Integer
Dim count As Integer
Dim Spacing As Integer



'show the multifixture page
Sheet11.Select

Application.Calculation = xlCalculationManual


Max_Spacing = 100
Spacing_Increments = 5

num_fixtures = Application.CountA(Sheet6.Range("Fixturechoices"))
lastRow = Sheet11.Cells.find(What:="*", after:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
lastCol = Sheet11.Cells.find(What:="*", after:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).column

ReDim EnergyPerKm(Max_Spacing / Spacing_Increments, num_fixtures)
ReDim PassFail(Max_Spacing / Spacing_Increments, num_fixtures)
ReDim LowestEnergySpacing(num_fixtures, 2)

For Spacing = 1 To Max_Spacing / Spacing_Increments

    Range("UPoleSpacing") = Spacing * 5

    For i = 1 To num_fixtures
        Range("Ufixturechoice") = i
        Application.Calculation = xlCalculationAutomatic
        Application.Calculation = xlCalculationManual

            Sheet11.Cells(lastRow + i, 1) = i
            For c = 2 To lastCol
                Sheet11.Cells(lastRow + i, c) = Sheet11.Cells(3, c)
            Next c

            'Col 7 is kWh/kilometer/yr, 16 is pass fail
            EnergyPerKm(Spacing, i) = Sheet11.Cells(3, 7)
            PassFail(Spacing, i) = Sheet11.Cells(3, 16)

            If LowestEnergySpacing(i) = "" Then
                LowestEnergySpacing(i, 1) = Spacing
                LowestEnergySpacing(i) = EnergyPerKm(Spacing, i)
            End If
            If PassFail(Spacing, i) = 1 And EnergyPerKm(Spacing, i) < LowestEnergySpacing(Spacing) Then
                LowestEnergySpacing(i, 1) = Spacing
                LowestEnergySpacing(i, 2) = EnergyPerKm(Spacing, i)
            End If

    Next i
Next Spacing





'sort the results in ascending order by kWh/km/yr
Sheet11.Sort.SortFields.Clear
Sheet11.Sort.SortFields.Add Key:=Range(Cells(4, 6), Cells(lastRow + i, 6)), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With Sheet11.Sort
        .SetRange Range(Cells(4, 1), Cells(lastRow + i, lastCol))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With




Application.Calculation = xlCalculationAutomatic
End Sub




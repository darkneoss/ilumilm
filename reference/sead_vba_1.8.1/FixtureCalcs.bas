Attribute VB_Name = "FixtureCalcs"
Option Explicit
Public ExitFlag As Boolean
Sub RunAllFixtures()


Dim i, c, j, k, r, x As Integer
Dim num_fixtures As Integer
Dim lastRow, lastCol, LastRow13, LastCol13 As Integer
Dim count, Namecheckcount, NameCheck As Integer
Dim GroupName, GroupNamei As Variant
Dim FixtureChoice As String
Dim continue, NameUsed As Boolean
Dim MResults() As Variant
Dim BaselineRow As Integer
Dim DeletePRompt2 As String
Dim ClearResults As Variant

Dim PromptName As String
Dim DefaultName As String, GroupNameHeader As String

Dim UTotalInstallCost, UTotalAnnualCost, InstallCostDifference, BTotalInstallCost, AnnualCostSavings, BTotalAnnualCost As Long

PromptName = Sheet25.Range("GroupPromptName")
DefaultName = Sheet25.Range("GroupDefaultName")
GroupNameHeader = Sheet25.Range("GroupNameHeader")

'GroupName = Application.InputBox(prompt:=PromptName, _
'          Title:=GroupNameHeader, Default:=DefaultName)
'If GroupName = False Then
'    Exit Sub
'Else
GroupName = Sheet25.Range("tResults")
'DeletePRompt2 = Sheet25.Range("DeletePrompt2")
'ClearResults = MsgBox(DeletePRompt2, vbYesNo)
'If ClearResults = vbYes Then
    DeleteResults
'End If
   
'First, count the total number of fixtures to be run
Dim FirstRow13 As Integer
Dim TotalFixtures As Integer

FirstRow13 = Sheet13.Range("IncludeHeader").row
'last row on the Fixtures tab
LastRow13 = Sheet13.Cells.find(What:="*", After:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
LastCol13 = Sheet13.Cells.find(What:="*", After:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).column


num_fixtures = Application.CountA(Range("Fixturechoices"))
TotalFixtures = 0
For i = FirstRow13 + 1 To LastRow13
    If Sheet13.Cells(i, 2) = "x" Or Sheet13.Cells(i, 2) = "X" Then TotalFixtures = TotalFixtures + 1
Next i
        
'````````````````````````````````````
'open the dialog for progress display
Dim pbProgBar As IProgressBar
Set pbProgBar = New FProgressBarIFace
pbProgBar.Title = Sheet25.Range("tStatusHeader")
pbProgBar.Text = Sheet25.Range("tPerformingCalcs")
pbProgBar.Min = 0
pbProgBar.max = TotalFixtures + 1
pbProgBar.Progress = 0
pbProgBar.Show
pbProgBar.Progress = 0.1
'````````````````````````````````````

    'show the multifixture page
    Sheet10.Select
    
    'Find the last column for writing output purposes
    lastCol = Sheet10.Cells.find(What:="*", After:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).column
    
    '------------------------
    'Run the Baseline fixture
    '------------------------
    Application.Calculation = xlCalculationAutomatic
    Sheet2.Range("Base_Upgrade_Choice").Value = "Baseline"
    'Sheets("FixtureData").Range("A6") = "Baseline"
    Call finalMatrices
    If ExitFlag = True Then Exit Sub
    
    '**then write results to row 4
            BaselineRow = 4
            'Start with the name
            Sheet10.Cells(BaselineRow, 2) = Range("BaselineTranslation")
            'Write the rest of the fields into the results section
            ReDim MResults(3 To lastCol)
            For c = 3 To lastCol
                MResults(c) = Sheet10.Cells(3, c)
            Next c
            Sheet10.Select
            Range(Cells(BaselineRow, 3), Cells(BaselineRow, c - 1)) = MResults
            
            'Set up the baseline costs for use in the ROI calculations later
            BTotalInstallCost = Range("TotalInstallCost")
            BTotalAnnualCost = Range("TotalAnnualCost")
            
            Dim BTotalAnnualCostwInflation(20) As Variant
            Dim ROI(20) As Variant
            Dim aWithInflationRow As Integer, aWithInflationCol As Integer
            aWithInflationRow = Range("TotalAnnualCostwInflation").row
            aWithInflationCol = Range("TotalAnnualCostwInflation").column
            BTotalAnnualCostwInflation(0) = BTotalInstallCost
            For i = 1 To 20
                BTotalAnnualCostwInflation(i) = Sheet4.Cells(aWithInflationRow + i + 1, aWithInflationCol).Value
            Next i
            
            For i = 0 To 20
                Sheet4.Cells(aWithInflationRow + i + 1, 18) = BTotalAnnualCostwInflation(i)
            Next i

        'counter for how many fixtures you've run
        count = 1
        pbProgBar.Progress = count
    
    
    '------------------------
    'Run the upgrade fixtures
    '------------------------
    Sheet2.Range("Base_Upgrade_Choice") = "Upgrade"
    'For each fixture with an 'x' on it in the 'Fixtures' tab, calculate the results
    '------------------------
    'last row on the results tab (find after writing the baseline to row 4)
    lastRow = Sheet10.Cells.find(What:="*", After:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row

    'For each fixture on the Fixtures tab
    For i = FirstRow13 + 1 To LastRow13
        'Check row for an x to decide whether to run or not
        If Sheet13.Cells(i, 2) <> "x" And Sheet13.Cells(i, 2) <> "X" Then
            'do nothing, go to next row
        ElseIf Sheet13.Cells(i, 2) = "x" Or Sheet13.Cells(i, 2) = "X" Then
            Application.Calculation = xlCalculationAutomatic

            'Write the ID number of the fixture in the array to the FixtureData tab to change calculations
            FixtureChoice = Application.WorksheetFunction.Match(Sheet13.Cells(i, 3), Range("FixtureChoices"), 0)
            Range("Ufixturechoice") = FixtureChoice
            Call finalMatrices
           
            'FLAG this should be turned into a subroutine if possible.
            'Write the group name and simulation count to the results tab; if the name already exists, add "_1 - "
            '**first check for unique names
                'GroupNamei is the name to be written - test it to make sure it doesn't exist
                GroupNamei = GroupName & " - " & count
                'initialize variables to check used names
                continue = True 'assume the name is already in use been used until confirmed otherwise
                Namecheckcount = 1
                
                Do While continue = True
                    NameUsed = False
                    For NameCheck = 4 To lastRow
                        If Sheet10.Cells(NameCheck, 2) = GroupNamei Then
                            GroupNamei = GroupName & "_" & Namecheckcount & " - " & count
                            Namecheckcount = Namecheckcount + 1
                            NameUsed = True
                        End If
                    Next NameCheck
                    
                    If NameUsed = True Then
                        continue = True
                    Else
                        continue = False
                    End If
                Loop
                
            '**then write SUMMMARY results to the next available row
            'Start with the name
            Sheet10.Cells(lastRow + count, 2) = GroupNamei
            'Write the rest of the fields into the results section
            ReDim MResults(3 To lastCol)
            For c = 3 To lastCol
                MResults(c) = Sheet10.Cells(3, c)
            Next c
            Sheet10.Select
            Range(Cells(lastRow + count, 3), Cells(lastRow + count, c - 1)) = MResults
            
            'Calculate the values that need to be compared to the baseline
            UTotalInstallCost = Sheet10.Range("Z3")
            UTotalAnnualCost = Sheet10.Range("AA3")
            InstallCostDifference = UTotalInstallCost - BTotalInstallCost
            AnnualCostSavings = BTotalAnnualCost - UTotalAnnualCost
            
            Sheet10.Cells(lastRow + count, 28) = InstallCostDifference
            Sheet10.Cells(lastRow + count, 29) = AnnualCostSavings
            If AnnualCostSavings = 0 Then
                Sheet10.Cells(lastRow + count, 30) = 0
            Else
                Sheet10.Cells(lastRow + count, 30) = InstallCostDifference / AnnualCostSavings
            End If
                       
            'Return on Investment
            For x = 1 To 20
                ROI(x) = Sheet4.Cells(23 + x, 21).Value
            Next x
            
            For x = 1 To 20
                Sheet10.Cells(lastRow + count, 51 + x) = ROI(x)
            Next x
            
            'Update the counter before moving onto the next row
            count = count + 1
            pbProgBar.Progress = count
        End If
    Next i

    pbProgBar.Hide

'If was in case they cancelled out of the dialog box - end it and wrap up the sub
'End If
End Sub


'================code used for running all fixtures - not used========
'    For i = 1 To num_fixtures
'        Range("Ufixturechoice") = i
'        RefreshIllCalcs
'
'            Sheet10.Cells(LastRow + i, 1) = i
'            Sheet10.Cells(LastRow + i, 2) = GroupName & " " & i
'            For c = 3 To LastCol
'                Sheet10.Cells(LastRow + i, c) = Sheet10.Cells(3, c)
'            Next c
'    Next i
'=====================================================================





Public Sub CenterShape(o As Shape)
o.Left = ActiveWindow.VisibleRange(1).Left + (ActiveWindow.VisibleRange.Width / 2 - o.Width / 2)
o.Top = ActiveWindow.VisibleRange(1).Top + (ActiveWindow.VisibleRange.Height / 2 - o.Height / 2)
End Sub




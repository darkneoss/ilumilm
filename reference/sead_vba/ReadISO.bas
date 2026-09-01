Attribute VB_Name = "ReadISO"
Dim filename As Variant
Dim filepath As Variant
Dim namelistrow As Integer
Dim ISOpath As String
Dim FixtureName As Variant
Dim FixtureType As Variant
Dim EquipmentCost As Double
Dim InstallationCost As Double
Dim Rebate As Double
Dim MaintenanceCost As Double
Dim MaintenanceInflation As Double
Dim LLD As Double
Dim LDD As Double
Dim BF As Double

Dim TooManyColumns As Boolean
Dim OtherError As Boolean

Public Enum EExcelVersion
    Excel2003 = 11
    Excel2007 = 12
    Excel2010 = 14
End Enum
'Needs to be edited -
Sub ListAllFiles()
    Dim fs As FileSearch, ws As Worksheet, i As Long, filenamelength As Integer
    Dim MyObj As Object, MySource As Object, file As Variant
    lastRow = Sheet21.Cells.find(What:="*", after:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    
    Message = Sheet25.Range("SelectIESPathmsg")
    Select Case MsgBox(Message, _
            vbOKCancel)
    Case vbCancel
        Exit Sub
    Case vbOK
        GetOpenFilename
        If filepath = False Then Exit Sub
        filenamelength = Len(filename)
        filepathlength = Len(filepath)
        filepathonly = Left(filepath, filepathlength - filenamelength)
    
        file = Dir(filepathonly)
        row = lastRow + 1
        While (file <> "")
          If InStr(file, ".ies") > 0 Or InStr(file, ".IES") > 0 Then
             Sheet21.Cells(row, 1) = filepathonly & file
             Sheet21.Cells(row, 2) = file
             row = row + 1
          End If
        file = Dir
        Wend
    End Select
    Message = Sheet25.Range("FixtureAddMessage")
    MsgBox (Message)
End Sub
Sub ReadMultipleISOFile()
'Get ready - unlock the Fixtures tab
Sheet13.Unprotect


 lastRow = Sheet21.Cells.find(What:="*", after:=[a1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
 
 Blankerror = False
 For row = 21 To lastRow
    ISOpath = Sheet21.Cells(row, 1)
    FixtureName = Sheet21.Cells(row, 2)
    FixtureType = Sheet21.Cells(row, 3)
    
    If ISOpath = "" Or FixtureName = "" Or FixtureType = "" Then Blankerror = True
    
 Next row
 
 If Blankerror = True Then
    Message = Sheet25.Range("Blankerror")
    Select Case MsgBox(Message, vbOKCancel)
    Case vbCancel
        Exit Sub
    End Select
 End If
    
'-------------
'Display the 'calculating' message
'open the dialog for progress display
Dim pbProgBar As IProgressBar
Set pbProgBar = New FProgressBarIFace
pbProgBar.Title = Sheet25.Range("tStatusHeader")
pbProgBar.Text = Sheet25.Range("tUploadingFixtures")
pbProgBar.Min = 0
pbProgBar.Max = lastRow - 20
pbProgBar.Progress = 0
pbProgBar.Show
pbProgBar.Progress = 0.1
'-----------------

Fixturecount = 1
ErrorCol = 12
For row = 21 To lastRow

    'Get info from the sheet about the file to be loaded - name, location, type
    pbProgBar.Progress = Fixturecount 'for use in the 'now uploading' box
    ISOpath = Sheet21.Cells(row, 1)
    FixtureName = Sheet21.Cells(row, 2)
    FixtureType = Sheet21.Cells(row, 3)
    'Translate the acronyms for FixtureType if necessary. See "FixtureTypeChoice" named range on the Assumptions tab
    '                 English              Spanish               French
    If FixtureType = "MH" Or FixtureType = "AM" Or FixtureType = "MH" Then FixtureType = "MH"
    If FixtureType = "HPS" Or FixtureType = "SAP" Or FixtureType = "SHP" Then FixtureType = "HPS"
    If FixtureType = "LED" Or FixtureType = "LED" Or FixtureType = "LED" Then FixtureType = "LED"
    
    'Get optional data
    'Initialize to defaults
    EquipmentCost = Range(FixtureType & "cost")
    InstallationCost = Range(FixtureType & "instcost")
    Rebate = 0
    MaintenanceCost = Range(FixtureType & "maintcost")
    MaintenanceInflation = Range(FixtureType & "maintinflate")
    LLD = Range(FixtureType & "LLD")
    LDD = Range(FixtureType & "DD")
    BF = Range(FixtureType & "BF")
    
    'If values are entered on the row, overwrite the default
    If Sheet21.Cells(row, 4).Value <> "" Then EquipmentCost = Sheet21.Cells(row, 4).Value
    If Sheet21.Cells(row, 5).Value <> "" Then InstallationCost = Sheet21.Cells(row, 5).Value
    If Sheet21.Cells(row, 6).Value <> "" Then Rebate = Sheet21.Cells(row, 6).Value
    If Sheet21.Cells(row, 7).Value <> "" Then MaintenanceCost = Sheet21.Cells(row, 7).Value
    If Sheet21.Cells(row, 8).Value <> "" Then MaintenanceInflation = Sheet21.Cells(row, 8).Value
    If Sheet21.Cells(row, 9).Value <> "" Then LLD = Sheet21.Cells(row, 9).Value
    If Sheet21.Cells(row, 10).Value <> "" Then LDD = Sheet21.Cells(row, 10).Value
    If Sheet21.Cells(row, 11).Value <> "" Then BF = Sheet21.Cells(row, 11).Value

    'Initialize Error variables
    TooManyColumns = False 'test to see if the Candela data will fit - set to false initialy
    OtherError = False 'again, set other error to false initially
    
    'get the data, write it to the sheet
    ReadISOfile
    
    'If errors prevented writing to the sheet, report the error
    If TooManyColumns = True Then
        Sheet21.Cells(row, ErrorCol) = "Too Many Columns"
    ElseIf OtherError = True Then
        Sheet21.Cells(row, ErrorCol) = "Other Error"
    End If
    
    Fixturecount = Fixturecount + 1
Next row
pbProgBar.Hide

msg = Sheet25.Range("tUploadComplete")
MsgBox (msg)

Sheet13.Protect
End Sub

Sub ReadSingleISOfile()
'This macro works as of version 1.6.0; however, it does not allow for optional input to override defaults
'Single upload option removed to simplify use of the tool, training, updates, etc.
'Read in the text values from the Translation tab to ensure proper translation
isoprompt = Sheet25.Range("isoprompt")
ISOFixtureprompt = Sheet25.Range("isofixtureprompt")
ISOFixturetitle = Sheet25.Range("isofixturetitle")
ISOFixtureDefault = Sheet25.Range("isofixturedefault")
ISOTypeDefault = Sheet25.Range("isotypedefault")
ISOTypeTitle = Sheet25.Range("isotypetitle")

ISOFixtureError = Sheet25.Range("ISOFixtureError")
ISOAngleError = Sheet25.Range("isoangleerror")


TooManyColumns = False 'test to see if the Candela data will fit - set to false initialy
OtherError = False 'again, set other error to false initially

Select Case MsgBox(isoprompt, _
            vbOKCancel)
Case vbCancel
    Exit Sub
Case vbOK
    GetOpenFilename
    If filepath = False Then Exit Sub
    ISOpath = filepath
    
    FixtureName = Application.InputBox(prompt:=ISOFixtureprompt, _
              Title:=ISOFixturetitle, Default:=ISOFixtureDefault)
    If FixtureName = False Then Exit Sub
    
    'Change - should be better input box, drop down choices, form?***************************************************************
    FixtureType = Application.InputBox(prompt:=ISOTypeDefault, _
              Title:=ISOTypeTitle, Default:="HPS, LED or MH")
    
    If FixtureType = False Then Exit Sub
    If FixtureType <> "LED" And FixtureType <> "HPS" And FixtureType <> "MH" Then
        MsgBox (ISOFixtureError)
        Exit Sub
    End If
    
    'Initialize to defaults
    EquipmentCost = Range(FixtureType & "cost")
    InstallationCost = Range(FixtureType & "instcost")
    Rebate = 0
    MaintenanceCost = Range(FixtureType & "maintcost")
    MaintenanceInflation = Range(FixtureType & "maintinflate")
    LLD = Range(FixtureType & "LLD")
    LDD = Range(FixtureType & "DD")
    BF = Range(FixtureType & "BF")
    
    ReadISOfile
    
    'Error alerts
    If TooManyColumns = True Then
        MsgBox (ISOAngleError)
    ElseIf OtherError = True Then
        '********************************************************need to change to translated error message************************************************
        MsgBox (FixtureName & " could not be uploaded due to an error. Please check the file and ensure it is in the proper format, or choose another file. This file will be skipped.")
    End If
    
End Select
End Sub
Sub check_compatibility()


Mode = ActiveWorkbook.Excel8CompatibilityMode


End Sub



Sub ReadISOfile()

Dim oFSO As New FileSystemObject
Dim oFS

Dim count As Integer
Dim NoKeywords As Boolean
Dim TiltRows As Integer
Dim DataRow1, DataRow2, VertAngleRow, HorizAngleRow As Integer
'Dim format As Boolean
Dim sText1, sText2 As Integer
Dim sValue As Variant
Dim sText As String
Dim sSpace As Integer
Dim Space As String
Dim VarCount1, VarCount2 As Integer
Dim NumLamps, Lumens, CandelaMult, NumVertAngles, NumHorizAngles, PhotoType, UnitType As Long
Dim VertAngleFilled, HorizAngleFilled As Boolean
Dim Manufac, Distribution, LumCat As String

Dim HorizAngles(), VertAngles(), CandelaValues() As Long

On Error GoTo Err1:

Set oFS = oFSO.OpenTextFile(ISOpath)

row = 1
NoKeywords = False

'set first position of candela array values
i = 1
j = 0

VertAngleFilled = False
HorizAngleFilled = False
AngleCount = 0


Do Until oFS.AtEndOfStream
    'Read line by line
    sText = oFS.ReadLine
    'Test if it's in 2002 format
'    If Row = 1 Then
'        If sText = "IESNA:LM-63-2002" Or sText = "IESNA:LM-63-1995" Then
'            format = True
'        Else
'            MsgBox "File is not in the appropriate format!"
'            Exit Sub
'        End If
'    End If
    
    'Look for Keyword Data
    If Left(sText, 9) = "[MANUFAC]" Then
        Manufac = Trim(Right(sText, Len(sText) - 9))
    End If
    
    If Left(sText, 14) = "[DISTRIBUTION]" Then
        Distribution = Trim(Right(sText, Len(sText) - 14))
    End If
    
    If Left(sText, 8) = "[LUMCAT]" Then
        LumCat = Trim(Right(sText, Len(sText) - 8))
    End If
    

    'Look for the end of the flexible variable keyword section
    If Left(sText, 4) = "TILT" Then
        If sText = "TILT=INCLUDE" Then
            TiltRows = 4
        End If
        
        'Initialize variables for identifying remaining data
        NoKeywords = True
        DataRow1 = row + TiltRows + 1   '<number of lamps> <lumens per lamp> <candela multiplier>
                                        '<number of vertical angles> <number of horizontal angles>
                                        '<photometric type><units type><width><length><height>
        
        DataRow2 = row + TiltRows + 2   '<ballast factor><future use><input watts>
        VertAngleRow = row + TiltRows + 3 'Horiz anglerow is initiated once vert is filled
    End If
    
    
Space = " "
    'Assign variables for Data row 1
    If row = DataRow1 Then
        VarCount1 = 0
        sValue = 0
        Do While Len(sText) > 0
            sSpace = InStr(sText, Space)
            If sSpace = 0 Then
                sValue = sText
                sText = ""
            ElseIf sSpace <> 0 Then
                sValue = Left(sText, sSpace - 1)
                sText = Right(sText, Len(sText) - sSpace)
            End If
            
            'Ignore duplicate spaces, otherwise count up the variables and assign them appropriately
            If sValue <> "" Then
                VarCount1 = VarCount1 + 1 'count the variables used in the calculation
            
                If VarCount1 = 1 Then NumLamps = sValue
                If VarCount1 = 2 Then Lumens = CLng(sValue)
                If VarCount1 = 3 Then CandelaMult = Val(sValue)
                If VarCount1 = 4 Then NumVertAngles = CInt(sValue)
                If VarCount1 = 5 Then NumHorizAngles = CInt(sValue)
                If VarCount1 = 6 Then PhotoType = sValue
                If VarCount1 = 7 Then UnitType = sValue
                'If VarCount1 = 8 Then Width = sValue
                'If VarCount1 = 9 Then Length = sValue
                'If VarCount1 = 10 Then Height = sValue
            End If
        Loop
        
        'Set up variables based on these new inputs (size of arrays for holding candela data)
        ReDim VertAngles(NumVertAngles)
        ReDim HorizAngles(NumHorizAngles)
        ReDim CandelaValues(NumHorizAngles, NumVertAngles)

    End If
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Assign variables for Data row 2
    If row = DataRow2 Then
        VarCount2 = 0
        sValue = 0
        Do While Len(sText) > 0
            sSpace = InStr(sText, Space)
            If sSpace = 0 Then
                sValue = sText
                sText = ""
            ElseIf sSpace <> 0 Then
                sValue = Left(sText, sSpace - 1)
                sText = Right(sText, Len(sText) - sSpace)
            End If
            
            'Ignore duplicate spaces, otherwise count up the variables and assign them appropriately
            If sValue <> "" Then
                VarCount2 = VarCount2 + 1 'count the variables used in the calculation
                If VarCount2 = 1 Then BallastFactor = sValue
                If VarCount2 = 2 Then FutureUse = sValue
                If VarCount2 = 3 Then
                    InputWatts = CLng(sValue)
                End If
            End If
        Loop
    End If

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Assign variables for Vertical Angles

   If row >= VertAngleRow And VertAngleFilled = False And NoKeywords = True Then
        sValue = 0
        Do While Len(sText) > 0
            sSpace = InStr(sText, Space)
            If sSpace = 0 Then
                sValue = sText
                sText = ""
            ElseIf sSpace <> 0 Then
                sValue = Left(sText, sSpace - 1)
                sText = Right(sText, Len(sText) - sSpace)
            End If
            
            'Ignore duplicate spaces, otherwise count up the variables and assign them appropriately
            If sValue <> "" Then
                If AngleCount < NumVertAngles Then
                    AngleCount = AngleCount + 1
                    VertAngles(AngleCount) = sValue
                End If
                If AngleCount = NumVertAngles Then
                    VertAngleFilled = True
                    AngleCount = 0
                    HorizAngleRow = row + 1
                End If
            End If
        Loop
    End If

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Assign variables for Horiz Angles

     If row >= HorizAngleRow And VertAngleFilled = True And HorizAngleFilled = False Then
        sValue = 0
        Do While Len(sText) > 0
            sSpace = InStr(sText, Space)
            If sSpace = 0 Then
                sValue = sText
                sText = ""
            ElseIf sSpace <> 0 Then
                sValue = Left(sText, sSpace - 1)
                sText = Right(sText, Len(sText) - sSpace)
            End If
            
            'Ignore duplicate spaces, otherwise count up the variables and assign them appropriately
            If sValue <> "" Then
                If AngleCount < NumHorizAngles Then
                    AngleCount = AngleCount + 1
                    HorizAngles(AngleCount) = sValue
                End If
                If AngleCount = NumHorizAngles Then
                    HorizAngleFilled = True
                    AngleCount = 0
                    CandelaRow = row + 1
                End If
            End If
        Loop
    End If

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'Fill remaining candela values into large array
If row >= CandelaRow And HorizAngleFilled = True Then
        sValue = 0
        'Take a whole row of space-separated values and turn it into an array of numbers
        Do While Len(sText) > 0
            sSpace = InStr(sText, Space)
            If sSpace = 0 Then
                sValue = sText
                sText = ""
            ElseIf sSpace <> 0 Then
                sValue = Left(sText, sSpace - 1)
                sText = Right(sText, Len(sText) - sSpace)
            End If
            
            'Ignore duplicate spaces, otherwise count up the variables and assign them appropriately
            If sValue <> "" And IsNumeric(sValue) Then
                If j = NumVertAngles Then
                    j = 1
                    i = i + 1
                Else
                    j = j + 1
                End If
                CandelaValues(i, j) = sValue
            End If
        Loop



    End If

row = row + 1
Loop

'=====================================================================================================
'Write data out of variables onto sheets

'Sheet2 - fixturedata
'Sheet13 - Fixtures
lastRow = Sheet2.Cells.find(What:="*", after:=[a1], SearchOrder:=xlByRows, _
              SearchDirection:=xlPrevious).row
LastRowSheet13 = Sheet13.Cells.find(What:="*", after:=[a1], SearchDirection:=xlPrevious).row

StartRow = lastRow + 1
startcolumn = 1
introrows = Sheet2.Range("Vrow")

Dim SecondLastChar As String
'Test if FixtureName is used already; if so, add a number to the end (incrementing until a non-unique one is found)
        FixtureExists = True
        increment = 1
        Do While FixtureExists = True
            Set FixtureNameTest = Range("FixtureNames").find(What:=FixtureName, LookIn:=xlValues, LookAt:= _
                    xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
                    , SearchFormat:=False)
            If FixtureNameTest Is Nothing Then
                FixtureExists = False
            Else
                SecondLastChar = Mid(FixtureName, Len(FixtureName) - 1, 1)
                If SecondLastChar = "_" Then
                    FixtureName = Left(FixtureName, Len(FixtureName) - 2)
                End If
                FixtureName = FixtureName & "_" & increment
                increment = increment + 1
            End If
        Loop

Application.Calculation = xlCalculationManual
'Write all the information to the tabs
'First check to see if there are too many columns

'Initialize
CompatibleColumns = ""

If NumVertAngles < 256 Then 'all good - no issues in any version
    CompatibleColumns = True
Else    'check to see if we're in compatibility mode or saved as Excel 2007 format
    ExcelVersion = Application.Version
    
    If ExcelVersion = "10.0" Or ExcelVersion = "11.0" Then 'Excel 2002 or 2003
    
    ElseIf ExcelVersion = "12.0" Or ExcelVersion = "14.0" Then
        CompatibilityMode = ActiveWorkbook.Excel8CompatibilityMode
        If CompatibilityMode = True Then
            CompatibleColumns = False 'We're in a compatible version, but we're in compatibility mode so would need to resave the file as .xlsx
        ElseIf CompatibilityMode = False Then
            CompatibleColumns = True 'In a 2007 or 2010 version, and not in compatibility mode - we have extra columns to use
        End If
    End If
    

End If



If CompatibleColumns = True Then
    'If loading generic fixtures, uncomment
'    Manufac = "Generic"
'    LumCat = "NA"
        
        
    'Write Fixture to Fixtures Sheet
    Sheet13.Rows(LastRowSheet13 & ":" & LastRowSheet13).Copy Destination:=Sheet13.Range("A" & LastRowSheet13 + 1)
    Sheet13.Cells(LastRowSheet13 + 1, 3) = FixtureName
        
      'Add fixture data
    Sheet13.Cells(LastRowSheet13 + 1, 5) = Manufac
    Sheet13.Cells(LastRowSheet13 + 1, 6) = FixtureType
    Sheet13.Cells(LastRowSheet13 + 1, 7) = LumCat
    Sheet13.Cells(LastRowSheet13 + 1, 8) = Distribution
    Sheet13.Cells(LastRowSheet13 + 1, 9) = ""
    Sheet13.Cells(LastRowSheet13 + 1, 10) = InputWatts
    
     'Add optional values (default or user input as assigned in previous subroutine)
    Sheet13.Cells(LastRowSheet13 + 1, 11) = EquipmentCost
    Sheet13.Cells(LastRowSheet13 + 1, 12) = InstallationCost
    Sheet13.Cells(LastRowSheet13 + 1, 13) = Rebate
    Sheet13.Cells(LastRowSheet13 + 1, 14) = MaintenanceCost
    Sheet13.Cells(LastRowSheet13 + 1, 15) = MaintenanceInflation
    Sheet13.Cells(LastRowSheet13 + 1, 16) = LLD
    Sheet13.Cells(LastRowSheet13 + 1, 17) = LDD
    Sheet13.Cells(LastRowSheet13 + 1, 18) = BF
        
    'Intro information
    Sheet2.Cells(StartRow, 2) = FixtureName
    Sheet2.Cells(StartRow + 1, 2) = Manufac
    Sheet2.Cells(StartRow + 2, 2) = FixtureType
    Sheet2.Cells(StartRow + 3, 2) = LumCat
    Sheet2.Cells(StartRow + 4, 2) = InputWatts
    Sheet2.Cells(StartRow + 5, 2) = Distribution
    Sheet2.Cells(StartRow + 6, 2) = ""
    Sheet2.Cells(StartRow + 7, 2) = NumVertAngles
    Sheet2.Cells(StartRow + 8, 2) = NumHorizAngles
    
    'Candela data
    For i = 1 To NumHorizAngles
        Sheet2.Cells(i + StartRow + introrows, startcolumn) = HorizAngles(i)
        For j = 1 To NumVertAngles
            Sheet2.Cells(StartRow + introrows, j + startcolumn) = VertAngles(j)
            tempVal = CandelaValues(i, j)
            tempVal = format(tempVal, "#.0###")
            tempVal = Val(tempVal)
            Sheet2.Cells(i + StartRow + introrows, j + startcolumn) = tempVal * CandelaMult     'then multiply the candela value by the candela multiplier from DataRow1
        Next j
    Next i

Else 'if there are more than 256 vertical angles, skip
    TooManyColumns = True
End If


Application.Calculation = xlCalculationAutomatic
errortest = 1 / NumHorizAngles

'Error handling
Err1:
    If Err.Number <> 0 Then
        'Moving this message box into the respective load single fixture and load all fixtures subroutines
        'MsgBox (FixtureName & " could not be uploaded due to an error. Please check the file and ensure it is in the proper format, or choose another file. This file will be skipped.")
        OtherError = True
        TestNumber = Err.Number
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    End If
            TestNumber = Err.Number
    Exit Sub
Resume

End Sub

Sub GetOpenFilename()
Dim Filter As String, Title As String
Dim FilterIndex As Integer
'Dim filename As Variant

' File filters
Filter = "IES Files (*.ies),*.ies," & _
        "Text Files (*.txt),*.txt," & _
        "All Files (*.*),*.*"

'Default Filter to *.*
FilterIndex = 3

' Set Dialog Caption
Title = "Select a File to Open"

With Application
    ' Set File Name to selected File
    filepath = .GetOpenFilename(Filter, FilterIndex, Title)
    ' Reset Start Drive/Path
    On Error Resume Next
    ChDrive (Left(.DefaultFilePath, 1))
    ChDir (.DefaultFilePath)
    On Error GoTo 0
End With
' Exit on Cancel
If filepath = False Then
    MsgBox "No file was selected."
    Exit Sub
End If

filename = GetFilenameFromPath(filepath)

End Sub


 Function GetFilenameFromPath(ByVal strPath As String) As String
' Returns the rightmost characters of a string upto but not including the rightmost '\'
' e.g. 'c:\winnt\win.ini' returns 'win.ini'
    
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function



Sub deleteISOdata()
Sheet13.Unprotect

'Identify the range to be deleted on teh Fixture Data tab
    Range("Base_Upgrade_Choice") = "Delete"
    
    firstRow = Range("FixLib_FirstSelected")
    lastRow = Range("FixLib_LastSelected")
    DeleteFixtureName = Range("FxLib_SelectedFixture")
    
    Sheet2.Range("A" & firstRow & ":A" & lastRow).EntireRow.Delete
    Range("Base_Upgrade_Choice") = "Upgrade"

'Delete the row on the Fixtures tab
    Sheet13.Cells.find(What:=DeleteFixtureName, after:=ActiveCell, LookIn:=xlFormulas, LookAt:= _
        xlWhole).Range("A1:A1").EntireRow.Delete

Sheet13.Protect
End Sub



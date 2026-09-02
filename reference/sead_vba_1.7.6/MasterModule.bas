Attribute VB_Name = "MasterModule"
'This macro was developed by p2w2.  http://p2w2.com/expert-in-microsoft-excel-consultants-consulting/index.php
'Please contact at CS@perceptive-analytics.com in case of any enquiry



' Macro assigned to button in inputs sheet
Sub masterProc()
Application.ScreenUpdating = False
Dim cMethod As String
cMethod = Range("calculationMethodChoice").Value
Dim IESmsg, CIEmsg, IES1msg, CIE2msg As String

IESmsg = Sheets("Translation").Range("IESmsg").Value
CIEmsg = Sheets("Translation").Range("CIEmsg").Value
IES1msg = Sheets("Translation").Range("IES1msg").Value
CIE2msg = Sheets("Translation").Range("CIE1msg").Value

'Perform some input validation
Dim rMaster(1 To 16) As Range
Dim iCheck As Long
Dim missingFlag As Boolean

missingFlag = False

Set rMaster(1) = Sheet5.Range("bNumLanes")
Set rMaster(2) = Sheet5.Range("bLaneWidth")
Set rMaster(3) = Sheet5.Range("bMedianWidth")
Set rMaster(4) = Sheet5.Range("bMountingHeight")
Set rMaster(5) = Sheet5.Range("bPoleSpacing")
Set rMaster(6) = Sheet5.Range("bPoleSetback")
Set rMaster(7) = Sheet5.Range("bArmLength")
Set rMaster(8) = Sheet5.Range("bFixtureArrangement")

Set rMaster(9) = Sheet5.Range("uNumLanes")
Set rMaster(10) = Sheet5.Range("uLaneWidth")
Set rMaster(11) = Sheet5.Range("uMedianWidth")
Set rMaster(12) = Sheet5.Range("uMountingHeight")
Set rMaster(13) = Sheet5.Range("uPoleSpacing")
Set rMaster(14) = Sheet5.Range("uPoleSetback")
Set rMaster(15) = Sheet5.Range("uArmLength")
Set rMaster(16) = Sheet5.Range("uFixtureArrangement")
                
For iCheck = LBound(rMaster) To UBound(rMaster)
    If IsEmpty(rMaster(iCheck).Value) Then missingFlag = True
Next iCheck

If missingFlag = True Then
    prompt = Sheet25.Range("tMissingRoadGeometry")
    MsgBox (prompt)
    Exit Sub
End If





'------------------------------
If cMethod = "IES" Then
    MsgBox IESmsg
    Sheets("FixtureData").Range("iescieGraphChoice").Value = "IES"
    Call RunAllFixtures
    If ExitFlag = True Then Exit Sub
    Call outFile(cMethod)
ElseIf cMethod = "CIE" Then
    MsgBox CIEmsg
    Sheets("FixtureData").Range("iescieGraphChoice").Value = "CIE"
    Call RunAllFixtures
    If ExitFlag = True Then Exit Sub
    Call outFile(cMethod)
Else
    Sheets("FixtureData").Range("iescieGraphChoice").Value = "IES"
    MsgBox IES1msg
    Call RunAllFixtures
    If ExitFlag = True Then Exit Sub
    Call outFile("IES")

    ThisWorkbook.Sheets("FixtureData").Range("iescieGraphChoice").Value = "CIE"
    MsgBox CIE2msg
    Call RunAllFixtures
    If ExitFlag = True Then Exit Sub
    Call outFile("CIE")
    
End If
Application.ScreenUpdating = True

ThisWorkbook.Sheets("Confirmation").Activate
End Sub



Sub ActivateLastSheet()
    Sheets(Sheet20.Range("A2").Value).Activate
End Sub

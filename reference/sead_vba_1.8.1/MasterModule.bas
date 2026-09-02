Attribute VB_Name = "MasterModule"
'This macro was developed by p2w2.  http://p2w2.com/expert-in-microsoft-excel-consultants-consulting/index.php
'Please contact at CS@perceptive-analytics.com in case of any enquiry
Public gbDebug As Boolean


' Macro assigned to button in inputs sheet
Sub masterProc()

'debug flag
gbDebug = False

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

'FLAG this should be updated to use the method in the finalMatrices module, read in from sheet one time here and passed as variable. Implementing within for now.
Set rMaster(1) = wksRoadGeometry.Range("bNumLanes")
Set rMaster(2) = wksRoadGeometry.Range("bLaneWidth")
Set rMaster(3) = wksRoadGeometry.Range("bMedianWidth")
Set rMaster(4) = wksRoadGeometry.Range("bMountingHeight")
Set rMaster(5) = wksRoadGeometry.Range("bPoleSpacing")
Set rMaster(6) = wksRoadGeometry.Range("bPoleSetback")
Set rMaster(7) = wksRoadGeometry.Range("bArmLength")
Set rMaster(8) = wksRoadGeometry.Range("bFixtureArrangement")

Set rMaster(9) = wksRoadGeometry.Range("uNumLanes")
Set rMaster(10) = wksRoadGeometry.Range("uLaneWidth")
Set rMaster(11) = wksRoadGeometry.Range("uMedianWidth")
Set rMaster(12) = wksRoadGeometry.Range("uMountingHeight")
Set rMaster(13) = wksRoadGeometry.Range("uPoleSpacing")
Set rMaster(14) = wksRoadGeometry.Range("uPoleSetback")
Set rMaster(15) = wksRoadGeometry.Range("uArmLength")
Set rMaster(16) = wksRoadGeometry.Range("uFixtureArrangement")
                
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


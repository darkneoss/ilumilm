Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
'Private rngLastLink As Range
'
'Private Sub Workbook_SheetFollowHyperlink(ByVal Sh As Object, ByVal Target As Hyperlink)
'    If UCase(Target.Parent.Value) = "BACK" Then
'        If rngLastLink Is Nothing Then
'            Application.EnableEvents = False
'            Target.Follow
'            Application.EnableEvents = True
'        Else
'            rngLastLink.Worksheet.Activate
'            rngLastLink.Activate
'        End If
'    Else
'        Set rngLastLink = Target.Parent
'    End If
'End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
'    CurrCalc = Application.Calculation
'    Application.Calculation = xlCalculationManual
    Sheet20.Range("A2").Value = Sheet20.Range("A1").Value
    Sheet20.Range("A1").Value = ActiveSheet.Name
'    Application.Calculation = CurrCalc
End Sub
Private Sub Workbook_Open()
    'wksMacroWarning.Visible = xlSheetHidden
    Sheet15.Activate
    
    
    Expiration = Sheet19.Range("ToolExpiration")
    ReleaseDate = Sheet19.Range("ReleaseDate")
    CurrentVersion = Sheet19.Range("VersionNumber")
    WebAddress = "www.superefficient.org"
    today = Now()
    If Expiration < today Then
'message box version
'        box = MsgBox("You are using the SEAD Street Lighting Toolkit version " & CurrentVersion & ", which was released on " & _
'        ReleaseDate & ". A more current version of the Toolkit may be available. Please consult " & WebAddress & _
'        " to download the latest version.", vbOKOnly, "Version out of date")

'form version
        VMessage = "You are using the SEAD Street Lighting Toolkit version " & CurrentVersion & ", which was released on " & _
            ReleaseDate & ". A more current version of the Toolkit may be available. Please visit the SEAD website " & _
            " to download the latest version."
            
        With VersionPopUp
            .Label1.Caption = VMessage
        End With
        DoEvents
        
        VersionPopUp.Show

    End If
    
'Application.ScreenUpdating = False
'Sheets("Input").Visible = False
'Sheets("Dashboard").Visible = False
'Sheets("Illuminance").Visible = False
'Sheets("Luminance").Visible = False
'Sheets("Annual Energy").Visible = False
'Sheets("Simple Payback").Visible = False
'Sheets("Net Present Value").Visible = False
'Sheets("IRR").Visible = False
'Application.ScreenUpdating = True

End Sub

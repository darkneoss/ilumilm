Attribute VB_Name = "VersionPopUp"
Attribute VB_Base = "0{33AD9674-167D-4DC4-BE62-59FAC9A9E257}{E24A1B41-331B-4C4B-9B18-2C38ACF5BEE1}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub Label2_Click()
    Link = "http://superefficient.org/en/Activities/Procurement/SEAD%20Street%20Lighting%20Evaluation%20Toolkit.aspx"
    On Error GoTo NoCanDo
    ActiveWorkbook.FollowHyperlink Address:=Link, NewWindow:=True
    Unload Me
    Exit Sub
NoCanDo:
    MsgBox "Cannot open " & Link
End Sub

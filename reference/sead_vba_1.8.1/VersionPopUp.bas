Attribute VB_Name = "VersionPopUp"
Attribute VB_Base = "0{AF90A54A-ED47-4FA6-B8A3-38A2F3C48CEA}{A680E484-AABF-4C18-82C0-F40BD8BD6360}"
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

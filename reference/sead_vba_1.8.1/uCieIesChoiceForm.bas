Attribute VB_Name = "uCieIesChoiceForm"
Attribute VB_Base = "0{051FA7BF-26B9-408B-B3F8-06D1361DDB69}{33C7D49C-A364-40E1-BDC8-CB587C7F73AD}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub CIE_Click()
Dim choice As String
choice = "CIE"
Unload Me
upgradePlot (choice)
End Sub

Private Sub IES_Click()
Dim choice As String
choice = "IES"
Unload Me
upgradePlot (choice)

End Sub


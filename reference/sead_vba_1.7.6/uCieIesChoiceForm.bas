Attribute VB_Name = "uCieIesChoiceForm"
Attribute VB_Base = "0{9A94FFE8-D30E-45FD-9EFA-0749209A86C9}{844641B8-E2AF-4BBD-BB37-680F098FE049}"
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


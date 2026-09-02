Attribute VB_Name = "bCieIesChoiceForm"
Attribute VB_Base = "0{2CF3482A-B9F1-46C9-8258-D470E4BF2C97}{7DBB65F0-E834-4CD8-AB41-579BCD7BA4F4}"
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
baselinePlot (choice)
End Sub

Private Sub IES_Click()
Dim choice As String
choice = "IES"
Unload Me
baselinePlot (choice)

End Sub



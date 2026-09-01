Attribute VB_Name = "bCieIesChoiceForm"
Attribute VB_Base = "0{D653556D-5E21-4B67-8D26-4E2011D775B7}{2ADC562D-B92F-4922-90D0-E7AD39E2511D}"
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



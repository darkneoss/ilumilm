Attribute VB_Name = "wksRoadGeometry"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
'Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'
'
''Sub infoboxtest()
'
'Dim strTitle As String
'Dim strMsg As String
'Dim lDVType As Long
'Dim sTemp As Shape
'Dim ws As Worksheet
'Dim InfoBoxName As String
'
'
'Application.EnableEvents = False
'
''Hide all callout boxes
'For Each sTemp In ActiveSheet.Shapes
'    If sTemp.Type = 2 Then sTemp.Visible = msoFalse
'Next sTemp
'
'
''Replace code so that each time a cell is clicked, the macro looks up the appropriate
''info box name from the lookup table sheet, and activates the name based on that.
'
''!!!!!!!!!!Change Range name for other sheets!!!!!!!!!!!!
'    For Each c In Sheet9.Range("RoadGeometryInfoBoxes").Cells
'        lookuprow = c.row
'
'        testcolumn = Sheet9.Cells(lookuprow, 2)
'        testrow = Sheet9.Cells(lookuprow, 3)
'
'        If Selection.row = testrow And Selection.column = testcolumn Then
'            InfoBoxName = c.Value
'            Set sTemp = ActiveSheet.Shapes(InfoBoxName)
'            sTemp.Visible = msoTrue
'        End If
'    Next
'
'On Error Resume Next
'On Error GoTo errHandler
'
'errHandler:
'  Application.EnableEvents = True
'  Exit Sub
'
'
'End Sub
'
'Private Sub worksheet_change(ByVal Target As Range)
'    ActiveSheet.Calculate
'End Sub
'
'Private Sub worksheet_activate()
'    ActiveSheet.Calculate
'End Sub

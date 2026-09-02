Attribute VB_Name = "Module3"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range("B3").Select
    Selection.ShowDependents
    ActiveCell.NavigateArrow TowardPrecedent:=False, arrowNumber:=1, _
        linkNumber:=1
End Sub

Attribute VB_Name = "Module1"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Sheets("Chart Data Baseline").Select
    Sheets("Chart Data Baseline").Copy
    ActiveWorkbook.SaveAs filename:= _
        "C:\Users\Perceptive Analytics\Documents\Book2.xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    Range("BW13").Select
End Sub

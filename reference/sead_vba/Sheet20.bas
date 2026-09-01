Attribute VB_Name = "Sheet20"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Application.ScreenUpdating = False
    ' Clear the color of all the cells
    Cells.Interior.ColorIndex = 0
    ' Highlight the active cell
    Target.Interior.Color = RGB(124, 150, 178)
    Application.GoTo ActiveCell, True
    Application.ScreenUpdating = True
End Sub

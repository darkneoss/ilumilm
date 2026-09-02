Attribute VB_Name = "IProgressBar"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'
' Description:  Class to define the IProgressBar interface
'
' Authors:      Rob Bovey, www.appspro.com
'               Stephen Bullen, www.oaltd.co.uk
'
Option Explicit


'Set and get the title
Public Property Let Title(ByVal sNew As String)
End Property

Public Property Get Title() As String
End Property

'Set and get the descriptive text
Public Property Let Text(ByVal sNew As String)
End Property

Public Property Get Text() As String
End Property

'Set and get the minimum value for the bar
Public Property Let Min(ByVal dNew As Double)
End Property

Public Property Get Min() As Double
End Property

'Set and get the maximum value for the bar
Public Property Let max(ByVal dNew As Double)
End Property

Public Property Get max() As Double
End Property

'Set and get the progress point
Public Property Let Progress(ByVal dNew As Double)
End Property

Public Property Get Progress() As Double
End Property

'Show the progress bar
Public Sub Show()
End Sub

'Hide the progress bar
Public Sub Hide()
End Sub







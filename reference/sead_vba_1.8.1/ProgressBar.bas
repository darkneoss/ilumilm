Attribute VB_Name = "ProgressBar"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments: Show a progress form that has the custom IProgressBar interface.
'           Take a look at the intellisense list for the pbProgBar variable.
'
' Arguments:    None
'
' Date          Developer       Action
' --------------------------------------------------------------
' 08 Jun 08     Stephen Bullen  Created
'
Sub ProgressFormInterface()

    'Using the IProgressBar interface,
    'the Intellisense list only shows the
    'properties and methods we choose to
    'expose, making it much easier to use.
    Dim pbProgBar As IProgressBar

    Dim lCounter As Long
    Dim lPause As Long

    Set pbProgBar = New FProgressBarIFace

    pbProgBar.Title = "Professional Excel Development"
    pbProgBar.Text = "Preparing report, please wait..."
    pbProgBar.Min = 0
    pbProgBar.max = 1000
    pbProgBar.Progress = 0
    pbProgBar.Show

    For lCounter = 0 To 1000
        For lPause = 1 To 100000
        Next lPause

        pbProgBar.Progress = lCounter
    Next lCounter

    pbProgBar.Hide

End Sub

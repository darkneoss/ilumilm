Attribute VB_Name = "UserForm"
Option Explicit

Sub DisplayVersionError()

    Expiration = Sheet19.Range("ToolExpiration")
    ReleaseDate = Sheet19.Range("ReleaseDate")
    CurrentVersion = Sheet19.Range("VersionNumber")
    WebAddress = "www.superefficient.org"
    today = Now()

    Verr1 = Sheet25.Range("Verr1")
    VErr2 = Sheet25.Range("Verr2")
    Verr3 = Sheet25.Range("Verr3")
    Verr4 = Sheet25.Range("Verr4")
    
    VMessage = Verr1 & CurrentVersion & VErr2 & _
        ReleaseDate & Verr3 & _
        Verr4
    
    VersionMessage.Text = VMessage
    
End Sub




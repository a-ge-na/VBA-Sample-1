Attribute VB_Name = "Module1"
Option Explicit

'
'
'
Public Sub testProcedure()
    
    Dim PS      As New ProcessSupport
    
    PS.doIt
    Application.StatusBar = "çÏã∆íÜ"
    
    Application.Wait Now() + TimeValue("00:00:05")

End Sub



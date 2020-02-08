VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessSupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private BackupStatusBar         As Variant
Private BackupScreenUpdating    As Boolean
Private BackupDisplayAlarts     As Boolean

'
'
'
Public Sub doIt(Optional ByVal ScreenUpdating As Boolean = False, Optional ByVal DisplayAlarts = False)

    BackupStatusBar = Application.StatusBar
    BackupScreenUpdating = Application.ScreenUpdating
    BackupDisplayAlarts = Application.DisplayAlerts
    Application.ScreenUpdating = ScreenUpdating
    Application.DisplayAlerts = DisplayAlarts
    
End Sub


'
'
'
Private Sub Class_Terminate()

    If BackupStatusBar = False Then
        Application.StatusBar = False
    Else
        Application.StatusBar = BackupStatusBar
    End If
    Application.DisplayAlerts = BackupDisplayAlarts
    Application.ScreenUpdating = BackupScreenUpdating

End Sub



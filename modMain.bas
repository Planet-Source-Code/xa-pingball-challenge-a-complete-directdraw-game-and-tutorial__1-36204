Attribute VB_Name = "modMain"
Public Sub Main()
    If LCase(Command$) = "settings" Then
        frmSets.Show
    Else
        Load frmMain
    End If
End Sub

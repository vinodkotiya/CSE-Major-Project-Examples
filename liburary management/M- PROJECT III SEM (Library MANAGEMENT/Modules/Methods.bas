Attribute VB_Name = "Methods"
Public fMainForm As frmMain
Public totalforms As Integer


Sub Main()
    Dim fLogin As New frmLogin
    fLogin.Show
End Sub


Public Sub wait(i As Integer)
    If i = 1 Then
        Load frmWait
        frmWait.Show
    Else
        frmWait.Hide
    End If
End Sub

Public Sub tot(x As Integer)
    totalforms = totalforms + x
    If totalforms <= 0 Then
        frmMain.mnuFileClose.Enabled = False
        frmMain.mnuFileCloseAll.Enabled = False
    Else
        frmMain.mnuFileClose.Enabled = True
        frmMain.mnuFileCloseAll.Enabled = True
    End If
End Sub

Public Function SetString(temStr As String) As String
    Dim temp As String
    temp = Replace(temStr, "'", "~")
    SetString = Replace(temp, "''", "#")
End Function

Public Function GetString(tempStr As String) As String
    Dim temp As String
    temp = Replace(tempStr, "~", "'")
    GetString = Replace(temp, "#", "''")
End Function



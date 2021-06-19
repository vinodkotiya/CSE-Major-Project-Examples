Attribute VB_Name = "Module11"
Public db As Connection
'Public Rs As New ADODB.Recordset
Public rs_staff As ADODB.Recordset

Public Sub con()
Set db = New Connection
db.Open "PROVIDER=Microsoft.jet.OLEDB.4.0;Data Source= " & App.Path & "/bank.mdb;"
'rs_staff.Open "select * from tran", db, adOpenDynamic, adLockOptimistic

End Sub

Public Function move_in_records(Rs As Recordset, Direction As String, cmdfirst As CommandButton, cmdnext As CommandButton, cmdprevious As CommandButton, cmdlast As CommandButton)

If Rs.BOF And Rs.EOF Then
cmdfirst.Enabled = False
cmdprevious.Enabled = False
cmdnext.Enabled = False
cmdlast.Enabled = False
End If
Select Case Direction

Case "movefirst":
Rs.MoveFirst
cmdfirst.Enabled = False
cmdprevious.Enabled = False
cmdnext.Enabled = True
cmdlast.Enabled = True

Case "movenext":
If Not Rs.EOF Then
Rs.MoveNext
cmdfirst.Enabled = True
cmdprevious.Enabled = True
cmdnext.Enabled = True
cmdlast.Enabled = True
End If
If Rs.EOF Then
Rs.MoveLast
cmdfirst.Enabled = True
cmdprevious.Enabled = True
cmdnext.Enabled = False
cmdlast.Enabled = False
End If

Case "moveprevious":
If Not Rs.BOF Then
Rs.MovePrevious
cmdfirst.Enabled = True
cmdprevious.Enabled = True
cmdnext.Enabled = True
cmdlast.Enabled = True
End If
If Rs.BOF Then
Rs.MoveFirst
cmdfirst.Enabled = False
cmdprevious.Enabled = False
cmdnext.Enabled = True
cmdlast.Enabled = True
End If

Case "movelast":
Rs.MoveLast
cmdfirst.Enabled = True
cmdprevious.Enabled = True
cmdnext.Enabled = False
cmdlast.Enabled = False

End Select

End Function



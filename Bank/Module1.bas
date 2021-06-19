Attribute VB_Name = "Module1"
Public db As Connection
Public Rs As ADODB.Recordset

Public Sub con()
Set db = New Connection
db.Open "PROVIDER=Microsoft.jet.OLEDB.4.0;Data Source= " & App.Path & "/bank.mdb;"
End Sub


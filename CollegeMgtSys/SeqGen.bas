Attribute VB_Name = "SeqGen"
'Variable Declaration

Public RecSet1 As New ADODB.Recordset
Public RecSet2 As New ADODB.Recordset
Public Conns As New ADODB.Connection

'Function To Generate Sequence No

Public Function SeqGen(CodeLength As Integer) As String
    Select Case CodeLength
        Case 1:
            SeqGen = "0000"
        Case 2:
            SeqGen = "000"
        Case 3:
            SeqGen = "00"
        Case 4:
            SeqGen = "0"
        Case 5:
            SeqGen = ""
End Function

'Procedure To Connect Database.

Public Sub ConnConnect()
    
    With Conns
        .Open "uid=college;pwd=college;dsn=oracle"
        .CursorLocation = adUseClient
    End With
    
End Sub

Public Sub dataConnect()
    
    With RecSet1
        .ActiveConnection = Conns
        .CursorType = adOpenDynamic
        .LockType = adLockBatchOptimistic
    End With
    
End Sub

'Procedure To Disconnect Connection

Public Sub DisConnect()
    'Conns.Close
End Sub
'procedure to hide and show the buttons


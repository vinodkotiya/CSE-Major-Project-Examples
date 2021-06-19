Attribute VB_Name = "Module1"
Option Explicit
Public age As Integer
Public sql1 As String
Public sql As String
Public Uname As String
Public Pass As String
Public cn As ADODB.Connection
Public rs As ADODB.Recordset
Public rs1 As ADODB.Recordset
Public p As Integer, q As Integer
Public r As Integer, s As Integer
Public pos1(10) As Integer
Public pos2(10) As Integer
Public mpos1(10) As Integer
Public mpos2(10) As Integer
Public ch(10) As String
Public chm(10) As String

'Procedure for accepting only number into a text box
Public Sub NumberOnly(arg1 As Integer)
Select Case arg1
Case 0 To 7:
arg1 = 0
Case 9 To 45:
arg1 = 0
arg1 = 0
Case 58 To 255:
arg1 = 0
End Select
End Sub
'Procedure for converting string into proper text form
Public Function Initcap(arg2 As String)
arg2 = Trim$(StrConv(arg2, vbProperCase))
Initcap = arg2
End Function
'Procedure for clculating age from date of birth
Public Function AgeCalc(arg3 As Date)
age = Round(DateDiff("d", arg3, Now) / 365)
AgeCalc = age
End Function
'Procedure for connecting to database
Public Sub Connect()
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
cn.Open "DSN=store"
End Sub
'Procedure for disconnecting to database
Public Sub Disconnect()
cn.Close
'rs.Close
Set cn = Nothing
Set rs = Nothing
Set rs1 = Nothing
End Sub

'Procedure for accepting only number into a text box
Public Sub DateOnly(arg1 As Integer)
Select Case arg1
Case 0 To 7:
arg1 = 0
Case 9 To 46:
arg1 = 0
Case 58 To 255:
arg1 = 0
End Select
End Sub

Public Function GetCollect(arg4 As String, arg5 As String)
p = 1
q = 1
pos1(p) = 0
pos2(q) = pos1(p) + 1
While (pos2(q) > pos1(p))
pos1(p) = InStr(pos2(q), arg4, arg5)
If pos1(p) = 0 Then
pos1(p) = pos2(q) + 1
Else
p = p + 1
q = q + 1
pos2(q) = pos1(p - 1) + 1
End If
Wend
For p = 1 To q - 1
ch(p) = Mid(arg4, pos1(p - 1) + 1, pos1(p) - pos1(p - 1) - 1)
Next p
GetCollect = ch
End Function


Public Function GetMedCollect(arg6 As String, arg7 As String)
r = 1
s = 1
mpos1(r) = 0
mpos2(s) = mpos1(r) + 1
While (mpos2(s) > mpos1(r))
mpos1(r) = InStr(mpos2(s), arg6, arg7)
If mpos1(r) = 0 Then
mpos1(r) = mpos2(s) + 1
Else
r = r + 1
s = s + 1
mpos2(s) = mpos1(r - 1) + 1
End If
Wend
For r = 1 To s - 1
chm(r) = Mid(arg6, mpos1(r - 1) + 1, mpos1(r) - mpos1(r - 1) - 1)
Next r
GetMedCollect = chm
End Function


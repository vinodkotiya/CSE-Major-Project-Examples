VERSION 5.00
Begin VB.Form changpassward 
   BackColor       =   &H00FFC0C0&
   Caption         =   "CHANG PASSWARD"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8085
   Icon            =   "changpassward.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   8085
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H000000FF&
      Caption         =   "EXIT"
      Height          =   495
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   1815
   End
   Begin VB.TextBox txtans 
      Height          =   495
      Left            =   3840
      TabIndex        =   10
      Top             =   3720
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txtque 
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   3000
      Width           =   3135
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H000000FF&
      Caption         =   "OK"
      Height          =   495
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox txtrepass 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3840
      PasswordChar    =   "$"
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox txtnewpass 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3840
      PasswordChar    =   "$"
      TabIndex        =   1
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txtpass 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3840
      PasswordChar    =   "$"
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "HINT ANS"
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Label lblque 
      BackColor       =   &H00FFC0C0&
      Caption         =   "HINT QUESTON"
      Height          =   495
      Left            =   1320
      TabIndex        =   7
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label lblrepass 
      BackColor       =   &H00FFC0C0&
      Caption         =   "CONFIRM NEW PASSWARD"
      Height          =   495
      Left            =   1440
      TabIndex        =   5
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label lblnewpass 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ENTER NEW PASSWARD"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label lblpass 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ENTER PASSWARD"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "changpassward"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub CMDEXIT_Click()
If rs.State = 1 Then rs.Close
If cn.State = 1 Then cn.Close
Unload Me
End
End Sub

Private Sub cmdok_Click()
If rs!passward = "" & txtpass.Text Then
If txtnewpass = txtrepass Then
rs!passward = "" & txtnewpass
rs!hintquestion = "" & txtque
rs!ans = "" & txtans
rs.UpdateBatch
Else
MsgBox "PLEAS  RE  ENTER THE NEW PASSWARD", vbCritical, "INFORMATION"
End If
Else
MsgBox "ENTER CORRECT PASSWARD", vbCritical, "WARNING"
End If
End Sub

Private Sub Form_Load()
cn.CursorLocation = adUseClient
cn.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\agrawal.mdb;Persist Security Info=False"
If rs.State = 1 Then rs.Close
rs.Open "select  *  from log_passward", cn, adOpenDynamic, adLockOptimistic
txtque = rs!hintquestion
txtans = rs!ans

End Sub

Private Sub Form_Unload(Cancel As Integer)
If rs.State = 1 Then rs.Close
If cn.State = 1 Then cn.Close
End Sub

Private Sub txtpass_Change()
If rs!passward = txtpass Then
txtans.Visible = True
End If
End Sub

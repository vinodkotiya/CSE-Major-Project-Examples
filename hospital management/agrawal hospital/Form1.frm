VERSION 5.00
Begin VB.Form login 
   BackColor       =   &H00C0E0FF&
   Caption         =   "AGRAWAL HOSPITAL'S INFORMATION SYSTEM"
   ClientHeight    =   8595
   ClientLeft      =   2280
   ClientTop       =   570
   ClientWidth     =   11205
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11205
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   5115
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1995
      Width           =   2295
   End
   Begin VB.CommandButton cmdforgot 
      Caption         =   "FORGOT PASSWARD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5235
      TabIndex        =   4
      Top             =   5745
      Width           =   1455
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2355
      TabIndex        =   3
      Top             =   5745
      Width           =   1455
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   111
      Left            =   11280
      Top             =   6360
   End
   Begin VB.CommandButton cmdchange 
      Caption         =   "CHANGE PASSWARD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   1
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "ENTER  PASSWARD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2115
      TabIndex        =   6
      Top             =   2025
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "THIS SOFTWARE IS DEVELOPED BY  NIKHIL , VIJAY,  ANKUR, VISHAL, AMZAD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   8730
      Width           =   7335
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub cmdchange_Click()
Load changpassward
changpassward.Show
End Sub

Private Sub cmdforgot_Click()
Load frmhint
frmhint.Show

End Sub

Private Sub cmdok_Click()
If Text1.Text = rs!passward Then
Unload Me
Load background
background.Show
Else
MsgBox "ENTER CORRECT PASSWARD", vbCritical, "INFORMATION!!!!!!!!"
Exit Sub
End If
If rs.State = 1 Then rs.Close
If cn.State = 1 Then cn.Close
End Sub

Private Sub CMDEXIT_Click()
If rs.State = 1 Then rs.Close
If cn.State = 1 Then cn.Close
End
End Sub

Private Sub Form_Load()
cn.CursorLocation = adUseClient
cn.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\agrawal.mdb;Persist Security Info=False"
If rs.State = 1 Then rs.Close
rs.Open "select  *  from log_passward", cn, adOpenDynamic, adLockOptimistic
End Sub

Private Sub Form_Unload(Cancel As Integer)
If rs.State = 1 Then rs.Close
If cn.State = 1 Then cn.Close
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdok_Click
End If
End Sub

Private Sub Timer1_Timer()
Label2.Left = 1000
End Sub


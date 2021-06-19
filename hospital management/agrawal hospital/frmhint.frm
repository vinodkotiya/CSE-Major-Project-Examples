VERSION 5.00
Begin VB.Form frmhint 
   BackColor       =   &H00FFC0C0&
   Caption         =   "HINT"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7905
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   7905
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      BackColor       =   &H000000FF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtans 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   1800
      Width           =   5535
   End
   Begin VB.TextBox txtque 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   960
      Width           =   5535
   End
   Begin VB.Label lblans 
      BackColor       =   &H00FFC0C0&
      Caption         =   "ENTER THE ANS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lblque 
      BackColor       =   &H00FFC0C0&
      Caption         =   "           HINT QUESTION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "frmhint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn1 As New ADODB.Connection
Dim rs1 As New ADODB.Recordset


Private Sub cmdok_Click()
If rs1!ans = txtans Then

If rs1.State = 1 Then rs1.Close
If cn1.State = 1 Then cn1.Close
Unload Me
Load background
background.Show

End If

End Sub

Private Sub Form_Load()
cn1.CursorLocation = adUseClient
cn1.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\agrawal.mdb;Persist Security Info=False"
If rs1.State = 1 Then rs1.Close
rs1.Open "select  *  from log_passward", cn1, adOpenDynamic, adLockOptimistic
txtque = rs1!hintquestion

End Sub

Private Sub Form_Unload(Cancel As Integer)
If rs1.State = 1 Then rs1.Close
If cn1.State = 1 Then cn1.Close
End Sub

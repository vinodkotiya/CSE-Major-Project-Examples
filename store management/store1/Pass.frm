VERSION 5.00
Begin VB.Form Form14 
   BackColor       =   &H00404040&
   Caption         =   "LOGIN FORM"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form14"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   3960
      TabIndex        =   2
      Top             =   2400
      Width           =   4575
      Begin VB.TextBox txtUname 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         TabIndex        =   0
         Text            =   " "
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton cmdOk 
         BackColor       =   &H00FFFF00&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1800
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFF00&
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtPass 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "$"
         TabIndex        =   1
         Text            =   " "
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   5
      Height          =   3255
      Left            =   3840
      Top             =   2280
      Width           =   4815
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Uname As String, Pass As String

Private Sub cmdCancel_Click()

Unload Me
End Sub
Private Sub cmdOk_Click()
If Trim(txtUname) = "" And Trim(txtPass) = "" Then
MsgBox "You Must Enter Username & Password", vbExclamation, "Warning"
txtUname.SetFocus
ElseIf Trim(txtUname) = "" Then
MsgBox "You Must Enter Username", vbExclamation, "Warning"
txtUname.SetFocus
ElseIf Trim(txtPass) = "" Then
MsgBox "You Must Enter Password", vbExclamation, "Warning"
txtPass.SetFocus
Else
Set rs = cn.Execute("select password from Login where username = '" & Trim(txtUname) & "' ")
If rs.EOF Then
MsgBox "Wrong Username", vbExclamation, "Warning"
Else
Set rs = cn.Execute("select password from Login where password = '" & Trim(txtPass) & "' ")
If rs.EOF Then
MsgBox "Wrong Password", vbExclamation, "Warning"
Else
Unload Me
frmSplash.Show
End If
End If
End If
End Sub

Private Sub Form_Load()
Call Connect
txtUname = ""
txtPass = ""
End Sub

Private Sub txtUname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
txtPass.SetFocus
End If
End Sub
Private Sub txtPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdOk_Click
End If
End Sub



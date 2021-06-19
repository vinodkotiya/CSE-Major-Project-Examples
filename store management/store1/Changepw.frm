VERSION 5.00
Begin VB.Form Form20 
   Caption         =   "Form20"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6090
   LinkTopic       =   "Form20"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   6090
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCancel 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   5
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox txtPass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   6600
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   " "
      Top             =   4080
      Width           =   2535
   End
   Begin VB.TextBox txtUname 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   6600
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   " "
      Top             =   3480
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   4
      Height          =   3255
      Left            =   3240
      Top             =   2520
      Width           =   6255
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   3240
      X2              =   9480
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ChangeYour Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   4320
      TabIndex        =   6
      Top             =   2640
      Width           =   4215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your New Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   3840
      TabIndex        =   1
      Top             =   4080
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Your Old Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3840
      TabIndex        =   0
      Top             =   3480
      Width           =   2655
   End
End
Attribute VB_Name = "Form20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Set rs = cn.Execute("select username from Login where password = '" & Trim(txtUname) & "'  ")
Set rs1 = cn.Execute("select username from Login where password = '" & Trim(txtPass) & "'  ")
If Trim(txtUname) = "" And Trim(txtPass) = "" Then
MsgBox "Enter Old & New Password", vbInformation, "Change Password"
txtUname.SetFocus
 Else
  If Trim(txtUname) = "" Then
  MsgBox "Enter Old Password", vbInformation, "Change Password"
  txtUname.SetFocus
  Else
    If Trim(txtPass) = "" Then
    MsgBox "Enter New Pssword", vbInformation, "Change Password"
    txtPass.SetFocus
    Else
     If rs.EOF Then
     MsgBox "Old Password Not Correct", vbInformation, "Change Password"
     Else
        If Not rs1.EOF Then
        MsgBox "New Password Not Accepted", vbInformation, "Change Password"
        Else
         cn.Execute ("update login set password = '" & Trim(txtPass) & "' where password = '" & Trim(txtUname) & "'")
         MsgBox "Password Changed", vbInformation, "Change Password"
         txtUname = ""
         txtPass = ""
     End If
   End If
  End If
End If
End If
End Sub

Private Sub Form_Load()
Call Connect
txtPass = ""
txtUname = ""
End Sub

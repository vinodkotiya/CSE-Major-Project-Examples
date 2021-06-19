VERSION 5.00
Begin VB.Form Form21 
   Caption         =   "Form21"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6075
   LinkTopic       =   "Form21"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   6075
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtUname 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   6960
      TabIndex        =   0
      Text            =   " "
      Top             =   4200
      Width           =   2535
   End
   Begin VB.TextBox txtPass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   6960
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   " "
      Top             =   4800
      Width           =   2535
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
      Left            =   4920
      TabIndex        =   2
      Top             =   5520
      Width           =   1455
   End
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
      Left            =   7440
      TabIndex        =   3
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   4
      Height          =   3255
      Left            =   4080
      Top             =   3120
      Width           =   6255
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   4080
      X2              =   10320
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Create New User"
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
      Left            =   5640
      TabIndex        =   6
      Top             =   3240
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Username"
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
      Left            =   4680
      TabIndex        =   5
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Password"
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
      Left            =   4680
      TabIndex        =   4
      Top             =   4800
      Width           =   2175
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Set rs = cn.Execute("select username from Login where password = '" & txtPass & "'  ")
If Trim(txtUname) = "" And Trim(txtPass) = "" Then
MsgBox "Enter Username and Pssword", vbInformation, "New User"
txtUname.SetFocus
 Else
  If Trim(txtUname) = "" Then
  MsgBox "Enter Username", vbInformation, "New User"
  txtUname.SetFocus
  Else
    If Trim(txtPass) = "" Then
    MsgBox "Enter Pssword", vbInformation, "New User"
    txtPass.SetFocus
    Else
     If Not rs.EOF Then
     MsgBox "Password not accepted", vbInformation, "New User"
     Else
      cn.Execute ("insert into login values('" & txtUname & "','" & txtPass & "')")
      MsgBox "Grant Succeded", vbInformation, "New User"
      txtUname = ""
      txtPass = ""
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

VERSION 5.00
Begin VB.Form frmmodify 
   Caption         =   "Modify Account Info"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancle 
      Caption         =   "Cancle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   1800
      Width           =   1695
   End
   Begin VB.CommandButton cmdmodify 
      Caption         =   "Modify"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtaccno 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Account No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frmmodify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancle_Click()
Unload Me
End Sub

Private Sub cmdmodify_Click()
With newacc
.Show
.SetFocus
.txtAcc_No.Locked = True
.txtAddress.Locked = False
.txtBalance.Locked = False
.txtname.Enabled = True
.txtAddress.Enabled = True
'.cmdmodifyacc.Visible = True
.CommandOk.Enabled = False
'.cmdmodifyacc.Enabled = True
.CommandCancle.Enabled = True
.txtAcc_No.Text = txtaccno.Text
End With
End Sub



VERSION 5.00
Begin VB.Form frmedit 
   Caption         =   "Edit Form"
   ClientHeight    =   4080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   5520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CommandQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   2
      Top             =   2760
      Width           =   2655
   End
   Begin VB.CommandButton CommandClosAcc 
      Caption         =   "Close Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   1800
      Width           =   2655
   End
   Begin VB.CommandButton CommandModAcc 
      Caption         =   "Modify Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   5040
      X2              =   5040
      Y1              =   240
      Y2              =   3600
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   480
      X2              =   480
      Y1              =   240
      Y2              =   3600
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   480
      X2              =   5040
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   480
      X2              =   5040
      Y1              =   240
      Y2              =   240
   End
End
Attribute VB_Name = "frmedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandClosAcc_Click(Index As Integer)
frmcloseacc.Show
End Sub

Private Sub CommandModAcc_Click(Index As Integer)
With newacc
.Show
.SetFocus
.txtAcc_No.Locked = True
.txtBalance.Locked = True
.txtName.Enabled = False
.txtAddress.Enabled = False
.cmdsave.Enabled = False
.CommandOk.Visible = False
.CommandOk.Enabled = False
  .cmdedit.Enabled = True
  .cmdfirst.Enabled = True
  .cmdnext.Enabled = True
  .cmdlast.Enabled = True
  .cmdprevious.Enabled = True
End With

End Sub

Private Sub CommandModTran_Click()
With frmTranDaily
.Show
.SetFocus
.txtaccno.Locked = True
.txttrantype.Locked = True
.txttranamount.Enabled = False
.txtpar.Enabled = False
.txttrantype.Enabled = False
.txttranmode.Enabled = False
.txtchno.Enabled = False
.cmbtranmode.Enabled = False
.cmbtrantype.Enabled = False
.DTPicker1.Enabled = False
.cmdsave.Visible = True
.cmdsave.Enabled = False
.CommandOk.Visible = False
.CommandOk.Enabled = False
.CommandCancle.Enabled = False
.Commandback.Visible = True
.Commandback.Enabled = True
.cmdedit.Visible = True

  .cmdfirst.Visible = True
  .cmdnext.Visible = True
  .cmdlast.Visible = True
  .cmdprevious.Visible = True
  .cmdedit.Enabled = False
  .cmdfirst.Enabled = True
  .cmdnext.Enabled = True
  .cmdlast.Enabled = True
  .cmdprevious.Enabled = True
End With
End Sub

Private Sub CommandQuit_Click(Index As Integer)
Unload Me
End Sub

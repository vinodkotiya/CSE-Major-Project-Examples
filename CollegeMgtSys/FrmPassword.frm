VERSION 5.00
Begin VB.Form FrmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Verification Module"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "FrmPassword.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4680
   Begin VB.Frame Frame1 
      Caption         =   "User Verification"
      Height          =   2460
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   4695
      Begin VB.Frame Frame2 
         Height          =   705
         Left            =   90
         TabIndex        =   5
         Top             =   1635
         Width           =   4515
         Begin VB.CommandButton CmdAction 
            Caption         =   "&Edit"
            Height          =   345
            Index           =   1
            Left            =   1890
            TabIndex        =   8
            ToolTipText     =   "Click To Edit Value"
            Top             =   210
            Width           =   810
         End
         Begin VB.CommandButton CmdAction 
            Cancel          =   -1  'True
            Caption         =   "Ca&ncel"
            Height          =   345
            Index           =   2
            Left            =   3450
            TabIndex        =   7
            ToolTipText     =   "Click To Cancel"
            Top             =   210
            Width           =   810
         End
         Begin VB.CommandButton CmdAction 
            Caption         =   "&Ok"
            Default         =   -1  'True
            Height          =   345
            Index           =   0
            Left            =   240
            TabIndex        =   6
            ToolTipText     =   "Click To Countinue"
            Top             =   195
            Width           =   810
         End
      End
      Begin VB.TextBox TxtUserDetails 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1350
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1110
         Width           =   2595
      End
      Begin VB.TextBox TxtUserDetails 
         Height          =   300
         Index           =   0
         Left            =   1365
         MaxLength       =   15
         TabIndex        =   3
         Top             =   600
         Width           =   2595
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         Height          =   195
         Index           =   1
         Left            =   375
         TabIndex        =   2
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   750
         Width           =   795
      End
   End
End
Attribute VB_Name = "FrmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    FrmPassword.Height = 2865
    FrmPassword.Width = 4770
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

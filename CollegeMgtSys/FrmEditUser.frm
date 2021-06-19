VERSION 5.00
Begin VB.Form FrmEditUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Edit Module"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "FrmEditUser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   4680
   Begin VB.Frame Frame1 
      Height          =   2865
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1560
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1770
         Width           =   2820
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1560
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1405
         Width           =   2820
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1560
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1040
         Width           =   2820
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   7
         Top             =   675
         Width           =   2820
      End
      Begin VB.Frame Frame2 
         Height          =   585
         Left            =   195
         TabIndex        =   6
         Top             =   2160
         Width           =   4215
         Begin VB.CommandButton CmdOperation 
            Caption         =   "Ca&ncel"
            Height          =   285
            Index           =   1
            Left            =   2745
            TabIndex        =   12
            Top             =   195
            Width           =   1005
         End
         Begin VB.CommandButton CmdOperation 
            Caption         =   "&Ok"
            Height          =   285
            Index           =   0
            Left            =   495
            TabIndex        =   11
            Top             =   195
            Width           =   1005
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Confirm Password"
         Height          =   195
         Index           =   3
         Left            =   105
         TabIndex        =   5
         Top             =   1860
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "New Password"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   4
         Top             =   1500
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Old Password"
         Height          =   195
         Index           =   1
         Left            =   390
         TabIndex        =   3
         Top             =   1140
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "User Name"
         Height          =   195
         Index           =   0
         Left            =   570
         TabIndex        =   2
         Top             =   780
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "Edit User Name/Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   195
         Width           =   2700
      End
   End
End
Attribute VB_Name = "FrmEditUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdOperation_Click(Index As Integer)
    Select Case Index
        Case 0: 'ok is pressed
        Case 1: 'Cancel is pressed
            Unload FrmEditUser
    End Select
End Sub

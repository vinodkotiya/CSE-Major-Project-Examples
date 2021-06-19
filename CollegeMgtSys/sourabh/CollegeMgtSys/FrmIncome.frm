VERSION 5.00
Begin VB.Form FrmIncome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Incomming Amount Module"
   ClientHeight    =   3990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6570
   Icon            =   "FrmIncome.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   6570
   Begin VB.CommandButton CmdOperation 
      Caption         =   "&Modify"
      Height          =   300
      Index           =   1
      Left            =   3510
      TabIndex        =   14
      Top             =   3390
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "Incomming Amount Module"
      Height          =   4020
      Left            =   15
      TabIndex        =   0
      Top             =   -30
      Width           =   6570
      Begin VB.TextBox TxtDetails 
         Height          =   285
         Index           =   4
         Left            =   2085
         TabIndex        =   12
         Top             =   2550
         Width           =   4335
      End
      Begin VB.TextBox TxtDetails 
         Height          =   285
         Index           =   3
         Left            =   2085
         TabIndex        =   11
         Top             =   2190
         Width           =   1575
      End
      Begin VB.TextBox TxtDetails 
         Height          =   600
         Index           =   2
         Left            =   2100
         TabIndex        =   10
         Top             =   1515
         Width           =   4335
      End
      Begin VB.TextBox TxtDetails 
         Height          =   285
         Index           =   1
         Left            =   2085
         TabIndex        =   9
         Top             =   1185
         Width           =   1575
      End
      Begin VB.TextBox TxtDetails 
         Height          =   285
         Index           =   0
         Left            =   2085
         TabIndex        =   8
         Top             =   810
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Caption         =   "Opeartions"
         Height          =   825
         Index           =   2
         Left            =   150
         TabIndex        =   7
         Top             =   3090
         Width           =   6285
         Begin VB.CommandButton CmdOperation 
            Caption         =   "Ca&ncel"
            Height          =   300
            Index           =   2
            Left            =   5055
            TabIndex        =   15
            Top             =   315
            Width           =   915
         End
         Begin VB.CommandButton CmdOperation 
            Caption         =   "&Ok"
            Height          =   300
            Index           =   0
            Left            =   1605
            TabIndex        =   13
            Top             =   330
            Width           =   915
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Recieved By"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   6
         Top             =   2640
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Recieving Date"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   5
         Top             =   2280
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Source"
         Height          =   195
         Index           =   2
         Left            =   765
         TabIndex        =   4
         Top             =   1905
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Amout"
         Height          =   195
         Index           =   1
         Left            =   825
         TabIndex        =   3
         Top             =   1230
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Incomming No"
         Height          =   195
         Index           =   0
         Left            =   255
         TabIndex        =   2
         Top             =   870
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Incomming Amout Module"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   360
         Left            =   1590
         TabIndex        =   1
         Top             =   240
         Width           =   3345
      End
   End
End
Attribute VB_Name = "FrmIncome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

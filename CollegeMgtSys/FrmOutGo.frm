VERSION 5.00
Begin VB.Form FrmOutGo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OutGoing Amount  Module"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6930
   Icon            =   "FrmOutGo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6930
   Begin VB.Frame Frame1 
      Caption         =   "Out Going Module"
      Height          =   4740
      Left            =   -15
      TabIndex        =   0
      Top             =   -30
      Width           =   6930
      Begin VB.TextBox TxtOutGoing 
         Height          =   285
         Index           =   6
         Left            =   1575
         TabIndex        =   16
         Top             =   3330
         Width           =   4170
      End
      Begin VB.TextBox TxtOutGoing 
         Height          =   285
         Index           =   5
         Left            =   1575
         TabIndex        =   15
         Top             =   2985
         Width           =   4170
      End
      Begin VB.TextBox TxtOutGoing 
         Height          =   285
         Index           =   4
         Left            =   1575
         TabIndex        =   14
         Top             =   2640
         Width           =   2010
      End
      Begin VB.TextBox TxtOutGoing 
         Height          =   750
         Index           =   3
         Left            =   1575
         TabIndex        =   13
         Top             =   1836
         Width           =   4170
      End
      Begin VB.TextBox TxtOutGoing 
         Height          =   285
         Index           =   2
         Left            =   1575
         TabIndex        =   12
         Top             =   1499
         Width           =   4170
      End
      Begin VB.TextBox TxtOutGoing 
         Height          =   285
         Index           =   1
         Left            =   1575
         TabIndex        =   11
         Top             =   1162
         Width           =   2010
      End
      Begin VB.TextBox TxtOutGoing 
         Height          =   285
         Index           =   0
         Left            =   1575
         TabIndex        =   10
         Top             =   825
         Width           =   2010
      End
      Begin VB.Frame Frame2 
         Caption         =   "Operations"
         Height          =   930
         Left            =   165
         TabIndex        =   9
         Top             =   3720
         Width           =   6615
         Begin VB.CommandButton CmdOperation 
            Caption         =   "&Ok"
            Height          =   285
            Index           =   0
            Left            =   525
            TabIndex        =   19
            Top             =   375
            Width           =   915
         End
         Begin VB.CommandButton CmdOperation 
            Caption         =   "Ca&ncel"
            Height          =   285
            Index           =   2
            Left            =   4650
            TabIndex        =   18
            Top             =   390
            Width           =   915
         End
         Begin VB.CommandButton CmdOperation 
            Caption         =   "&Modify"
            Height          =   285
            Index           =   1
            Left            =   2580
            TabIndex        =   17
            Top             =   420
            Width           =   915
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ordered By"
         Height          =   195
         Index           =   7
         Left            =   600
         TabIndex        =   8
         Top             =   3420
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Issued By"
         Height          =   195
         Index           =   6
         Left            =   705
         TabIndex        =   7
         Top             =   3090
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Issue Date"
         Height          =   195
         Index           =   5
         Left            =   630
         TabIndex        =   6
         Top             =   2730
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Purpose"
         Height          =   195
         Index           =   4
         Left            =   795
         TabIndex        =   5
         Top             =   2400
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Issued To"
         Height          =   195
         Index           =   3
         Left            =   690
         TabIndex        =   4
         Top             =   1620
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Amout"
         Height          =   195
         Index           =   2
         Left            =   945
         TabIndex        =   3
         Top             =   1260
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Out Going No"
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   2
         Top             =   930
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "OutGoing Amount Module"
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
         Index           =   0
         Left            =   1725
         TabIndex        =   1
         Top             =   240
         Width           =   3315
      End
   End
End
Attribute VB_Name = "FrmOutGo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

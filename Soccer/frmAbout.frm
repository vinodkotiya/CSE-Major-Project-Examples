VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2265
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   3270
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   3270
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1652
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3012
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Gagan Sahoo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   372
         Left            =   240
         TabIndex        =   3
         Top             =   1260
         Width           =   2532
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "One-On-One Soccer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   252
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   2532
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   1320
         Picture         =   "frmAbout.frx":030A
         Top             =   720
         Width           =   480
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   270
      Left            =   1080
      TabIndex        =   0
      Top             =   1820
      Width           =   1092
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

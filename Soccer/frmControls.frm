VERSION 5.00
Begin VB.Form frmControls 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "One Player Controls"
   ClientHeight    =   3084
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   3156
   Icon            =   "frmControls.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3084
   ScaleWidth      =   3156
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStartGame 
      Caption         =   "Start Game"
      Height          =   324
      Left            =   960
      TabIndex        =   4
      Top             =   2400
      Width           =   1212
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   2280
      Picture         =   "frmControls.frx":030A
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   300
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   2280
      Picture         =   "frmControls.frx":074C
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   2280
      Picture         =   "frmControls.frx":0B8E
      Stretch         =   -1  'True
      Top             =   840
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   2280
      Picture         =   "frmControls.frx":0FD0
      Stretch         =   -1  'True
      Top             =   360
      Width           =   300
   End
   Begin VB.Label lControls 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Left:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   26
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   528
   End
   Begin VB.Label lControls 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Down:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   25
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Width           =   744
   End
   Begin VB.Label lControls 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Up:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   24
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   420
   End
   Begin VB.Label lControls 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Right:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   23
      Left            =   600
      TabIndex        =   0
      Top             =   1800
      Width           =   684
   End
End
Attribute VB_Name = "frmControls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStartGame_Click()
Unload frmStartUp
Unload frmControls
frmPlayer1.Show
End Sub

Private Sub cmdStartGame_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload frmControls
    frmStartUp.Show
End If
End Sub

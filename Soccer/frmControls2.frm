VERSION 5.00
Begin VB.Form frmControls2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Two Player Controls"
   ClientHeight    =   3696
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   4584
   Icon            =   "frmControls2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3696
   ScaleWidth      =   4584
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdStartGame 
      Caption         =   "Start Game"
      Height          =   324
      Left            =   1680
      TabIndex        =   14
      Top             =   3120
      Width           =   1212
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   3960
      Picture         =   "frmControls2.frx":030A
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   300
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   3960
      Picture         =   "frmControls2.frx":074C
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   3960
      Picture         =   "frmControls2.frx":0B8E
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   3960
      Picture         =   "frmControls2.frx":0FD0
      Stretch         =   -1  'True
      Top             =   960
      Width           =   300
   End
   Begin VB.Label lControls 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   1
      Left            =   720
      TabIndex        =   13
      Top             =   120
      Width           =   972
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
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   1440
      Width           =   744
   End
   Begin VB.Label lControls 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Player 2"
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
      Index           =   2
      Left            =   3000
      TabIndex        =   11
      Top             =   120
      Width           =   972
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
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   3
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   528
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
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   5
      Left            =   240
      TabIndex        =   9
      Top             =   960
      Width           =   420
   End
   Begin VB.Label lControls 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   9
      Left            =   1608
      TabIndex        =   8
      Top             =   960
      Width           =   276
   End
   Begin VB.Label lControls 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   10
      Left            =   1644
      TabIndex        =   7
      Top             =   1440
      Width           =   204
   End
   Begin VB.Label lControls 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   11
      Left            =   1644
      TabIndex        =   6
      Top             =   2400
      Width           =   204
   End
   Begin VB.Label lControls 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   12
      Left            =   1644
      TabIndex        =   5
      Top             =   1920
      Width           =   204
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
      Left            =   2640
      TabIndex        =   4
      Top             =   2400
      Width           =   684
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
      Left            =   2640
      TabIndex        =   3
      Top             =   960
      Width           =   420
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
      Left            =   2640
      TabIndex        =   2
      Top             =   1440
      Width           =   744
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
      Left            =   2640
      TabIndex        =   1
      Top             =   1920
      Width           =   528
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
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   6
      Left            =   240
      TabIndex        =   0
      Top             =   2400
      Width           =   684
   End
End
Attribute VB_Name = "frmControls2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdStartGame_Click()
Unload frmStartUp
Unload frmControls2
frmPlayer2.Show
End Sub

Private Sub cmdStartGame_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload frmControls2
    frmStartUp.Show
End If
End Sub

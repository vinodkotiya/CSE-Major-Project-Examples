VERSION 5.00
Begin VB.Form frmControls2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Play Against Friend"
   ClientHeight    =   5220
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4050
   Icon            =   "frmControls2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   4050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "Allow snakes to go through walls"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   3732
   End
   Begin VB.CommandButton cmdStartGame 
      Caption         =   "Start Game"
      Height          =   324
      Left            =   1440
      TabIndex        =   14
      Top             =   4800
      Width           =   1212
   End
   Begin VB.Label lControls 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmControls2.frx":030A
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1620
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   3816
   End
   Begin VB.Image nSnake 
      Height          =   120
      Index           =   1
      Left            =   600
      Picture         =   "frmControls2.frx":03BE
      Stretch         =   -1  'True
      Top             =   600
      Width           =   120
   End
   Begin VB.Image nSnake 
      Height          =   120
      Index           =   2
      Left            =   720
      Picture         =   "frmControls2.frx":05B0
      Stretch         =   -1  'True
      Top             =   600
      Width           =   120
   End
   Begin VB.Image nSnake 
      Height          =   120
      Index           =   3
      Left            =   840
      Picture         =   "frmControls2.frx":07A2
      Stretch         =   -1  'True
      Top             =   600
      Width           =   120
   End
   Begin VB.Image nSnake 
      Height          =   120
      Index           =   4
      Left            =   960
      Picture         =   "frmControls2.frx":0994
      Stretch         =   -1  'True
      Top             =   600
      Width           =   120
   End
   Begin VB.Image nSnake 
      Height          =   120
      Index           =   5
      Left            =   1080
      Picture         =   "frmControls2.frx":0B86
      Stretch         =   -1  'True
      Top             =   600
      Width           =   120
   End
   Begin VB.Image nSnake 
      Height          =   120
      Index           =   6
      Left            =   1200
      Picture         =   "frmControls2.frx":0D78
      Stretch         =   -1  'True
      Top             =   600
      Width           =   120
   End
   Begin VB.Image nSnake2 
      Height          =   120
      Index           =   6
      Left            =   3480
      Picture         =   "frmControls2.frx":0F6A
      Stretch         =   -1  'True
      Top             =   600
      Width           =   120
   End
   Begin VB.Image nSnake2 
      Height          =   120
      Index           =   5
      Left            =   3360
      Picture         =   "frmControls2.frx":115C
      Stretch         =   -1  'True
      Top             =   600
      Width           =   120
   End
   Begin VB.Image nSnake2 
      Height          =   120
      Index           =   4
      Left            =   3240
      Picture         =   "frmControls2.frx":134E
      Stretch         =   -1  'True
      Top             =   600
      Width           =   120
   End
   Begin VB.Image nSnake2 
      Height          =   120
      Index           =   3
      Left            =   3120
      Picture         =   "frmControls2.frx":1540
      Stretch         =   -1  'True
      Top             =   600
      Width           =   120
   End
   Begin VB.Image nSnake2 
      Height          =   120
      Index           =   2
      Left            =   3000
      Picture         =   "frmControls2.frx":1732
      Stretch         =   -1  'True
      Top             =   600
      Width           =   120
   End
   Begin VB.Image nSnake2 
      Height          =   120
      Index           =   1
      Left            =   2880
      Picture         =   "frmControls2.frx":1924
      Stretch         =   -1  'True
      Top             =   600
      Width           =   120
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   1320
      Picture         =   "frmControls2.frx":1B16
      Stretch         =   -1  'True
      Top             =   2340
      Width           =   300
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   1320
      Picture         =   "frmControls2.frx":1F58
      Stretch         =   -1  'True
      Top             =   1860
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   1320
      Picture         =   "frmControls2.frx":239A
      Stretch         =   -1  'True
      Top             =   1380
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   1320
      Picture         =   "frmControls2.frx":27DC
      Stretch         =   -1  'True
      Top             =   900
      Width           =   300
   End
   Begin VB.Label lControls 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Player 2"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   348
      Index           =   1
      Left            =   2760
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
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   348
      Index           =   4
      Left            =   2400
      TabIndex        =   12
      Top             =   1320
      Width           =   696
   End
   Begin VB.Label lControls 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   348
      Index           =   2
      Left            =   480
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
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   348
      Index           =   3
      Left            =   2400
      TabIndex        =   10
      Top             =   1800
      Width           =   600
   End
   Begin VB.Label lControls 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Up:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   348
      Index           =   5
      Left            =   2400
      TabIndex        =   9
      Top             =   840
      Width           =   420
   End
   Begin VB.Label lControls 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   348
      Index           =   9
      Left            =   3648
      TabIndex        =   8
      Top             =   840
      Width           =   276
   End
   Begin VB.Label lControls 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   348
      Index           =   10
      Left            =   3696
      TabIndex        =   7
      Top             =   1320
      Width           =   180
   End
   Begin VB.Label lControls 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   348
      Index           =   11
      Left            =   3696
      TabIndex        =   6
      Top             =   2280
      Width           =   180
   End
   Begin VB.Label lControls 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   348
      Index           =   12
      Left            =   3684
      TabIndex        =   5
      Top             =   1800
      Width           =   204
   End
   Begin VB.Label lControls 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Right:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   348
      Index           =   23
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label lControls 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Up:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   348
      Index           =   24
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   420
   End
   Begin VB.Label lControls 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Down:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   348
      Index           =   25
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   696
   End
   Begin VB.Label lControls 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Left:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   348
      Index           =   26
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   600
   End
   Begin VB.Label lControls 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Right:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   348
      Index           =   6
      Left            =   2400
      TabIndex        =   0
      Top             =   2280
      Width           =   720
   End
End
Attribute VB_Name = "frmControls2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Starting game with Frined
Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload frmControls2
    frmMain.Show
ElseIf KeyCode = vbKeyReturn Then
    Call cmdStartGame_Click
End If
End Sub

Private Sub cmdStartGame_Click()
Unload frmMain
Unload frmControls2
frmSnakesTwoPlayer.Show
End Sub

Private Sub cmdStartGame_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload frmControls2
    frmMain.Show
End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 0 Then
    No_Walls = False
Else
    No_Walls = True
End If
End Sub

Private Sub Form_Activate()
If No_Walls = False Then
    Check1.Value = 0
Else
    Check1.Value = 1
End If
End Sub

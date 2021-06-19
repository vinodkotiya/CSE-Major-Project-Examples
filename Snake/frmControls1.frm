VERSION 5.00
Begin VB.Form frmControls1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Play Against Computer"
   ClientHeight    =   3420
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   3180
   Icon            =   "frmControls1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   3180
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "No Walls"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   9
      Top             =   2560
      Width           =   1575
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Hard"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   972
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Medium"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Value           =   -1  'True
      Width           =   1092
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Easy"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   972
   End
   Begin VB.CommandButton cmdStartGame 
      Caption         =   "Start Game"
      Height          =   324
      Left            =   960
      TabIndex        =   4
      Top             =   3000
      Width           =   1092
   End
   Begin VB.Label lControls 
      BackColor       =   &H00008000&
      BackStyle       =   0  'Transparent
      Caption         =   "Defeat the computer by making it crash into your tail. Collect food to make your tail grow."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   780
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Top             =   120
      Width           =   3108
   End
   Begin VB.Image Image4 
      Height          =   300
      Left            =   2640
      Picture         =   "frmControls1.frx":030A
      Stretch         =   -1  'True
      Top             =   2580
      Width           =   300
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   2640
      Picture         =   "frmControls1.frx":074C
      Stretch         =   -1  'True
      Top             =   2100
      Width           =   300
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   2640
      Picture         =   "frmControls1.frx":0B8E
      Stretch         =   -1  'True
      Top             =   1620
      Width           =   300
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   2640
      Picture         =   "frmControls1.frx":0FD0
      Stretch         =   -1  'True
      Top             =   1140
      Width           =   300
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
      Left            =   1680
      TabIndex        =   3
      Top             =   2520
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
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
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
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
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
      Left            =   1680
      TabIndex        =   0
      Top             =   2040
      Width           =   600
   End
End
Attribute VB_Name = "frmControls1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 0 Then
    No_Walls = False
Else
    No_Walls = True
End If
End Sub

Private Sub cmdStartGame_Click()
Unload frmMain
Unload frmControls1
frmSnakesTwoPlayer.Show
End Sub

Private Sub cmdStartGame_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload frmControls1
    frmMain.Show
End If
End Sub

Private Sub Form_Activate()
If Controls1_Initialized = False Then
    Controls1_Initialized = True
    AI_Level = 2
    Option1.Value = False
    Option2.Value = True
    Option3.Value = False
Else
    If AI_Level = 1 Then
        Option1.Value = True
        Option2.Value = False
        Option3.Value = False
    ElseIf AI_Level = 2 Then
        Option1.Value = False
        Option2.Value = True
        Option3.Value = False
    ElseIf AI_Level = 3 Then
        Option1.Value = False
        Option2.Value = False
        Option3.Value = True
    End If
End If
If No_Walls = False Then
    Check1.Value = 0
Else
    Check1.Value = 1
End If
End Sub

Private Sub Option1_Click()
AI_Level = 1
Option1.Value = True
Option2.Value = False
Option3.Value = False
End Sub

Private Sub Option2_Click()
AI_Level = 2
Option1.Value = False
Option2.Value = True
Option3.Value = False
End Sub

Private Sub Option3_Click()
AI_Level = 3
Option1.Value = False
Option2.Value = False
Option3.Value = True
End Sub

Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call cmdStartGame_Click
ElseIf KeyCode = vbKeyEscape Then
    Unload frmControls1
    frmMain.Show
End If
End Sub

Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call cmdStartGame_Click
ElseIf KeyCode = vbKeyEscape Then
    Unload frmControls1
    frmMain.Show
End If
End Sub

Private Sub Option3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call cmdStartGame_Click
ElseIf KeyCode = vbKeyEscape Then
    Unload frmControls1
    frmMain.Show
End If
End Sub

Private Sub Check1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call cmdStartGame_Click
ElseIf KeyCode = vbKeyEscape Then
    Unload frmControls1
    frmMain.Show
End If
End Sub


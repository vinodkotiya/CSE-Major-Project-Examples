VERSION 5.00
Begin VB.Form frmPlayer1 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "One-On-One Soccer"
   ClientHeight    =   8304
   ClientLeft      =   48
   ClientTop       =   624
   ClientWidth     =   11904
   Icon            =   "frmSoccerForeverPlayer1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8304
   ScaleWidth      =   11904
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCountdown 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   11040
      Top             =   7440
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "  AI "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   2760
      TabIndex        =   10
      Top             =   480
      Visible         =   0   'False
      Width           =   1212
      Begin VB.Timer SwitchBrainOn 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   720
         Top             =   480
      End
      Begin VB.Timer ComputerBrain 
         Interval        =   1
         Left            =   240
         Top             =   480
      End
   End
   Begin VB.Timer tmrMovePlayers 
      Interval        =   10
      Left            =   7440
      Top             =   2040
   End
   Begin VB.Timer tmrMoveBall 
      Interval        =   40
      Left            =   6960
      Top             =   2040
   End
   Begin VB.Timer tmrBallTap 
      Interval        =   1
      Left            =   6480
      Top             =   2040
   End
   Begin VB.Line Line9 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   11760
      X2              =   11760
      Y1              =   3000
      Y2              =   5280
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   120
      X2              =   120
      Y1              =   3000
      Y2              =   5280
   End
   Begin VB.Label RedGoalLine 
      BackColor       =   &H00000000&
      Height          =   732
      Left            =   240
      TabIndex        =   15
      Top             =   3840
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image Portrait 
      Height          =   864
      Index           =   5
      Left            =   7320
      Picture         =   "frmSoccerForeverPlayer1.frx":030A
      Top             =   6240
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Portrait 
      Height          =   864
      Index           =   4
      Left            =   6480
      Picture         =   "frmSoccerForeverPlayer1.frx":35EC
      Top             =   6240
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Portrait 
      Height          =   864
      Index           =   3
      Left            =   5520
      Picture         =   "frmSoccerForeverPlayer1.frx":68CE
      Top             =   6240
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Portrait 
      Height          =   864
      Index           =   2
      Left            =   4560
      Picture         =   "frmSoccerForeverPlayer1.frx":9BB0
      Top             =   6240
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Image Portrait 
      Height          =   864
      Index           =   1
      Left            =   3720
      Picture         =   "frmSoccerForeverPlayer1.frx":CE92
      Top             =   6240
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label lblDirection 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5640
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label lblPause 
      Caption         =   "False"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   4440
      TabIndex        =   11
      Top             =   1920
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Image Ball 
      Height          =   384
      Left            =   5640
      Picture         =   "frmSoccerForeverPlayer1.frx":10174
      Top             =   3960
      Width           =   384
   End
   Begin VB.Shape Blue 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   10800
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   495
   End
   Begin VB.Shape Red 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   360
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   495
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   5880
      X2              =   5880
      Y1              =   120
      Y2              =   8160
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   120
      X2              =   11760
      Y1              =   8160
      Y2              =   8160
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   120
      X2              =   11760
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   11760
      X2              =   11760
      Y1              =   120
      Y2              =   3000
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   11760
      X2              =   11760
      Y1              =   5280
      Y2              =   8160
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   120
      X2              =   120
      Y1              =   5280
      Y2              =   8160
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   3000
   End
   Begin VB.Label BlueLeft 
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   252
      Left            =   6000
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label BlueRight 
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   252
      Left            =   6960
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label BlueDown 
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   252
      Left            =   6480
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label BlueUp 
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   252
      Left            =   6480
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label RedLeft 
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   4080
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label RedRight 
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   5160
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label RedDown 
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   4680
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label RedUp 
      Caption         =   "Up"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Left            =   4680
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label RedGoals 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1332
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   588
   End
   Begin VB.Label BlueGoals 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1332
      Left            =   11064
      TabIndex        =   1
      Top             =   120
      Width           =   588
   End
   Begin VB.Label lblOpponent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Opponent :"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   240
      TabIndex        =   14
      Top             =   7760
      Width           =   1284
   End
   Begin VB.Image OpponentPortrait 
      Height          =   576
      Left            =   240
      Picture         =   "frmSoccerForeverPlayer1.frx":1047E
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   480
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   1480
      Y1              =   5800
      Y2              =   5800
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   1480
      X2              =   1480
      Y1              =   2436
      Y2              =   5800
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   1480
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   10400
      X2              =   11760
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   10400
      X2              =   10400
      Y1              =   2430
      Y2              =   5800
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   10400
      X2              =   11760
      Y1              =   5800
      Y2              =   5800
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time Remaining : 90"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   9240
      TabIndex        =   13
      Top             =   7764
      Width           =   2412
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   2304
      Left            =   4800
      Shape           =   2  'Oval
      Top             =   3000
      Width           =   2172
   End
   Begin VB.Menu mnuFIle 
      Caption         =   "&File"
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New Game"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuPause 
         Caption         =   "&Pause"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmPlayer1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Ball_x_vel As Integer
Private Ball_y_vel As Integer
Private Opponent_velocity As Integer
Private Opponent_start_method As Integer
      
Public Sub SetPlayerDirection(Player As Object, Direction As String)
Direction = LCase(Direction)
Select Case Direction
    Case Is = "up"
        RedUp.Caption = "Down"
    Case Is = "up-right"
        RedUp.Caption = "Down"
        RedRight.Caption = "Down"
    Case Is = "right"
        RedRight.Caption = "Down"
    Case Is = "down-right"
        RedDown.Caption = "Down"
        RedRight.Caption = "Down"
    Case Is = "down"
        RedDown.Caption = "Down"
    Case Is = "down-left"
        RedDown.Caption = "Down"
        RedLeft.Caption = "Down"
    Case Is = "left"
        RedLeft.Caption = "Down"
    Case Is = "up-left"
        RedLeft.Caption = "Down"
        RedUp.Caption = "Down"
End Select
End Sub
        
Public Sub SetPlayerDirection2(Player As Object, Direction As Integer)
Select Case Direction
    Case Is = 1
        RedUp.Caption = "Down"
    Case Is = 2
        RedUp.Caption = "Down"
        RedRight.Caption = "Down"
    Case Is = 3
        RedRight.Caption = "Down"
    Case Is = 4
        RedDown.Caption = "Down"
        RedRight.Caption = "Down"
    Case Is = 5
        RedDown.Caption = "Down"
    Case Is = 6
        RedDown.Caption = "Down"
        RedLeft.Caption = "Down"
    Case Is = 7
        RedLeft.Caption = "Down"
    Case Is = 8
        RedLeft.Caption = "Down"
        RedUp.Caption = "Down"
End Select
End Sub

Public Sub MovePlayer(Player As Object, Direction As String)
Direction = LCase(Direction)
Select Case Direction
    Case Is = "up"
        Player.Left = Player.Left
        Player.Top = Player.Top - 240
    Case Is = "up-right"
        Player.Left = Player.Left + 240
        Player.Top = Player.Top - 240
    Case Is = "right"
        Player.Left = Player.Left + 240
        Player.Top = Player.Top
    Case Is = "down-right"
        Player.Left = Player.Left + 240
        Player.Top = Player.Top + 240
    Case Is = "down"
        Player.Left = Player.Left
        Player.Top = Player.Top + 240
    Case Is = "down-left"
        Player.Left = Player.Left - 240
        Player.Top = Player.Top + 240
    Case Is = "left"
        Player.Left = Player.Left - 240
        Player.Top = Player.Top
    Case Is = "up-left"
        Player.Left = Player.Left - 240
        Player.Top = Player.Top - 240
End Select
End Sub

Public Sub MovePlayer2(Player As Object, Direction As Integer)
Select Case Direction
    Case Is = 1
        Player.Left = Player.Left
        Player.Top = Player.Top - 240
    Case Is = 2
        Player.Left = Player.Left + 240
        Player.Top = Player.Top - 240
    Case Is = 3
        Player.Left = Player.Left + 240
        Player.Top = Player.Top - 240
    Case Is = 4
        Player.Left = Player.Left + 240
        Player.Top = Player.Top + 240
    Case Is = 5
        Player.Left = Player.Left
        Player.Top = Player.Top + 240
    Case Is = 6
        Player.Left = Player.Left - 240
        Player.Top = Player.Top + 240
    Case Is = 7
        Player.Left = Player.Left - 240
        Player.Top = Player.Top
    Case Is = 8
        Player.Left = Player.Left - 240
        Player.Top = Player.Top - 240
End Select
End Sub

Public Function BallComing(Player As Object) As Boolean
Ball_x_vel = 240 And Ball_y_vel = 240
Select Case GameLibrary.Direction(Ball, Player)
    Case Is = 1
        If Ball_x_vel = 0 And Ball_y_vel = -240 Then
            BallComing = True
            Exit Function
        End If
    Case 2 Or 3 Or 4
        If Ball_x_vel = 240 And Ball_y_vel = -240 Then
            BallComing = True
            Exit Function
        End If
    Case Is = 5
        If Ball_x_vel = 240 And Ball_y_vel = 0 Then
            BallComing = True
            Exit Function
        End If
    Case 6 Or 7 Or 8
        If Ball_x_vel = 240 And Ball_y_vel = 240 Then
            BallComing = True
            Exit Function
        End If
    Case Is = 9
        If Ball_x_vel = 0 And Ball_y_vel = 240 Then
            BallComing = True
            Exit Function
        End If
    Case Is = 10 Or 11 Or 12
        If Ball_x_vel = -240 And Ball_y_vel = 240 Then
            BallComing = True
            Exit Function
        End If
    Case Is = 13
        If Ball_x_vel = -240 And Ball_y_vel = 0 Then
            BallComing = True
            Exit Function
        End If
    Case Is = 14 Or 15 Or 16
        If Ball_x_vel = -240 And Ball_y_vel = -240 Then
            BallComing = True
            Exit Function
        End If
End Select
BallComing = False
End Function

Public Function Section(Player_or_Ball As Object) As Integer
If Player_or_Ball.Top <= 2760 Then
    'Object is in the top part
    Section = 1
    Exit Function
ElseIf Player_or_Ball.Top <= 5040 Then
    'Object is in the goal part
    Section = 2
    Exit Function
Else
    'Object is in the bottom part
    Section = 3
    Exit Function
End If
End Function

Public Sub MovePlayerTowardsBall(Player As Object)
Select Case GameLibrary.Direction(Player, Ball)
    Case Is = 1
        Call SetPlayerDirection2(Player, 1)
    Case Is = 2
        Call SetPlayerDirection2(Player, 2)
    Case Is = 3
        Call SetPlayerDirection2(Player, 2)
    Case Is = 4
        Call SetPlayerDirection2(Player, 2)
    Case Is = 5
        Call SetPlayerDirection2(Player, 3)
    Case Is = 6
        Call SetPlayerDirection2(Player, 4)
    Case Is = 7
        Call SetPlayerDirection2(Player, 4)
    Case Is = 8
        Call SetPlayerDirection2(Player, 4)
    Case Is = 9
        Call SetPlayerDirection2(Player, 5)
    Case Is = 10
        Call SetPlayerDirection2(Player, 6)
    Case Is = 11
        Call SetPlayerDirection2(Player, 6)
    Case Is = 12
        Call SetPlayerDirection2(Player, 6)
    Case Is = 13
        Call SetPlayerDirection2(Player, 7)
    Case Is = 14
        Call SetPlayerDirection2(Player, 8)
    Case Is = 15
        Call SetPlayerDirection2(Player, 8)
    Case Is = 16
        Call SetPlayerDirection2(Player, 8)
    Case Else
        Call SetPlayerDirection2(Player, 3)
End Select
End Sub

Public Sub MoveRedTowardsX(XObject As Object)
Select Case GameLibrary.Direction(Red, XObject)
    Case Is = 1
        Call SetPlayerDirection2(Red, 1)
    Case Is = 2
        Call SetPlayerDirection2(Red, 2)
    Case Is = 3
        Call SetPlayerDirection2(Red, 2)
    Case Is = 4
        Call SetPlayerDirection2(Red, 2)
    Case Is = 5
        Call SetPlayerDirection2(Red, 3)
    Case Is = 6
        Call SetPlayerDirection2(Red, 4)
    Case Is = 7
        Call SetPlayerDirection2(Red, 4)
    Case Is = 8
        Call SetPlayerDirection2(Red, 4)
    Case Is = 9
        Call SetPlayerDirection2(Red, 5)
    Case Is = 10
        Call SetPlayerDirection2(Red, 6)
    Case Is = 11
        Call SetPlayerDirection2(Red, 6)
    Case Is = 12
        Call SetPlayerDirection2(Red, 6)
    Case Is = 13
        Call SetPlayerDirection2(Red, 7)
    Case Is = 14
        Call SetPlayerDirection2(Red, 8)
    Case Is = 15
        Call SetPlayerDirection2(Red, 8)
    Case Is = 16
        Call SetPlayerDirection2(Red, 8)
    Case Else
        Call SetPlayerDirection2(Red, 3)
End Select
End Sub

Private Sub ComputerBrain_Timer()
RedUp.Caption = "Up"
RedDown.Caption = "Up"
RedLeft.Caption = "Up"
RedRight.Caption = "Up"
lblDirection.Caption = GameLibrary.Direction(Red, Ball)
Select Case Opponent
    Case Is = 1
        Call Intelligence1
    Case Is = 2
        Call Intelligence2
    Case Is = 3
        Call Intelligence3
    Case Is = 4
        Call Intelligence4
    Case Is = 5
        Call Intelligence5
End Select
End Sub

Private Sub Form_Activate()
Call mnuNewGame_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Exit Game
If KeyCode = vbKeyEscape Then
    Call mnuExit_Click
End If

'Movements For Red Player
'If Blue.Left <> 360 Then
    If KeyCode = vbKeyLeft Then
        BlueLeft.Caption = "Down"
    End If
'If Blue.Left <> 8760 Then
    If KeyCode = vbKeyRight Then
        BlueRight.Caption = "Down"
    End If
'If Blue.Top <> 240 Then
    If KeyCode = vbKeyUp Then
        BlueUp.Caption = "Down"
    End If
'If Blue.Top <> 6000 Then
    If KeyCode = vbKeyDown Then
        BlueDown.Caption = "Down"
    End If

If KeyCode = vbKeyControl Then
    lblDirection.Visible = True
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'Movements For Blue Player
'If Red.Left <> 360 Then
    If KeyCode = vbKeyA Then
        RedLeft.Caption = "Up"
    End If
'If Red.Left <> 8760 Then
    If KeyCode = vbKeyD Then
        RedRight.Caption = "Up"
    End If
'If Red.Top <> 240 Then
    If KeyCode = vbKeyW Then
        RedUp.Caption = "Up"
    End If
'If Red.Top <> 6000 Then
    If KeyCode = vbKeyS Then
        RedDown.Caption = "Up"
    End If

'Movements For Red Player
'If Blue.Left <> 360 Then
    If KeyCode = vbKeyLeft Then
        BlueLeft.Caption = "Up"
    End If
'If Blue.Left <> 8760 Then
    If KeyCode = vbKeyRight Then
        BlueRight.Caption = "Up"
    End If
'If Blue.Top <> 240 Then
    If KeyCode = vbKeyUp Then
        BlueUp.Caption = "Up"
    End If
'If Blue.Top <> 6000 Then
    If KeyCode = vbKeyDown Then
        BlueDown.Caption = "Up"
    End If

If KeyCode = vbKeyControl Then
    lblDirection.Visible = False
End If
End Sub

Private Sub Form_Load()
With Red
    .Left = 360
    .Top = 3960
End With
With Blue
    .Left = 10920
    .Top = 3960
End With
With Ball
    .Left = 5640
    .Top = 3960
End With
Ball_x_vel = 0
Ball_y_vel = 240
BlueGoals.Caption = 0
RedGoals.Caption = 0
Time_Left = Default_Time
lblTime.Caption = "Time Remaining : " & Default_Time
tmrCountdown.Enabled = True
RedUp.Caption = "Up"
RedDown.Caption = "Up"
RedLeft.Caption = "Up"
RedRight.Caption = "Up"
BlueUp.Caption = "Up"
BlueDown.Caption = "Up"
BlueLeft.Caption = "Up"
BlueRight.Caption = "Up"
End Sub

Private Sub mnuExit_Click()
frmStartUp.Show
Call mnuNewGame_Click
lblPause.Caption = "True"
frmPlayer1.Visible = False
End Sub

Private Sub mnuNewGame_Click()
Randomize
With Red
    .Left = 360
    .Top = 3960
End With
With Blue
    .Left = 10920
    .Top = 3960
End With
With Ball
    .Left = 5640
    .Top = 3960
End With
Ball_x_vel = 0
Ball_y_vel = 240
BlueGoals.Caption = 0
RedGoals.Caption = 0
Time_Left = Default_Time
lblTime.Caption = "Time Remaining : " & Default_Time
tmrCountdown.Enabled = False
tmrCountdown.Interval = 2000
tmrCountdown.Enabled = True
RedUp.Caption = "Up"
RedDown.Caption = "Up"
RedLeft.Caption = "Up"
RedRight.Caption = "Up"
BlueUp.Caption = "Up"
BlueDown.Caption = "Up"
BlueLeft.Caption = "Up"
BlueRight.Caption = "Up"
lblPause.Caption = "False"
Opponent_start_method = Int(Rnd * 3) + 1
End Sub

Private Sub mnuPause_Click()
If mnuPause.Caption = "&Pause" Then
    mnuPause.Caption = "&Continue"
    lblPause.Caption = "True"
Else
    mnuPause.Caption = "&Pause"
    lblPause.Caption = "False"
End If
End Sub

Private Sub SwitchBrainOn_Timer()
ComputerBrain.Enabled = True
SwitchBrainOn.Enabled = False
End Sub

Private Sub tmrBallTap_Timer()
'Ball is Moved by Red Player
With Red
    'Postion A
    If .Top = Ball.Top - 240 And .Left = Ball.Left - 240 Then
        Ball_x_vel = 240
        Ball_y_vel = 240
    End If
    'Postion B
    If .Top = Ball.Top - 240 And .Left = Ball.Left Then
        Ball_x_vel = 0
        Ball_y_vel = 240
    End If
    'Postion C
    If .Top = Ball.Top - 240 And .Left = Ball.Left + 240 Then
        Ball_x_vel = -240
        Ball_y_vel = 240
    End If
    'Postion D
    If .Top = Ball.Top And .Left = Ball.Left - 240 Then
        Ball_x_vel = 240
        Ball_y_vel = 0
    End If
    'Postion F
    If .Top = Ball.Top And .Left = Ball.Left + 240 Then
        Ball_x_vel = -240
        Ball_y_vel = 0
    End If
    'Postion G
    If .Top = Ball.Top + 240 And .Left = Ball.Left - 240 Then
        Ball_x_vel = 240
        Ball_y_vel = -240
    End If
    'Postion H
    If .Top = Ball.Top + 240 And .Left = Ball.Left Then
        Ball_x_vel = 0
        Ball_y_vel = -240
    End If
    'Postion I
    If .Top = Ball.Top + 240 And .Left = Ball.Left + 240 Then
        Ball_x_vel = -240
        Ball_y_vel = -240
    End If
End With

'Ball is Moved by Blue Player
With Blue
    'Postion A
    If .Top = Ball.Top - 240 And .Left = Ball.Left - 240 Then
        Ball_x_vel = 240
        Ball_y_vel = 240
    End If
    'Postion B
    If .Top = Ball.Top - 240 And .Left = Ball.Left Then
        Ball_x_vel = 0
        Ball_y_vel = 240
    End If
    'Postion C
    If .Top = Ball.Top - 240 And .Left = Ball.Left + 240 Then
        Ball_x_vel = -240
        Ball_y_vel = 240
    End If
    'Postion D
    If .Top = Ball.Top And .Left = Ball.Left - 240 Then
        Ball_x_vel = 240
        Ball_y_vel = 0
    End If
    'Postion F
    If .Top = Ball.Top And .Left = Ball.Left + 240 Then
        Ball_x_vel = -240
        Ball_y_vel = 0
    End If
    'Postion G
    If .Top = Ball.Top + 240 And .Left = Ball.Left - 240 Then
        Ball_x_vel = 240
        Ball_y_vel = -240
    End If
    'Postion H
    If .Top = Ball.Top + 240 And .Left = Ball.Left Then
        Ball_x_vel = 0
        Ball_y_vel = -240
    End If
    'Postion I
    If .Top = Ball.Top + 240 And .Left = Ball.Left + 240 Then
        Ball_x_vel = -240
        Ball_y_vel = -240
    End If
End With
End Sub

Private Sub tmrCountdown_Timer()
If lblPause.Caption = "False" Then
    Time_Left = Time_Left - 1
    lblTime.Caption = "Time Remaining : " & Time_Left
    tmrCountdown.Interval = 1000
    
    If Time_Left = 0 Then
        lblPause.Caption = "True"
        If Val(RedGoals.Caption) > Val(BlueGoals.Caption) Then
            MsgBox "You Lost to " & frmSelectPlayer.lblName(Opponent).Caption & " " & Val(RedGoals.Caption) & " - " & Val(BlueGoals.Caption) & " !", vbOKOnly + vbExclamation, "One-On-One Soccer"
        ElseIf Val(RedGoals.Caption) < Val(BlueGoals.Caption) Then
            MsgBox "You Defeated " & frmSelectPlayer.lblName(Opponent).Caption & " " & Val(BlueGoals.Caption) & " - " & Val(RedGoals.Caption) & " !", vbOKOnly + vbExclamation, "One-On-One Soccer"
        Else
            MsgBox "Its a Tie ! " & Val(RedGoals.Caption) & " - " & Val(BlueGoals.Caption), vbOKOnly + vbExclamation, "One-On-One Soccer"
        End If
        lblPause.Caption = "False"
        tmrCountdown.Enabled = False
        Call mnuExit_Click
    End If
End If
End Sub

Private Sub tmrMoveBall_Timer()
'Collision - Ball Rebounds From the Sideline
If Ball.Left <= 120 Then
     If Ball.Top < 3000 Or Ball.Top >= 5160 Then
        Ball_x_vel = Ball_x_vel * -1
    Else
        'Goal For Blue
        BlueGoals.Caption = Val(BlueGoals.Caption) + 1
        With Red
            .Left = 360
            .Top = 3960
        End With
        With Blue
            .Left = 10920
            .Top = 3960
        End With
        With Ball
            .Left = 5640
            .Top = 3960
        End With
        Ball_x_vel = 0
        Ball_y_vel = 240
        Opponent_start_method = Int(Rnd * 2) + 1
    End If
End If
If Ball.Left >= 10920 + 240 Then
    If Ball.Top < 3000 Or Ball.Top >= 5160 Then
        Ball_x_vel = Ball_x_vel * -1
    Else
        'Goal For Red
        RedGoals.Caption = Val(RedGoals.Caption) + 1
        With Red
            .Left = 360
            .Top = 3960
        End With
        With Blue
            .Left = 10920
            .Top = 3960
        End With
        With Ball
            .Left = 5640
            .Top = 3960
        End With
        Ball_x_vel = 0
        Ball_y_vel = 240
        Opponent_start_method = Int(Rnd * 2) + 1
    End If
End If
If Ball.Top <= 240 Then
    Ball_y_vel = Ball_y_vel * -1
End If
If Ball.Top >= 7560 Then
    Ball_y_vel = Ball_y_vel * -1
End If

'Move Ball
If lblPause.Caption = "False" Then
    Ball.Left = Ball.Left + Ball_x_vel
    Ball.Top = Ball.Top + Ball_y_vel
End If
End Sub

Private Sub tmrMovePlayers_Timer()
If lblPause.Caption = "False" Then
    'Move Blue Player
    If Blue.Top <> 360 Then
        If BlueUp.Caption = "Down" Then
            Blue.Top = Blue.Top - 240
        End If
    End If
    If Blue.Top <> 7560 Then
        If BlueDown.Caption = "Down" Then
            Blue.Top = Blue.Top + 240
        End If
    End If
    If Blue.Left <> 360 Then
        If BlueLeft.Caption = "Down" Then
            Blue.Left = Blue.Left - 240
        End If
    End If
    If Blue.Left <> 10920 Then
        If BlueRight.Caption = "Down" Then
            Blue.Left = Blue.Left + 240
        End If
    End If
    'Move Red Player
    Opponent_velocity = 240
    If Red.Top <> 360 Then
        If RedUp.Caption = "Down" Then
            Red.Top = Red.Top - Opponent_velocity
        End If
    End If
    If Red.Top <> 7560 Then
        If RedDown.Caption = "Down" Then
            Red.Top = Red.Top + Opponent_velocity
        End If
    End If
    If Red.Left <> 360 Then
        If RedLeft.Caption = "Down" Then
            Red.Left = Red.Left - Opponent_velocity
        End If
    End If
    If Red.Left <> 10920 Then
        If RedRight.Caption = "Down" Then
            Red.Left = Red.Left + Opponent_velocity
        End If
    End If
End If
End Sub

Public Sub Intelligence1()
Select Case GameLibrary.Direction(Red, Ball)
    Case Is = 0
        RedLeft.Caption = "Down"
    Case Is = 1
        RedLeft.Caption = "Down"
    Case Is = 2
        RedUp.Caption = "Down"
        RedRight.Caption = "Down"
    Case Is = 3
        'RedUp.Caption = "Down"
        'RedRight.Caption = "Down"
    Case Is = 4
        RedDown.Caption = "Down"
    Case Is = 5
        End
        'Special
        If Ball.Top < 3960 Then
            RedUp.Caption = "Down"
        ElseIf Ball.Top > 3960 Then
            RedDown.Caption = "Down"
        Else
            RedLeft.Caption = "Down"
        End If
    Case Is = 6
        RedLeft.Caption = "Down"
        RedDown.Caption = "Down"
    Case Is = 7
        RedRight.Caption = "Down"
        RedLeft.Caption = "Down"
    Case Is = 8
        RedUp.Caption = "Down"
    Case Is = 9
        RedLeft.Caption = "Down"
    Case Is = 10
        RedLeft.Caption = "Down"
    Case Is = 11
        RedLeft.Caption = "Down"
    Case Is = 12
        RedUp.Caption = "Down"
        RedLeft.Caption = "Down"
    Case Is = 13
        'Special
        If Ball.Top < 3960 Then
            RedDown.Caption = "Down"
            RedLeft.Caption = "Down"
        ElseIf Ball.Top > 3960 Then
            RedUp.Caption = "Down"
            RedLeft.Caption = "Down"
        Else
            RedDown.Caption = "Down"
            RedLeft.Caption = "Down"
        End If
    Case Is = 14
        RedRight.Caption = "Down"
    Case Is = 15
        'RedLeft.Caption = "Down"
        'RedUp.Caption = "Down"
    Case Is = 16
        If Red.Left < (frmPlayer1.Width / 3) Then
            RedUp.Caption = "Down"
            RedRight.Caption = "Down"
        Else
            RedUp.Caption = "Down"
            RedLeft.Caption = "Down"
        End If
End Select
End Sub

Public Sub Intelligence2()
If Int(Rnd * 3) <> 1 Then
    Select Case GameLibrary.Direction(Red, Ball)
        Case Is = 0
            RedLeft.Caption = "Down"
        Case Is = 1
            RedLeft.Caption = "Down"
        Case Is = 2
            RedUp.Caption = "Down"
        Case Is = 3
            'RedUp.Caption = "Down"
            'RedRight.Caption = "Down"
        Case Is = 4
            If Red.Top > Line6.Y1 And Red.Top < Line6.Y2 Then
                RedRight.Caption = "Down"
            ElseIf Red.Top < Line6.Y1 Then
                RedRight.Caption = "Down"
            Else
                RedRight.Caption = "Down"
            End If
        Case Is = 5
            End
            'Special
            If Ball.Top < 3960 Then
                RedUp.Caption = "Down"
            ElseIf Ball.Top > 3960 Then
                RedDown.Caption = "Down"
            Else
                RedLeft.Caption = "Down"
            End If
        Case Is = 6
            RedRight.Caption = "Down"
            RedDown.Caption = "Down"
        Case Is = 7
            RedRight.Caption = "Down"
            RedLeft.Caption = "Down"
        Case Is = 8
            RedDown.Caption = "Down"
        Case Is = 9
            RedLeft.Caption = "Down"
        Case Is = 10
            RedLeft.Caption = "Down"
        Case Is = 11
            RedLeft.Caption = "Down"
        Case Is = 12
            RedUp.Caption = "Down"
            RedLeft.Caption = "Down"
        Case Is = 13
            'Special
            If Ball.Top < 3960 Then
                RedDown.Caption = "Down"
                RedLeft.Caption = "Down"
            ElseIf Ball.Top > 3960 Then
                RedUp.Caption = "Down"
                RedLeft.Caption = "Down"
            Else
                RedDown.Caption = "Down"
                RedLeft.Caption = "Down"
            End If
        Case Is = 14
            If Red.Height < (frmPlayer1.Height / 2) Then
                RedDown.Caption = "Down"
            Else
                RedUp.Caption = "Down"
            End If
        Case Is = 15
            RedRight.Caption = "Down"
            RedUp.Caption = "Down"
        Case Is = 16
            If Red.Left < (frmPlayer1.Width / 3) Then
                RedUp.Caption = "Down"
                RedRight.Caption = "Down"
            Else
                RedUp.Caption = "Down"
                RedLeft.Caption = "Down"
            End If
    End Select
End If
End Sub

Public Sub Intelligence3()
If Ball_x_vel = 0 And Ball.Left = 5640 Then
    'This is the start of the game
    If Opponent_start_method = 1 Then
        If Blue.Left <= 9600 Then
            If Red.Left < 5400 Then
                RedRight.Caption = "Down"
            Else
                RedDown.Caption = "Down"
            End If
        End If
    Else
        If Blue.Left <= 7600 Then
            Call MovePlayerTowardsBall(Red)
        End If
    End If
ElseIf Ball.Left <= 600 And Ball.Top <= 600 Then
    Call MoveRedTowardsX(RedGoalLine)
ElseIf Ball.Left >= 11760 - 600 And Ball.Top <= 600 Then
    Call MoveRedTowardsX(RedGoalLine)
ElseIf Ball.Left <= 600 And Ball.Top >= 8100 - 600 Then
    Call MoveRedTowardsX(RedGoalLine)
ElseIf Ball.Left >= 11760 - 720 And Ball.Top >= 8100 - 600 Then
    Call MoveRedTowardsX(RedGoalLine)
Else
    Select Case GameLibrary.Direction(Red, Ball)
        Case Is = 0
            RedLeft.Caption = "Down"
        Case Is = 1
            RedLeft.Caption = "Down"
        Case Is = 2
            If Opponent_start_method <= 2 Then
                RedUp.Caption = "Down"
            End If
        Case Is = 3
            RedUp.Caption = "Down"
            RedRight.Caption = "Down"
        Case Is = 4
                RedRight.Caption = "Down"
            
        Case Is = 5
            'Special
            If Ball.Top < 3960 Then
                RedUp.Caption = "Down"
            ElseIf Ball.Top > 3960 Then
                RedDown.Caption = "Down"
            Else
                RedLeft.Caption = "Down"
            End If
        Case Is = 6
            RedRight.Caption = "Down"
            RedDown.Caption = "Down"
'        Case Is = 7
'            RedRight.Caption = "Down"
'            RedLeft.Caption = "Down"
'        Case Is = 8
'            RedDown.Caption = "Down"
        Case Is = 9
            RedLeft.Caption = "Down"
        Case Is = 10
            RedLeft.Caption = "Down"
        Case Is = 11
            RedLeft.Caption = "Down"
        Case Is = 12
            RedUp.Caption = "Down"
            RedLeft.Caption = "Down"
        Case Is = 13
            'Special
            If Ball.Top < 3960 Then
                RedDown.Caption = "Down"
                RedLeft.Caption = "Down"
            ElseIf Ball.Top > 3960 Then
                RedUp.Caption = "Down"
                RedLeft.Caption = "Down"
            Else
                RedDown.Caption = "Down"
                RedLeft.Caption = "Down"
            End If
'        Case Is = 14
'            If Opponent_start_method = 2 Then
'                If Red.Left < (frmPlayer1.Height / 2) Then
'                    RedLeft.Caption = "Down"
'                    RedDown.Caption = "Down"
'                Else
'                    RedLeft.Caption = "Down"
'                    RedUp.Caption = "Down"
'                End If
'            End If
'        Case Is = 15
'            RedLeft.Caption = "Down"
'            RedUp.Caption = "Down"
        Case Is = 16
            If Red.Left < (frmPlayer1.Width / 3) Then
                RedUp.Caption = "Down"
                RedRight.Caption = "Down"
            Else
                RedUp.Caption = "Down"
                RedLeft.Caption = "Down"
            End If
    End Select
End If
End Sub

Public Sub Intelligence4()
Dim Direction
Direction = GameLibrary.Direction(Red, Ball)
Randomize

If Ball_x_vel = 0 And Ball.Left = 5640 Then
    'This is the start of the game
    If Opponent_start_method = 1 Then
        If Red.Left < 5400 Then
            RedRight.Caption = "Down"
        Else
            RedDown.Caption = "Down"
        End If
    Else
        If Red.Left < 5400 Then
            RedRight.Caption = "Down"
        End If
    End If
ElseIf Ball_x_vel = 0 And Ball.Left <= 360 Then
    'In Your Own Half
    If Section(Ball) = 1 Then
        If Section(Red) = 1 Then
            If Red.Left < 600 Then
                RedRight.Caption = "Down"
            ElseIf Red.Left > 600 Then
                RedLeft.Caption = "Down"
            End If
        Else
            Call MovePlayerTowardsBall(Red)
        End If
    ElseIf Section(Ball) = 2 Then
        RedLeft.Caption = "Down"
        If Ball.Top < Red.Top Then
            RedUp.Caption = "Down"
        Else
           RedDown.Caption = "Down"
        End If
    ElseIf Section(Ball) = 3 Then
        If Section(Red) = 3 Then
            If Red.Left < 600 Then
                RedRight.Caption = "Down"
            ElseIf Red.Left > 600 Then
                RedLeft.Caption = "Down"
            End If
        Else
            Call MovePlayerTowardsBall(Red)
        End If
    Else
        Call MovePlayerTowardsBall(Red)
    End If
ElseIf Ball_x_vel = 0 And Ball.Left >= 11040 And Section(Ball) = 2 And GameLibrary.DistanceBetween(Red, Ball, 0, 0) < 480 Then
    RedLeft.Caption = "Down"
ElseIf Ball_y_vel = 0 Then
    If Ball_x_vel = 240 Then
        'Ball is travelling to our opponent's goals
        If Section(Ball) = 1 Then
            'Top Part
            If Ball.Left >= 6480 And GameLibrary.DistanceBetween(Red, Ball, 0, 0) < 480 Then
                RedDown.Caption = "Down"
            Else
                RedRight.Caption = "Down"
            End If
        ElseIf Section(Ball) = 2 Then
            'Goal Part
            If Direction = 4 Or Direction = 5 Or Direction = 6 Then
                RedRight.Caption = "Down"
            ElseIf Direction = 13 Or Direction = 12 Or Direction = 14 Then
                RedDown.Caption = "Down"
            Else
                RedLeft.Caption = "Down"
                RedUp.Caption = "Down"
            End If
        Else
            'Bottom Part
            If Ball.Left >= 6480 Then
                'RedLeft.Caption = "Down"
                RedUp.Caption = "Down"
            Else
                RedRight.Caption = "Down"
            End If
        End If
    ElseIf Ball_x_vel = -240 Then
        'Ball is travelling to our goal !
        If Red.Left < Ball.Left Then
            RedRight.Caption = "Down"
        Else
            RedLeft.Caption = "Down"
            If Ball.Top > Red.Top + 120 Then
                RedDown.Caption = "Down"
            Else
                RedUp.Caption = "Down"
            End If
        End If
    End If
Else
    Select Case Direction
        Case Is = 0
            RedLeft.Caption = "Down"
        Case Is = 1
            RedLeft.Caption = "Down"
        Case Is = 2
            RedUp.Caption = "Down"
        Case Is = 3
            If BallComing(Red) = True Then
                RedUp.Caption = "Down"
                RedRight.Caption = "Down"
            Else
                RedRight.Caption = "Down"
            End If
        Case Is = 4
            RedRight.Caption = "Down"
        Case Is = 5
            RedRight.Caption = "Down"
        Case Is = 6
            RedRight.Caption = "Down"
            RedDown.Caption = "Down"
        Case Is = 7
            RedRight.Caption = "Down"
            RedLeft.Caption = "Down"
        'Case Is = 8
        '    RedDown.Caption = "Down"
        Case Is = 9
            RedLeft.Caption = "Down"
        Case Is = 10
            RedLeft.Caption = "Down"
        Case Is = 11
            RedLeft.Caption = "Down"
        'Case Is = 12
        '    RedUp.Caption = "Down"
        '    RedLeft.Caption = "Down"
        Case Is = 13
            'Special
            If Ball.Top < 3960 Then
                RedDown.Caption = "Down"
                RedLeft.Caption = "Down"
            ElseIf Ball.Top > 3960 Then
                RedUp.Caption = "Down"
                RedLeft.Caption = "Down"
            Else
                RedDown.Caption = "Down"
                RedLeft.Caption = "Down"
            End If
        Case Is = 14
            If Red.Height < (frmPlayer1.Height / 2) Then
                RedDown.Caption = "Down"
            Else
                RedUp.Caption = "Down"
            End If
        Case Is = 15
            RedLeft.Caption = "Down"
            RedUp.Caption = "Down"
        Case Is = 16
            If Red.Left < (frmPlayer1.Width / 3) Then
                RedUp.Caption = "Down"
                RedRight.Caption = "Down"
            Else
                RedUp.Caption = "Down"
                RedLeft.Caption = "Down"
            End If
    End Select
End If
End Sub

Public Sub Intelligence5_FAILURE_1()
Dim Direction
Direction = GameLibrary.Direction(Red, Ball)
Randomize

If Ball_x_vel = 0 And Ball.Left = 5640 Then
    'This is the start of the game
    If Opponent_start_method <= 2 Then
        If Red.Left < 5400 Then
            RedRight.Caption = "Down"
        Else
            RedDown.Caption = "Down"
        End If
    Else
        Call MovePlayerTowardsBall(Red)
    End If
ElseIf Ball_x_vel = 0 And Ball.Left <= 360 Then
    'In Your Own Half
    If Section(Ball) = 1 Then
        If Section(Red) = 1 Then
            If Red.Left < 600 Then
                RedRight.Caption = "Down"
            ElseIf Red.Left > 600 Then
                RedLeft.Caption = "Down"
            End If
        Else
            Call MovePlayerTowardsBall(Red)
        End If
    ElseIf Section(Ball) = 2 Then
        RedLeft.Caption = "Down"
        If Ball.Top < Red.Top Then
            RedUp.Caption = "Down"
        Else
           RedDown.Caption = "Down"
        End If
    ElseIf Section(Ball) = 3 Then
        If Section(Red) = 3 Then
            If Red.Left < 600 Then
                RedRight.Caption = "Down"
            ElseIf Red.Left > 600 Then
                RedLeft.Caption = "Down"
            End If
        Else
            Call MovePlayerTowardsBall(Red)
        End If
    Else
        Call MovePlayerTowardsBall(Red)
    End If
ElseIf Ball_y_vel = 0 Then
    If GameLibrary.DistanceBetween(Red, Ball, 0, 0) < GameLibrary.DistanceBetween(Blue, Ball, 0, 0) Then
        If Ball_x_vel = 240 Then
            'Ball is travelling to our opponent's goals
            If Ball.Left < Red.Left Then
                If Red.Left > 2400 Then
                    RedLeft.Caption = "Down"
                Else
                    If Section(Ball) = 1 Then
                        If Red.Top < Ball.Top + 240 Then
                            RedDown.Caption = "Down"
                        End If
                        If Red.Top = Ball.Top + 240 Then
                            ComputerBrain.Enabled = False
                            SwitchBrainOn.Interval = 500
                            SwitchBrainOn.Enabled = True
                        End If
                    Else
                        If Red.Top > Ball.Top - 240 Then
                            RedUp.Caption = "Down"
                        End If
                        If Red.Top = Ball.Top - 240 Then
                            ComputerBrain.Enabled = False
                            SwitchBrainOn.Interval = 500
                            SwitchBrainOn.Enabled = True
                        End If
                    End If
                End If
            Else
                If Section(Ball) = 1 Then
                    'Top Part
                    If Ball.Left >= 6480 And GameLibrary.DistanceBetween(Red, Ball, 0, 0) < 480 Then
                        RedDown.Caption = "Down"
                    Else
                        RedRight.Caption = "Down"
                    End If
                ElseIf Section(Ball) = 2 Then
                    'Goal Part
                    If Direction = 4 Or Direction = 5 Or Direction = 6 Then
                        RedRight.Caption = "Down"
                    ElseIf Direction = 13 Or Direction = 12 Or Direction = 14 Then
                        RedDown.Caption = "Down"
                    Else
                        RedLeft.Caption = "Down"
                        RedUp.Caption = "Down"
                    End If
                Else
                    'Bottom Part
                    If Ball.Left >= 6480 Then
                        Call MovePlayerTowardsBall(Red)
                    Else
                        RedRight.Caption = "Down"
                    End If
                End If
            End If
        ElseIf Ball_x_vel = -240 Then
            'Ball is travelling to our goal !
            If Red.Left < Ball.Left Then
                RedRight.Caption = "Down"
            Else
                If Red.Left <= 2040 Then
                    RedLeft.Caption = "Down"
                    If Ball.Top = Red.Top Then
                        RedDown.Caption = "Down"
                    ElseIf Ball.Top > 600 Then
                        RedUp.Caption = "Down"
                    End If
                Else
                    If Red.Left > 2400 Then
                        RedLeft.Caption = "Down"
                    End If
                    'If Ball.Top > Red.Top + 120 Then
                    '    RedDown.Caption = "Down"
                    'Else
                    '    RedUp.Caption = "Down"
                    'End If
                End If
            End If
        End If
    Else
        Call MovePlayerTowardsBall(Red)
    End If
Else
    Select Case Direction
        Case Is = 0
            RedLeft.Caption = "Down"
        Case Is = 1
            RedLeft.Caption = "Down"
        Case Is = 2
            RedUp.Caption = "Down"
        Case Is = 3
            If BallComing(Red) = True Then
                RedUp.Caption = "Down"
                RedRight.Caption = "Down"
            Else
                RedRight.Caption = "Down"
            End If
        Case Is = 4
            RedRight.Caption = "Down"
        Case Is = 5
            RedRight.Caption = "Down"
        Case Is = 6
            RedRight.Caption = "Down"
            RedDown.Caption = "Down"
        Case Is = 7
            RedRight.Caption = "Down"
            RedLeft.Caption = "Down"
        Case Is = 8
            RedDown.Caption = "Down"
        Case Is = 9
            Call MovePlayerTowardsBall(Red)
        Case Is = 10
            Call MovePlayerTowardsBall(Red)
        Case Is = 11
            Call MovePlayerTowardsBall(Red)
        Case Is = 12
            RedUp.Caption = "Down"
            RedLeft.Caption = "Down"
        Case Is = 13
            'Special
            If Ball.Top < 3960 Then
                RedDown.Caption = "Down"
                RedLeft.Caption = "Down"
            ElseIf Ball.Top > 3960 Then
                RedUp.Caption = "Down"
                RedLeft.Caption = "Down"
            Else
                RedDown.Caption = "Down"
                RedLeft.Caption = "Down"
            End If
        Case Is = 14
            If Red.Height < (frmPlayer1.Height / 2) Then
                RedDown.Caption = "Down"
            Else
                RedUp.Caption = "Down"
            End If
        Case Is = 15
            RedLeft.Caption = "Down"
            RedUp.Caption = "Down"
        Case Is = 16
            If Red.Left < (frmPlayer1.Width / 3) Then
                RedUp.Caption = "Down"
                RedRight.Caption = "Down"
            Else
                RedUp.Caption = "Down"
                RedLeft.Caption = "Down"
            End If
    End Select
End If
End Sub

Public Sub Intelligence5_FAILURE_2()
Dim Direction
Direction = GameLibrary.Direction(Red, Ball)
Randomize

If Ball_x_vel = 0 And Ball.Left = 5640 Then
    'This is the start of the game
    If Opponent_start_method <= 2 Then
        If Red.Left < 5400 Then
            RedRight.Caption = "Down"
        Else
            RedDown.Caption = "Down"
        End If
    Else
        Call MovePlayerTowardsBall(Red)
    End If
ElseIf Ball_x_vel = 0 And Ball.Left <= 360 Then
    'In Your Own Half
    If Section(Ball) = 1 Then
        If Section(Red) = 1 Then
            If Red.Left < 600 Then
                RedRight.Caption = "Down"
            ElseIf Red.Left > 600 Then
                RedLeft.Caption = "Down"
            End If
        Else
            Call MovePlayerTowardsBall(Red)
        End If
    ElseIf Section(Ball) = 2 Then
        RedLeft.Caption = "Down"
        If Ball.Top < Red.Top Then
            RedUp.Caption = "Down"
        Else
           RedDown.Caption = "Down"
        End If
    ElseIf Section(Ball) = 3 Then
        If Section(Red) = 3 Then
            If Red.Left < 600 Then
                RedRight.Caption = "Down"
            ElseIf Red.Left > 600 Then
                RedLeft.Caption = "Down"
            End If
        Else
            Call MovePlayerTowardsBall(Red)
        End If
    Else
        Call MovePlayerTowardsBall(Red)
    End If
ElseIf Ball_y_vel = 0 Then
    If Ball_x_vel = 240 Then
        'Ball is travelling to our opponent's goals
        If Section(Ball) = 1 Then
            'Top Part
            If Ball.Left >= 6480 And GameLibrary.DistanceBetween(Red, Ball, 0, 0) < 480 Then
                RedDown.Caption = "Down"
            Else
                RedRight.Caption = "Down"
            End If
        ElseIf Section(Ball) = 2 Then
            'Goal Part
            If Direction = 4 Or Direction = 5 Or Direction = 6 Then
                RedRight.Caption = "Down"
            ElseIf Direction = 13 Or Direction = 12 Or Direction = 14 Then
                RedDown.Caption = "Down"
            Else
                RedLeft.Caption = "Down"
                RedUp.Caption = "Down"
            End If
        Else
            'Bottom Part
            If Ball.Left >= 6480 Then
                Call MovePlayerTowardsBall(Red)
            Else
                RedRight.Caption = "Down"
            End If
        End If
    ElseIf Ball_x_vel = -240 Then
        'Ball is travelling to our goal !
        If Red.Left < Ball.Left Then
            RedRight.Caption = "Down"
        Else
            RedLeft.Caption = "Down"
            If Ball.Top > Red.Top + 120 Then
                RedDown.Caption = "Down"
            Else
                RedUp.Caption = "Down"
            End If
        End If
    End If
Else
    Select Case Direction
        Case Is = 0
            RedLeft.Caption = "Down"
        Case Is = 1
            RedLeft.Caption = "Down"
        Case Is = 2
            RedUp.Caption = "Down"
        Case Is = 3
            If BallComing(Red) = True Then
                RedUp.Caption = "Down"
                RedRight.Caption = "Down"
            Else
                RedRight.Caption = "Down"
            End If
        Case Is = 4
            RedRight.Caption = "Down"
        Case Is = 5
            RedRight.Caption = "Down"
        Case Is = 6
            RedRight.Caption = "Down"
            RedDown.Caption = "Down"
        Case Is = 7
            RedRight.Caption = "Down"
            RedLeft.Caption = "Down"
        Case Is = 8
            RedDown.Caption = "Down"
        Case Is = 9
            Call MovePlayerTowardsBall(Red)
        Case Is = 10
            Call MovePlayerTowardsBall(Red)
        Case Is = 11
            Call MovePlayerTowardsBall(Red)
        Case Is = 12
            RedUp.Caption = "Down"
            RedLeft.Caption = "Down"
        Case Is = 13
            'Special
            If Ball.Top < 3960 Then
                RedDown.Caption = "Down"
                RedLeft.Caption = "Down"
            ElseIf Ball.Top > 3960 Then
                RedUp.Caption = "Down"
                RedLeft.Caption = "Down"
            Else
                RedDown.Caption = "Down"
                RedLeft.Caption = "Down"
            End If
        Case Is = 14
            If Red.Height < (frmPlayer1.Height / 2) Then
                RedDown.Caption = "Down"
            Else
                RedUp.Caption = "Down"
            End If
        Case Is = 15
            RedLeft.Caption = "Down"
            RedUp.Caption = "Down"
        Case Is = 16
            If Red.Left < (frmPlayer1.Width / 3) Then
                RedUp.Caption = "Down"
                RedRight.Caption = "Down"
            Else
                RedUp.Caption = "Down"
                RedLeft.Caption = "Down"
            End If
    End Select
End If
End Sub

Public Sub Intelligence5()
Dim Direction
Direction = GameLibrary.Direction(Red, Ball)
Randomize

If Ball_x_vel = 0 And Ball.Left = 5640 Then
    'This is the start of the game
    If Opponent_start_method <= 2 Then
        If Red.Left < 5400 Then
            RedRight.Caption = "Down"
        Else
            RedDown.Caption = "Down"
        End If
    Else
        Call MovePlayerTowardsBall(Red)
    End If
ElseIf Ball_x_vel = 0 And Ball.Left <= 360 Then
    'In Your Own Half
    If Section(Ball) = 1 Then
        If Section(Red) = 1 Then
            If Red.Left < 600 Then
                RedRight.Caption = "Down"
            ElseIf Red.Left > 600 Then
                RedLeft.Caption = "Down"
            End If
        Else
            Call MovePlayerTowardsBall(Red)
        End If
    ElseIf Section(Ball) = 2 Then
        RedLeft.Caption = "Down"
        If Ball.Top < Red.Top Then
            RedUp.Caption = "Down"
        Else
           RedDown.Caption = "Down"
        End If
    ElseIf Section(Ball) = 3 Then
        If Section(Red) = 3 Then
            If Red.Left < 600 Then
                RedRight.Caption = "Down"
            ElseIf Red.Left > 600 Then
                RedLeft.Caption = "Down"
            End If
        Else
            Call MovePlayerTowardsBall(Red)
        End If
    Else
        Call MovePlayerTowardsBall(Red)
    End If
ElseIf Ball_x_vel = 0 And Ball.Left >= 11040 And Section(Ball) = 2 And GameLibrary.DistanceBetween(Red, Ball, 0, 0) < 480 Then
    RedLeft.Caption = "Down"
ElseIf Ball_y_vel = 0 Then
    If Ball_x_vel = 240 Then
        'Ball is travelling to our opponent's goals
        If Section(Ball) = 1 Then
            'Top Part
            If Direction <> 15 Then
                If Ball.Left >= 6480 And GameLibrary.DistanceBetween(Red, Ball, 0, 0) < 480 Then
                    RedDown.Caption = "Down"
                Else
                    'If Ball.Left > Red.Left And GameLibrary.DistanceBetween(Red, Ball, 0, 0) > 480 Then
                        If Red.Left < 6400 Then
                            RedRight.Caption = "Down"
                        End If
                    'End If
                End If
            End If
        ElseIf Section(Ball) = 2 Then
            'Goal Part
            If Direction = 4 Or Direction = 5 Or Direction = 6 Then
                RedRight.Caption = "Down"
            ElseIf Direction = 13 Or Direction = 12 Or Direction = 14 Then
                RedDown.Caption = "Down"
            Else
                RedLeft.Caption = "Down"
                RedUp.Caption = "Down"
            End If
        Else
            'Bottom Part
            If Ball.Left >= 6480 Then
                Call MovePlayerTowardsBall(Red)
            Else
                RedRight.Caption = "Down"
            End If
        End If
    ElseIf Ball_x_vel = -240 Then
        'Ball is travelling to our goal !
        If Ball.Left < 6400 Then
            If Red.Left < Ball.Left Then
                RedRight.Caption = "Down"
            Else
                RedLeft.Caption = "Down"
                If Ball.Top > Red.Top + 120 Then
                    RedDown.Caption = "Down"
                Else
                    RedUp.Caption = "Down"
                End If
            End If
        Else
            If Red.Left > 7200 Then
                RedLeft.Caption = "Down"
            End If
        End If
    End If
Else
    Select Case Direction
        Case Is = 0
            RedLeft.Caption = "Down"
        Case Is = 1
            RedLeft.Caption = "Down"
        Case Is = 2
            RedUp.Caption = "Down"
        Case Is = 3
            If BallComing(Red) = True Then
                RedUp.Caption = "Down"
                RedRight.Caption = "Down"
            Else
                RedRight.Caption = "Down"
            End If
        Case Is = 4
            RedRight.Caption = "Down"
        Case Is = 5
            RedRight.Caption = "Down"
        Case Is = 6
            RedRight.Caption = "Down"
            RedDown.Caption = "Down"
        Case Is = 7
            RedDown.Caption = "Down"
            RedRight.Caption = "Down"
        Case Is = 8
            RedDown.Caption = "Down"
        Case Is = 9
            Call MovePlayerTowardsBall(Red)
        Case Is = 10
            If Red.Left > 2500 Then
                Call MovePlayerTowardsBall(Red)
            End If
        Case Is = 11
            Call MovePlayerTowardsBall(Red)
        Case Is = 12
            RedUp.Caption = "Down"
            RedLeft.Caption = "Down"
        Case Is = 13
            'Special
            If Ball.Top < 3960 Then
                RedDown.Caption = "Down"
                RedLeft.Caption = "Down"
            ElseIf Ball.Top > 3960 Then
                RedUp.Caption = "Down"
                RedLeft.Caption = "Down"
            Else
                RedDown.Caption = "Down"
                RedLeft.Caption = "Down"
            End If
        Case Is = 14
            Call MovePlayerTowardsBall(Red)
        Case Is = 15
            RedLeft.Caption = "Down"
            RedUp.Caption = "Down"
        Case Is = 16
            RedUp.Caption = "Down"
            RedLeft.Caption = "Down"
    End Select
End If
End Sub


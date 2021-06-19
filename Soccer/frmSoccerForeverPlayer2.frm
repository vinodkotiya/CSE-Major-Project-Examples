VERSION 5.00
Begin VB.Form frmPlayer2 
   Appearance      =   0  'Flat
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "One-On-One Soccer"
   ClientHeight    =   8304
   ClientLeft      =   48
   ClientTop       =   624
   ClientWidth     =   11880
   Icon            =   "frmSoccerForeverPlayer2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8304
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCountdown 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   11040
      Top             =   7440
   End
   Begin VB.Timer tmrMovePlayers 
      Interval        =   10
      Left            =   7800
      Top             =   5040
   End
   Begin VB.Timer tmrMoveBall 
      Interval        =   40
      Left            =   6960
      Top             =   5040
   End
   Begin VB.Timer tmrBallTap 
      Interval        =   1
      Left            =   6480
      Top             =   5040
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   10404
      X2              =   11764
      Y1              =   5800
      Y2              =   5800
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   10404
      X2              =   10404
      Y1              =   2430
      Y2              =   5800
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   10404
      X2              =   11764
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   1480
      Y1              =   2430
      Y2              =   2430
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   1476
      X2              =   1476
      Y1              =   2430
      Y2              =   5800
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   120
      X2              =   1480
      Y1              =   5800
      Y2              =   5800
   End
   Begin VB.Label lblPause 
      BackStyle       =   0  'Transparent
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
      Left            =   4320
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   1092
   End
   Begin VB.Image Ball 
      Height          =   384
      Left            =   5640
      Picture         =   "frmSoccerForeverPlayer2.frx":030A
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
   Begin VB.Line Line9 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   11760
      X2              =   11760
      Y1              =   3000
      Y2              =   5280
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
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      X1              =   120
      X2              =   120
      Y1              =   3000
      Y2              =   5280
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   3000
   End
   Begin VB.Label RedKeyLeft 
      BackStyle       =   0  'Transparent
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
      TabIndex        =   9
      Top             =   840
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label RedKeyRight 
      BackStyle       =   0  'Transparent
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
      Left            =   7440
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label RedKeyDown 
      BackStyle       =   0  'Transparent
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
      TabIndex        =   7
      Top             =   1200
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label RedKeyUp 
      BackStyle       =   0  'Transparent
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
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label BlueKeyLeft 
      BackStyle       =   0  'Transparent
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
      Left            =   3960
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label BlueKeyRight 
      BackStyle       =   0  'Transparent
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
      Left            =   5040
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label BlueKeyDown 
      BackStyle       =   0  'Transparent
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
      Left            =   4560
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   612
   End
   Begin VB.Label BlueKeyUp 
      BackStyle       =   0  'Transparent
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
      Left            =   4560
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
      Left            =   9216
      TabIndex        =   11
      Top             =   7764
      Width           =   2412
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      Height          =   2302
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
Attribute VB_Name = "frmPlayer2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Ball_x_vel As Integer
Private Ball_y_vel As Integer

Public Sub MovePlayer(Player As Object, Direction As String)
Select Case Direction
    Case Is = "Up"
        Player.Left = Player.Left
        Player.Top = Player.Top - 240
    Case Is = "Up-Right"
        Player.Left = Player.Left + 240
        Player.Top = Player.Top - 240
    Case Is = "Right"
        Player.Left = Player.Left + 240
        Player.Top = Player.Top
    Case Is = "Down-Right"
        Player.Left = Player.Left + 240
        Player.Top = Player.Top + 240
    Case Is = "Down"
        Player.Left = Player.Left
        Player.Top = Player.Top + 240
    Case Is = "Down-Left"
        Player.Left = Player.Left - 240
        Player.Top = Player.Top + 240
    Case Is = "Left"
        Player.Left = Player.Left - 240
        Player.Top = Player.Top
    Case Is = "Up-Left"
        Player.Left = Player.Left - 240
        Player.Top = Player.Top - 240
End Select
End Sub

Private Sub ComputerBrain_Timer()
BlueKeyUp.Caption = "Up"
BlueKeyDown.Caption = "Up"
BlueKeyLeft.Caption = "Up"
BlueKeyRight.Caption = "Up"
'If GameLibrary.LengthBetween(Red, Ball, 1, 1) < GameLibrary.LengthBetween(Blue, Ball, 1, 1) Then
    'Blue is closer than Red to the ball
    If GameLibrary.DistanceBetween(Red, Ball, 1, 1) < 2160 Then
        'Blue is close to the Ball
        'If BallComing(Red) = False Then
        
        'Else
        
        Select Case GameLibrary.Direction(Red, Ball)
        Case Is = 0
            BlueKeyLeft.Caption = "Down"
        Case Is = 1
            BlueKeyLeft.Caption = "Down"
        Case Is = 2
            BlueKeyUp.Caption = "Down"
        Case Is = 3
            BlueKeyUp.Caption = "Down"
            BlueKeyRight.Caption = "Down"
        Case Is = 4
            BlueKeyRight.Caption = "Down"
        Case Is = 5
            'Special
            If Ball.Top < 3960 Then
                BlueKeyUp.Caption = "Down"
            ElseIf Ball.Top > 3960 Then
                BlueKeyDown.Caption = "Down"
            Else
                BlueKeyLeft.Caption = "Down"
            End If
        Case Is = 6
            BlueKeyRight.Caption = "Down"
            BlueKeyDown.Caption = "Down"
        Case Is = 7
            BlueKeyRight.Caption = "Down"
            BlueKeyLeft.Caption = "Down"
        Case Is = 8
            BlueKeyDown.Caption = "Down"
        Case Is = 9
            BlueKeyLeft.Caption = "Down"
        Case Is = 10
            BlueKeyLeft.Caption = "Down"
        Case Is = 11
            BlueKeyLeft.Caption = "Down"
        Case Is = 12
            BlueKeyUp.Caption = "Down"
            BlueKeyLeft.Caption = "Down"
        Case Is = 13
            'Special
            If Ball.Top < 3960 Then
                BlueKeyDown.Caption = "Down"
                BlueKeyLeft.Caption = "Down"
            ElseIf Ball.Top > 3960 Then
                BlueKeyUp.Caption = "Down"
                BlueKeyLeft.Caption = "Down"
            Else
                BlueKeyDown.Caption = "Down"
                BlueKeyLeft.Caption = "Down"
            End If
        Case Is = 14
            BlueKeyLeft.Caption = "Down"
        Case Is = 15
            BlueKeyLeft.Caption = "Down"
            BlueKeyUp.Caption = "Down"
        Case Is = 16
            BlueKeyUp.Caption = "Down"
            BlueKeyLeft.Caption = "Down"
        End Select
        'End If
    Else
    'Blue is not close to the Ball
    Select Case GameLibrary.Direction(Red, Ball)
        Case Is = 1
            BlueKeyUp.Caption = "Down"
        Case Is = 2
            BlueKeyUp.Caption = "Down"
            BlueKeyRight.Caption = "Down"
        Case Is = 3
            BlueKeyUp.Caption = "Down"
            BlueKeyRight.Caption = "Down"
        Case Is = 4
            BlueKeyUp.Caption = "Down"
            BlueKeyRight.Caption = "Down"
        Case Is = 5
            BlueKeyRight.Caption = "Down"
        Case Is = 6
            BlueKeyDown.Caption = "Down"
            BlueKeyRight.Caption = "Down"
        Case Is = 7
            BlueKeyDown.Caption = "Down"
            BlueKeyRight.Caption = "Down"
        Case Is = 8
            BlueKeyDown.Caption = "Down"
            BlueKeyRight.Caption = "Down"
        Case Is = 9
            BlueKeyDown.Caption = "Down"
        Case Is = 10
            BlueKeyDown.Caption = "Down"
            BlueKeyLeft.Caption = "Down"
        Case Is = 11
            BlueKeyDown.Caption = "Down"
            BlueKeyLeft.Caption = "Down"
        Case Is = 12
            BlueKeyDown.Caption = "Down"
            BlueKeyLeft.Caption = "Down"
        Case Is = 13
            BlueKeyDown.Caption = "Down"
        Case Is = 14
            BlueKeyUp.Caption = "Down"
            BlueKeyLeft.Caption = "Down"
        Case Is = 15
            BlueKeyUp.Caption = "Down"
            BlueKeyLeft.Caption = "Down"
        Case Is = 16
            BlueKeyUp.Caption = "Down"
            BlueKeyLeft.Caption = "Down"
    End Select
    End If
'Else
   'Red is closer than Blue to the ball
'End If

'Movements Allowed For Blue Player
'If Red.Left <> 360 Then
'    If KeyCode = vbKeyA Then
'        BlueKeyLeft.Caption = "Down"
'    End If
'If Red.Left <> 8760 Then
'    If KeyCode = vbKeyD Then
'        BlueKeyRight.Caption = "Down"
'    End If
'If Red.Top <> 240 Then
'    If KeyCode = vbKeyW Then
'        BlueKeyUp.Caption = "Down"
'    End If
'If Red.Top <> 6000 Then
'    If KeyCode = vbKeyS Then
'        BlueKeyDown.Caption = "Down"
'    End If
End Sub

Private Sub Form_Activate()
Call mnuNewGame_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'Exit Game
If KeyCode = vbKeyEscape Then
    Call mnuExit_Click
End If

'Movements For Blue Player
If KeyCode = vbKeyA Then
    BlueKeyLeft.Caption = "Down"
End If
If KeyCode = vbKeyD Then
    BlueKeyRight.Caption = "Down"
End If
If KeyCode = vbKeyW Then
    BlueKeyUp.Caption = "Down"
End If
If KeyCode = vbKeyS Then
    BlueKeyDown.Caption = "Down"
End If

'Movements For Red Player
If KeyCode = vbKeyLeft Then
    RedKeyLeft.Caption = "Down"
End If
If KeyCode = vbKeyRight Then
    RedKeyRight.Caption = "Down"
End If
If KeyCode = vbKeyUp Then
    RedKeyUp.Caption = "Down"
End If
If KeyCode = vbKeyDown Then
    RedKeyDown.Caption = "Down"
End If

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'Movements For Blue Player
    If KeyCode = vbKeyA Then
        BlueKeyLeft.Caption = "Up"
    End If

    If KeyCode = vbKeyD Then
        BlueKeyRight.Caption = "Up"
    End If

    If KeyCode = vbKeyW Then
        BlueKeyUp.Caption = "Up"
    End If

    If KeyCode = vbKeyS Then
        BlueKeyDown.Caption = "Up"
    End If

'Movements For Red Player
    If KeyCode = vbKeyLeft Then
        RedKeyLeft.Caption = "Up"
    End If

    If KeyCode = vbKeyRight Then
        RedKeyRight.Caption = "Up"
    End If

    If KeyCode = vbKeyUp Then
        RedKeyUp.Caption = "Up"
    End If

    If KeyCode = vbKeyDown Then
        RedKeyDown.Caption = "Up"
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
BlueKeyUp.Caption = "Up"
BlueKeyDown.Caption = "Up"
BlueKeyLeft.Caption = "Up"
BlueKeyRight.Caption = "Up"
RedKeyUp.Caption = "Up"
RedKeyDown.Caption = "Up"
RedKeyLeft.Caption = "Up"
RedKeyRight.Caption = "Up"
End Sub

Private Sub mnuExit_Click()
frmStartUp.Show
Call mnuNewGame_Click
frmPlayer2.Visible = False
lblPause.Caption = "True"
End Sub

Private Sub mnuNewGame_Click()
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
BlueKeyUp.Caption = "Up"
BlueKeyDown.Caption = "Up"
BlueKeyLeft.Caption = "Up"
BlueKeyRight.Caption = "Up"
RedKeyUp.Caption = "Up"
RedKeyDown.Caption = "Up"
RedKeyLeft.Caption = "Up"
RedKeyRight.Caption = "Up"
lblPause.Caption = "False"
End Sub

Private Sub mnuPause_Click()
If Red.Visible = True Then
    If mnuPause.Caption = "&Pause" Then
        mnuPause.Caption = "&Continue"
        lblPause.Caption = "True"
    Else
        mnuPause.Caption = "&Pause"
        lblPause.Caption = "False"
    End If
End If
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
            MsgBox "Red Player Wins ! " & Val(RedGoals.Caption) & " - " & Val(BlueGoals.Caption), vbOKOnly + vbExclamation, "One-On-One Soccer"
        ElseIf Val(RedGoals.Caption) < Val(BlueGoals.Caption) Then
            MsgBox "Blue Player Wins ! " & Val(BlueGoals.Caption) & " - " & Val(RedGoals.Caption), vbOKOnly + vbExclamation, "One-On-One Soccer"
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
        With Ball
            .Top = 3960
            .Left = 5640
        End With
        With Red
            .Top = 3960
            .Left = 360
        End With
        With Blue
            .Top = 3960
            .Left = 10920
        End With
        Ball_x_vel = 0
        Ball_y_vel = 240
    End If
End If
If Ball.Left >= 10920 + 240 Then
    If Ball.Top < 3000 Or Ball.Top >= 5160 Then
        Ball_x_vel = Ball_x_vel * -1
    Else
        'Goal For Red
        RedGoals.Caption = Val(RedGoals.Caption) + 1
        With Ball
             .Top = 3960
             .Left = 5640
        End With
        With Red
             .Top = 3960
             .Left = 360
        End With
        With Blue
            .Top = 3960
            .Left = 10920
        End With
        Ball_x_vel = 0
        Ball_y_vel = 240
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
        If RedKeyUp.Caption = "Down" Then
            Blue.Top = Blue.Top - 240
        End If
    End If
    If Blue.Top <> 7560 Then
        If RedKeyDown.Caption = "Down" Then
            Blue.Top = Blue.Top + 240
        End If
    End If
    If Blue.Left <> 360 Then
        If RedKeyLeft.Caption = "Down" Then
            Blue.Left = Blue.Left - 240
        End If
    End If
    If Blue.Left <> 10920 Then
        If RedKeyRight.Caption = "Down" Then
            Blue.Left = Blue.Left + 240
        End If
    End If
    'Move Red Player
    If Red.Top <> 360 Then
        If BlueKeyUp.Caption = "Down" Then
            Red.Top = Red.Top - 240
        End If
    End If
    If Red.Top <> 7560 Then
        If BlueKeyDown.Caption = "Down" Then
            Red.Top = Red.Top + 240
        End If
    End If
    If Red.Left <> 360 Then
        If BlueKeyLeft.Caption = "Down" Then
            Red.Left = Red.Left - 240
        End If
    End If
    If Red.Left <> 10920 Then
        If BlueKeyRight.Caption = "Down" Then
            Red.Left = Red.Left + 240
        End If
    End If
End If
End Sub

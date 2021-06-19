VERSION 5.00
Begin VB.Form frmSnakesTwoPlayer 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snakes - 2 Player"
   ClientHeight    =   7695
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   5895
   Icon            =   "frmSnakesTwoPlayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTimer 
      Height          =   252
      Left            =   2280
      ScaleHeight     =   195
      ScaleWidth      =   1260
      TabIndex        =   5
      Top             =   7420
      Width           =   1320
      Begin VB.Shape Timer 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   228
         Left            =   0
         Top             =   0
         Width           =   1284
      End
   End
   Begin VB.Timer tmrTimer 
      Interval        =   1000
      Left            =   4080
      Top             =   3360
   End
   Begin VB.Timer tmrMoveSnake 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1320
      Top             =   2280
   End
   Begin VB.Label lblGamePaused 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Game Paused"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2400
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   1116
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level : Medium"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   2364
      TabIndex        =   4
      Top             =   7140
      Width           =   1164
   End
   Begin VB.Shape Shape4 
      FillStyle       =   0  'Solid
      Height          =   132
      Left            =   120
      Top             =   6960
      Width           =   5652
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  'Solid
      Height          =   132
      Left            =   120
      Top             =   0
      Width           =   5652
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   7092
      Left            =   5760
      Top             =   0
      Width           =   132
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   7092
      Left            =   0
      Top             =   0
      Width           =   132
   End
   Begin VB.Label lblPlayer1Lives 
      AutoSize        =   -1  'True
      Caption         =   "Player 1 Lives : 3"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   60
      TabIndex        =   3
      Top             =   7404
      Width           =   1560
   End
   Begin VB.Label lblPlayer2Lives 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Player 2 Lives : 3"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4272
      TabIndex        =   2
      Top             =   7404
      Width           =   1560
   End
   Begin VB.Image imgHead2 
      Height          =   180
      Index           =   3
      Left            =   2040
      Picture         =   "frmSnakesTwoPlayer.frx":030A
      Top             =   3480
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgHead2 
      Height          =   180
      Index           =   2
      Left            =   2280
      Picture         =   "frmSnakesTwoPlayer.frx":04FC
      Top             =   3360
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgHead2 
      Height          =   180
      Index           =   4
      Left            =   1800
      Picture         =   "frmSnakesTwoPlayer.frx":06EE
      Top             =   3360
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgHead2 
      Height          =   180
      Index           =   1
      Left            =   2040
      Picture         =   "frmSnakesTwoPlayer.frx":08E0
      Top             =   3120
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgTail2 
      Height          =   180
      Index           =   2
      Left            =   3120
      Picture         =   "frmSnakesTwoPlayer.frx":0AD2
      Top             =   3360
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgTail2 
      Height          =   180
      Index           =   4
      Left            =   2640
      Picture         =   "frmSnakesTwoPlayer.frx":0CC4
      Top             =   3360
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgTail2 
      Height          =   180
      Index           =   3
      Left            =   2880
      Picture         =   "frmSnakesTwoPlayer.frx":0EB6
      Top             =   3600
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgTail2 
      Height          =   180
      Index           =   1
      Left            =   2880
      Picture         =   "frmSnakesTwoPlayer.frx":10A8
      Top             =   3120
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgBody2 
      Height          =   180
      Left            =   2520
      Picture         =   "frmSnakesTwoPlayer.frx":129A
      Top             =   3720
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image nSnake2 
      Height          =   120
      Index           =   1
      Left            =   2280
      Picture         =   "frmSnakesTwoPlayer.frx":148C
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image nSnake2 
      Height          =   120
      Index           =   2
      Left            =   2400
      Picture         =   "frmSnakesTwoPlayer.frx":167E
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image nSnake2 
      Height          =   120
      Index           =   3
      Left            =   2520
      Picture         =   "frmSnakesTwoPlayer.frx":1870
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image nSnake2 
      Height          =   120
      Index           =   4
      Left            =   2640
      Picture         =   "frmSnakesTwoPlayer.frx":1A62
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image nSnake2 
      Height          =   120
      Index           =   5
      Left            =   2760
      Picture         =   "frmSnakesTwoPlayer.frx":1C54
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image nSnake2 
      Height          =   120
      Index           =   6
      Left            =   2880
      Picture         =   "frmSnakesTwoPlayer.frx":1E46
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label lblPlayer2Score 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Player 2 Score : 0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4188
      TabIndex        =   1
      Top             =   7116
      Width           =   1644
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3360
      Picture         =   "frmSnakesTwoPlayer.frx":2038
      Top             =   2280
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image Food2 
      Appearance      =   0  'Flat
      Height          =   120
      Left            =   1800
      Picture         =   "frmSnakesTwoPlayer.frx":27E6
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   3360
      Picture         =   "frmSnakesTwoPlayer.frx":2F94
      Top             =   2760
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Line Line1 
      X1              =   5880
      X2              =   0
      Y1              =   7090
      Y2              =   7090
   End
   Begin VB.Label lblPlayer1Score 
      AutoSize        =   -1  'True
      Caption         =   "Player 1 Score : 0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   60
      TabIndex        =   0
      Top             =   7116
      Width           =   1644
   End
   Begin VB.Image Snake2 
      Height          =   144
      Index           =   0
      Left            =   2400
      Picture         =   "frmSnakesTwoPlayer.frx":3742
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   144
   End
   Begin VB.Image nSnake 
      Height          =   120
      Index           =   6
      Left            =   2640
      Picture         =   "frmSnakesTwoPlayer.frx":3934
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image Food 
      Appearance      =   0  'Flat
      Height          =   120
      Left            =   1560
      Picture         =   "frmSnakesTwoPlayer.frx":3B26
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   120
   End
   Begin VB.Image imgTail 
      Height          =   180
      Index           =   1
      Left            =   2760
      Picture         =   "frmSnakesTwoPlayer.frx":42D4
      Top             =   1680
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgTail 
      Height          =   180
      Index           =   3
      Left            =   2760
      Picture         =   "frmSnakesTwoPlayer.frx":44C6
      Top             =   2160
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgTail 
      Height          =   180
      Index           =   4
      Left            =   2520
      Picture         =   "frmSnakesTwoPlayer.frx":46B8
      Top             =   1920
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgTail 
      Height          =   180
      Index           =   2
      Left            =   3000
      Picture         =   "frmSnakesTwoPlayer.frx":48AA
      Top             =   1920
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgBody 
      Height          =   180
      Left            =   2400
      Picture         =   "frmSnakesTwoPlayer.frx":4A9C
      Top             =   2280
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgHead 
      Height          =   180
      Index           =   1
      Left            =   1920
      Picture         =   "frmSnakesTwoPlayer.frx":4C8E
      Top             =   1680
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgHead 
      Height          =   180
      Index           =   4
      Left            =   1680
      Picture         =   "frmSnakesTwoPlayer.frx":4E80
      Top             =   1920
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgHead 
      Height          =   180
      Index           =   2
      Left            =   2160
      Picture         =   "frmSnakesTwoPlayer.frx":5072
      Top             =   1920
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgHead 
      Height          =   180
      Index           =   3
      Left            =   1920
      Picture         =   "frmSnakesTwoPlayer.frx":5264
      Top             =   2040
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image nSnake 
      Height          =   120
      Index           =   5
      Left            =   2520
      Picture         =   "frmSnakesTwoPlayer.frx":5456
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image nSnake 
      Height          =   120
      Index           =   4
      Left            =   2400
      Picture         =   "frmSnakesTwoPlayer.frx":5648
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image nSnake 
      Height          =   120
      Index           =   3
      Left            =   2280
      Picture         =   "frmSnakesTwoPlayer.frx":583A
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image nSnake 
      Height          =   120
      Index           =   2
      Left            =   2160
      Picture         =   "frmSnakesTwoPlayer.frx":5A2C
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image nSnake 
      Height          =   120
      Index           =   1
      Left            =   2040
      Picture         =   "frmSnakesTwoPlayer.frx":5C1E
      Stretch         =   -1  'True
      Top             =   2640
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image Snake 
      Height          =   120
      Index           =   0
      Left            =   2280
      Picture         =   "frmSnakesTwoPlayer.frx":5E10
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   120
   End
End
Attribute VB_Name = "frmSnakesTwoPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Two players code module

Option Explicit
Private Direction As Integer
Private Direction2 As Integer
Private Snake_Head As Node
Private Snake_Tail As Node
Private Snake_Body_Count As Integer
Private Snake_Head2 As Node
Private Snake_Tail2 As Node
Private Snake_Body_Count2 As Integer
Private Food_X As Integer
Private Food_Y As Integer
Private Player1_Collect_Food As Integer
Private Player2_Collect_Food As Integer
Const Columns = 48
Const Rows = 58
Private Is_Game_Paused As Boolean
Private Gain_Life1 As Integer
Private Gain_Life2 As Integer

Private Sub Form_Activate()
On Error Resume Next
Randomize
Snake_Body_Count = 0
Snake_Body_Count2 = 0
Direction = 2
Direction2 = 4
Food_X = Int((Rnd * (Columns - 2)) + 1)
Food_Y = Int((Rnd * (Rows - 2)) + 1)
Food.Left = Food_X * 120
Food.Top = Food_Y * 120
Player1_Score = 0
Player2_Score = 0
lblPlayer1Score.Caption = "Player 1 Score : " & Player1_Score
lblPlayer2Score.Caption = "Player 2 Score : " & Player2_Score
Player1_Lives = 5
Player2_Lives = 5
lblPlayer1Lives.Caption = "Player 1 Lives : " & Player1_Lives
lblPlayer2Lives.Caption = "Player 2 Lives : " & Player2_Lives
If AI_Level = 1 Then
    lblLevel.Caption = "Level : Easy"
ElseIf AI_Level = 2 Then
    lblLevel.Caption = "Level : Medium"
ElseIf AI_Level = 3 Then
    lblLevel.Caption = "Level : Hard"
End If
Set Snake_Head = Nothing
Set Snake_Tail = Nothing
Set Snake_Head2 = Nothing
Set Snake_Tail2 = Nothing
Insert_Node_At_Front2 46, 56
Insert_Node_At_Front2 45, 56
Insert_Node_At_Front2 44, 56
Insert_Node_At_Front2 43, 56
Insert_Node_At_Front2 42, 56
Insert_Node_At_Front2 41, 56

Insert_Node_At_Front 3, 3
Insert_Node_At_Front 4, 3
Insert_Node_At_Front 5, 3
Insert_Node_At_Front 6, 3
Insert_Node_At_Front 7, 3
Insert_Node_At_Front 8, 3

tmrMoveSnake.Enabled = True
End Sub

Private Sub picTimer_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyP Then
    If Is_Game_Paused = False Then
        Is_Game_Paused = True
        lblGamePaused.Visible = True
    Else
        Is_Game_Paused = False
        lblGamePaused.Visible = False
    End If
End If
If Is_Game_Paused = True Then Exit Sub
Select Case KeyCode
    Case Is = vbKeyUp
        If Direction <> 3 And Snake_Head.X <> Snake_Head.NextNode.X And Snake_Head.Y - 1 <> Snake_Head.NextNode.Y Then
            Direction = 1
        End If
    Case Is = vbKeyRight
        If Direction <> 4 And Snake_Head.X + 1 <> Snake_Head.NextNode.X And Snake_Head.Y <> Snake_Head.NextNode.Y Then
            Direction = 2
        End If
    Case Is = vbKeyDown
        If Direction <> 1 And Snake_Head.X <> Snake_Head.NextNode.X And Snake_Head.Y + 1 <> Snake_Head.NextNode.Y Then
            Direction = 3
        End If
    Case Is = vbKeyLeft
        If Direction <> 2 And Snake_Head.X - 1 <> Snake_Head.NextNode.X And Snake_Head.Y <> Snake_Head.NextNode.Y Then
            Direction = 4
        End If
    Case Is = vbKeyF
        tmrMoveSnake.Interval = 1
    Case Is = vbKeyD
        tmrMoveSnake.Interval = 10
    Case Is = vbKeyS
        tmrMoveSnake.Interval = 100
    Case Is = vbKeyEscape
        Unload frmSnakesTwoPlayer
        frmMain.Show
End Select

If AI_On = False Then
    Select Case KeyCode
        Case Is = vbKeyW
            If Direction2 <> 3 And Snake_Head2.X <> Snake_Head2.NextNode.X And Snake_Head2.Y - 1 <> Snake_Head2.NextNode.Y Then
                Direction2 = 1
            End If
        Case Is = vbKeyD
            If Direction2 <> 4 And Snake_Head2.X + 1 <> Snake_Head2.NextNode.X And Snake_Head2.Y <> Snake_Head2.NextNode.Y Then
                Direction2 = 2
            End If
        Case Is = vbKeyS
            If Direction2 <> 1 And Snake_Head2.X <> Snake_Head2.NextNode.X And Snake_Head2.Y + 1 <> Snake_Head2.NextNode.Y Then
                Direction2 = 3
            End If
        Case Is = vbKeyA
            If Direction2 <> 2 And Snake_Head2.X + 1 <> Snake_Head2.NextNode.X And Snake_Head2.Y <> Snake_Head2.NextNode.Y Then
                Direction2 = 4
            End If
    End Select
End If
End Sub

Private Sub picTimer_LostFocus()
On Error Resume Next
picTimer.SetFocus
End Sub

Private Sub tmrMoveSnake_Timer()
If Player1_Lives = 0 Or Player2_Lives = 0 Or Is_Game_Paused = True Then
    Exit Sub
End If
Dim temp_node As Node
Dim i As Integer
Dim Occupied(1 To 4) As Boolean
Occupied(1) = False
Occupied(2) = False
Occupied(3) = False
Occupied(4) = False
'Check if snake ate any food
If Snake_Head.X = Food_X And Snake_Head.Y = Food_Y Then
    Food_X = Int((Rnd * (Columns - 2)) + 1)
    Food_Y = Int((Rnd * (Rows - 2)) + 1)
    Food.Left = Food_X * 120
    Food.Top = Food_Y * 120
    For i = 1 To Snake_Body_Count
        If Snake(i).Left = Food.Left And Snake(i).Top = Food.Top Then
            Food_X = Int((Rnd * (Columns - 2)) + 1)
            Food_Y = Int((Rnd * (Rows - 2)) + 1)
            Food.Left = Food_X * 120
            Food.Top = Food_Y * 120
        End If
    Next i
    For i = 1 To Snake_Body_Count2
        If Snake2(i).Left = Food.Left And Snake2(i).Top = Food.Top Then
            Food_X = Int((Rnd * (Columns - 2)) + 1)
            Food_Y = Int((Rnd * (Rows - 2)) + 1)
            Food.Left = Food_X * 120
            Food.Top = Food_Y * 120
        End If
    Next i
    Player1_Score = Player1_Score + Points
    lblPlayer1Score.Caption = "Player 1 Score : " & Player1_Score
    Player1_Collect_Food = 4
    If Player1_Score > Gain_Life1 * 1000 + 1000 Then
        Player1_Lives = Player1_Lives + 1
        Gain_Life1 = Gain_Life1 + 1
        lblPlayer1Lives.Caption = "Player 1 Lives : " & Player1_Lives
    End If
ElseIf Snake_Head2.X = Food_X And Snake_Head2.Y = Food_Y Then
    Food_X = Int((Rnd * (Columns - 2)) + 1)
    Food_Y = Int((Rnd * (Rows - 2)) + 1)
    Food.Left = Food_X * 120
    Food.Top = Food_Y * 120
    For i = 1 To Snake_Body_Count
        If Snake(i).Left = Food.Left And Snake(i).Top = Food.Top Then
            Food_X = Int((Rnd * (Columns - 2)) + 1)
            Food_Y = Int((Rnd * (Rows - 2)) + 1)
            Food.Left = Food_X * 120
            Food.Top = Food_Y * 120
        End If
    Next i
    For i = 1 To Snake_Body_Count2
        If Snake2(i).Left = Food.Left And Snake2(i).Top = Food.Top Then
            Food_X = Int((Rnd * (Columns - 2)) + 1)
            Food_Y = Int((Rnd * (Rows - 2)) + 1)
            Food.Left = Food_X * 120
            Food.Top = Food_Y * 120
        End If
    Next i
    Player2_Score = Player2_Score + Points
    lblPlayer2Score.Caption = "Player 2 Score : " & Player2_Score
    Player2_Collect_Food = 4
    If Player2_Score > Gain_Life2 * 1000 + 1000 Then
        Player2_Lives = Player2_Lives + 1
        Gain_Life2 = Gain_Life2 + 1
        lblPlayer2Lives.Caption = "Player 2 Lives : " & Player2_Lives
    End If
    'New Artifical Intelligence
    If AI_On = True And AI_Level > 1 Then
        Select Case GameLibrary.Direction(Snake2(Snake_Head2.Index), Food)
            Case Is = 1
                If Direction2 <> 3 Then
                    Direction2 = 1
                Else
                    If Snake_Head.X > Food_X Then
                        Direction2 = 2
                    Else
                        Direction2 = 4
                    End If
                End If
            Case Is <= 8
                If Direction2 <> 4 Then
                    Direction2 = 2
                Else
                    If Snake_Head.Y > Food_Y Then
                        Direction2 = 3
                    Else
                        Direction2 = 1
                    End If
                End If
            Case Is = 9
                If Direction2 <> 1 Then
                    Direction2 = 3
                Else
                    If Snake_Head.X > Food_X Then
                        Direction2 = 2
                    Else
                        Direction2 = 4
                    End If
                End If
            Case Else
                If Direction2 <> 2 Then
                    Direction2 = 4
                Else
                    If Snake_Head.Y > Food_Y Then
                        Direction2 = 3
                    Else
                        Direction2 = 1
                    End If
                End If
        End Select
    End If
End If

'Collision Detection
If No_Walls = False Then
    If Snake_Head.X <= 0 Or Snake_Head.Y <= 0 Or Snake_Head.X >= Columns Or Snake_Head.Y >= Rows Then
        Call Player1_Crashes
        Exit Sub
    End If
    If Snake_Head2.X <= 0 Or Snake_Head2.Y <= 0 Or Snake_Head2.X >= Columns Or Snake_Head2.Y >= Rows Then
        Call Player2_Crashes
        Exit Sub
    End If
End If
If Snake_Head.X = Snake_Head2.X And Snake_Head.Y = Snake_Head2.Y Then
    If Player1_Lives = 1 And Player2_Lives = 1 Then
        tmrMoveSnake.Enabled = False
        Player1_Lives = 0
        Player2_Lives = 0
        lblPlayer1Lives.Caption = "Player 1 Lives : " & Player1_Lives
        lblPlayer2Lives.Caption = "Player 2 Lives : " & Player2_Lives
        MsgBox "It's a TIE !", vbOKOnly + vbExclamation, "Snakes"
        Unload frmSnakesTwoPlayer
        frmMain.Show
        Exit Sub
    ElseIf Player1_Lives = 1 Then
        tmrMoveSnake.Enabled = False
        Player1_Lives = 0
        lblPlayer1Lives.Caption = "Player 1 Lives : " & Player1_Lives
        lblPlayer2Lives.Caption = "Player 2 Lives : " & Player2_Lives
        If AI_On = True Then
            MsgBox "You Lose !", vbOKOnly + vbExclamation, "Snakes"
        Else
            MsgBox "Player 2 Wins !", vbOKOnly + vbExclamation, "Snakes"
        End If
        Unload frmSnakesTwoPlayer
        frmMain.Show
        Exit Sub
    ElseIf Player2_Lives = 1 Then
        tmrMoveSnake.Enabled = False
        Player2_Lives = 0
        lblPlayer1Lives.Caption = "Player 1 Lives : " & Player1_Lives
        lblPlayer2Lives.Caption = "Player 2 Lives : " & Player2_Lives
        If AI_On = True Then
            MsgBox "You Win !", vbOKOnly + vbExclamation, "Snakes"
        Else
            MsgBox "Player 1 Wins !", vbOKOnly + vbExclamation, "Snakes"
        End If
        Unload frmSnakesTwoPlayer
        frmMain.Show
        Exit Sub
    Else
        Call Player1_Crashes
        Call Player2_Crashes
        Exit Sub
    End If
Else
    If Snake_Body_Count >= 6 Then
        Set temp_node = New Node
        Set temp_node = Snake_Head
        For i = 1 To Snake_Body_Count - 1
        'While Not temp_node.NextNode Is Snake_Tail
            Set temp_node = temp_node.NextNode
            If temp_node.X = Snake_Head.X And temp_node.Y = Snake_Head.Y Then
                Call Player1_Crashes
                Exit Sub
            ElseIf temp_node.X = Snake_Head2.X And temp_node.Y = Snake_Head2.Y Then
                Call Player2_Crashes
                Exit Sub
            End If
                       
            If AI_On = True And AI_Level > 1 Then
                If AI_Level = 3 Then
                    'Hardest Artifical Intelligence
                    If temp_node.X = Snake_Head2.X And temp_node.Y + 1 = Snake_Head2.Y Then
                        Occupied(1) = True
                    End If
                    If temp_node.X - 1 = Snake_Head2.X And temp_node.Y = Snake_Head2.Y Then
                        Occupied(2) = True
                    End If
                    If temp_node.X = Snake_Head2.X And temp_node.Y - 1 = Snake_Head2.Y Then
                        Occupied(3) = True
                    End If
                    If temp_node.X + 1 = Snake_Head2.X And temp_node.Y = Snake_Head2.Y Then
                        Occupied(4) = True
                    End If
                                    
                    If Occupied(1) = True Then
                        If Food_X < Snake_Head.X Then
                            If Occupied(4) = False Then
                                Direction2 = 4
                            Else
                                Direction2 = 2
                            End If
                        Else
                            If Occupied(2) = False Then
                                Direction2 = 2
                            Else
                                Direction2 = 4
                            End If
                        End If
                    ElseIf Occupied(2) = True Then
                        If Food_Y < Snake_Head.Y Then
                            If Occupied(1) = False Then
                                Direction2 = 1
                            Else
                                Direction2 = 3
                            End If
                        Else
                            If Occupied(3) = False Then
                                Direction2 = 3
                            Else
                                Direction2 = 1
                            End If
                        End If
                    ElseIf Occupied(3) = True Then
                        If Food_X < Snake_Head.X Then
                            If Occupied(4) = False Then
                                Direction2 = 4
                            Else
                                Direction2 = 2
                            End If
                        Else
                            If Occupied(2) = False Then
                                Direction2 = 2
                            Else
                                Direction2 = 4
                            End If
                        End If
                    ElseIf Occupied(4) = True Then
                        If Food_Y < Snake_Head.Y Then
                            If Occupied(1) = False Then
                                Direction2 = 1
                            Else
                                Direction2 = 3
                            End If
                        Else
                            If Occupied(3) = False Then
                                Direction2 = 3
                            Else
                                Direction2 = 1
                            End If
                        End If
                    End If
                Else
                    'Simple Artifical Intelligence
                    Select Case Direction2
                        Case Is = 1
                            If temp_node.X = Snake_Head2.X And temp_node.Y + 1 = Snake_Head2.Y Then
                                If Food_X < Snake_Head.X Then
                                    Direction2 = 4
                                Else
                                    Direction2 = 2
                                End If
                            End If
                        Case Is = 2
                            If temp_node.X - 1 = Snake_Head2.X And temp_node.Y = Snake_Head2.Y Then
                                If Food_Y < Snake_Head.Y Then
                                    Direction2 = 1
                                Else
                                    Direction2 = 3
                                End If
                            End If
                        Case Is = 3
                            If temp_node.X = Snake_Head2.X And temp_node.Y - 1 = Snake_Head2.Y Then
                                If Food_X < Snake_Head.X Then
                                    Direction2 = 4
                                Else
                                    Direction2 = 2
                                End If
                            End If
                        Case Is = 4
                            If temp_node.X + 1 = Snake_Head2.X And temp_node.Y = Snake_Head2.Y Then
                                If Food_Y < Snake_Head.Y Then
                                    Direction2 = 1
                                Else
                                    Direction2 = 3
                                End If
                            End If
                    End Select
                End If
            End If
        'Wend
        Next i
    End If
    If Snake_Body_Count2 >= 6 Then
        Set temp_node = New Node
        Set temp_node = Snake_Head2
        Dim Simple_AI_Done As Boolean
        For i = 1 To Snake_Body_Count2 - 1
        Occupied(1) = False
        Occupied(2) = False
        Occupied(3) = False
        Occupied(4) = False
        Simple_AI_Done = False
        'While Not temp_node Is Snake_Tail2
            Set temp_node = temp_node.NextNode
            If temp_node.X = Snake_Head.X And temp_node.Y = Snake_Head.Y Then
                Call Player1_Crashes
                Exit Sub
            ElseIf temp_node.X = Snake_Head2.X And temp_node.Y = Snake_Head2.Y Then
                Call Player2_Crashes
                Exit Sub
            End If
            If AI_On = True And AI_Level > 1 Then
                If AI_Level = 3 Then
                    Select Case Direction2
                        Case Is = 1
                            If temp_node.X = Snake_Head2.X And temp_node.Y + 1 = Snake_Head2.Y Then
                                Occupied(1) = True
                            End If
                        Case Is = 2
                            If temp_node.X - 1 = Snake_Head2.X And temp_node.Y = Snake_Head2.Y Then
                                Occupied(2) = True
                            End If
                        Case Is = 3
                            If temp_node.X = Snake_Head2.X And temp_node.Y - 1 = Snake_Head2.Y Then
                                Occupied(3) = True
                            End If
                        Case Is = 4
                            If temp_node.X + 1 = Snake_Head2.X And temp_node.Y = Snake_Head2.Y Then
                                Occupied(4) = True
                            End If
                    End Select
                    If No_Walls = False Then
                        If Snake_Head2.X < 2 Then
                            Occupied(4) = True
                        ElseIf Snake_Head2.X > Columns - 2 Then
                            Occupied(2) = True
                        ElseIf Snake_Head2.Y < 2 Then
                            Occupied(1) = True
                        ElseIf Snake_Head2.X > Rows - 2 Then
                            Occupied(3) = True
                        End If
                    End If
                    If Occupied(1) = True Then
                        If Occupied(4) = False Then
                            Direction2 = 4
                        Else
                            Direction2 = 2
                        End If
                    ElseIf Occupied(2) = True Then
                        If Occupied(1) = False Then
                            Direction2 = 1
                        Else
                            Direction2 = 3
                        End If
                    ElseIf Occupied(3) = True Then
                        If Occupied(4) = False Then
                            Direction2 = 4
                        Else
                            Direction2 = 2
                        End If
                    ElseIf Occupied(4) = True Then
                        If Occupied(1) = False Then
                            Direction2 = 1
                        Else
                            Direction2 = 3
                        End If
                    End If
                End If
            ElseIf Simple_AI_Done = False Then
                'Simple Artifical Intelligence
                Select Case Direction2
                    Case Is = 1
                        If temp_node.X = Snake_Head2.X And temp_node.Y + 2 = Snake_Head2.Y Then
                            If Food_X < Snake_Head.X Then
                                Direction2 = 4
                            Else
                                Direction2 = 2
                            End If
                            Simple_AI_Done = True
                        End If
                    Case Is = 2
                        If temp_node.X - 2 = Snake_Head2.X And temp_node.Y = Snake_Head2.Y Then
                            If Food_Y < Snake_Head.Y Then
                                Direction2 = 1
                            Else
                                Direction2 = 3
                            End If
                            Simple_AI_Done = True
                        End If
                    Case Is = 3
                        If temp_node.X = Snake_Head2.X And temp_node.Y - 2 = Snake_Head2.Y Then
                            If Food_X < Snake_Head.X Then
                                Direction2 = 4
                            Else
                                Direction2 = 2
                            End If
                            Simple_AI_Done = True
                        End If
                    Case Is = 4
                        If temp_node.X + 2 = Snake_Head2.X And temp_node.Y = Snake_Head2.Y Then
                            If Food_Y < Snake_Head.Y Then
                                Direction2 = 1
                            Else
                                Direction2 = 3
                            End If
                            Simple_AI_Done = True
                        End If
                End Select
            End If
        'Wend
        Next i
    End If
End If

If Player1_Lives = 0 Or Player2_Lives = 0 Then
    Exit Sub
End If

'Move Snake
Select Case Direction
    Case Is = 1
        If Player1_Collect_Food <= 0 Then
            Set temp_node = Snake_Head
            Set Snake_Head = Remove_Node_At_Back
            Snake_Head.NextNode = temp_node
            Snake_Head.Y = Snake_Head.NextNode.Y - 1
            Snake_Head.X = Snake_Head.NextNode.X
        Else
            Player1_Collect_Food = Player1_Collect_Food - 1
            Insert_Node_At_Front Snake_Head.X, Snake_Head.Y - 1
        End If
    Case Is = 2
        If Player1_Collect_Food <= 0 Then
            Set temp_node = Snake_Head
            Set Snake_Head = Remove_Node_At_Back
            Snake_Head.NextNode = temp_node
            Snake_Head.X = Snake_Head.NextNode.X + 1
            Snake_Head.Y = Snake_Head.NextNode.Y
        Else
            Player1_Collect_Food = Player1_Collect_Food - 1
            Insert_Node_At_Front Snake_Head.X + 1, Snake_Head.Y
        End If
    Case Is = 3
        If Player1_Collect_Food <= 0 Then
            Set temp_node = Snake_Head
            Set Snake_Head = Remove_Node_At_Back
            Snake_Head.NextNode = temp_node
            Snake_Head.Y = Snake_Head.NextNode.Y + 1
            Snake_Head.X = Snake_Head.NextNode.X
        Else
            Player1_Collect_Food = Player1_Collect_Food - 1
            Insert_Node_At_Front Snake_Head.X, Snake_Head.Y + 1
        End If
    Case Is = 4
        If Player1_Collect_Food <= 0 Then
            Set temp_node = Snake_Head
            Set Snake_Head = Remove_Node_At_Back
            Snake_Head.NextNode = temp_node
            Snake_Head.X = Snake_Head.NextNode.X - 1
            Snake_Head.Y = Snake_Head.NextNode.Y
        Else
            Player1_Collect_Food = Player1_Collect_Food - 1
            Insert_Node_At_Front Snake_Head.X - 1, Snake_Head.Y
        End If
End Select

Select Case Direction2
    Case Is = 1
        If Player2_Collect_Food <= 0 Then
            Set temp_node = Snake_Head2
            Set Snake_Head2 = Remove_Node_At_Back2
            Snake_Head2.NextNode = temp_node
            Snake_Head2.Y = Snake_Head2.NextNode.Y - 1
            Snake_Head2.X = Snake_Head2.NextNode.X
        Else
            Player2_Collect_Food = Player2_Collect_Food - 1
            Insert_Node_At_Front2 Snake_Head2.X, Snake_Head2.Y - 1
        End If
    Case Is = 2
        If Player2_Collect_Food <= 0 Then
            Set temp_node = Snake_Head2
            Set Snake_Head2 = Remove_Node_At_Back2
            Snake_Head2.NextNode = temp_node
            Snake_Head2.X = Snake_Head2.NextNode.X + 1
            Snake_Head2.Y = Snake_Head2.NextNode.Y
        Else
            Player2_Collect_Food = Player2_Collect_Food - 1
            Insert_Node_At_Front2 Snake_Head2.X + 1, Snake_Head2.Y
        End If
    Case Is = 3
        If Player2_Collect_Food <= 0 Then
            Set temp_node = Snake_Head2
            Set Snake_Head2 = Remove_Node_At_Back2
            Snake_Head2.NextNode = temp_node
            Snake_Head2.Y = Snake_Head2.NextNode.Y + 1
            Snake_Head2.X = Snake_Head2.NextNode.X
        Else
            Player2_Collect_Food = Player2_Collect_Food - 1
            Insert_Node_At_Front2 Snake_Head2.X, Snake_Head2.Y + 1
        End If
    Case Is = 4
        If Player2_Collect_Food <= 0 Then
            Set temp_node = Snake_Head2
            Set Snake_Head2 = Remove_Node_At_Back2
            Snake_Head2.NextNode = temp_node
            Snake_Head2.X = Snake_Head2.NextNode.X - 1
            Snake_Head2.Y = Snake_Head2.NextNode.Y
        Else
            Player2_Collect_Food = Player2_Collect_Food - 1
            Insert_Node_At_Front2 Snake_Head2.X - 1, Snake_Head2.Y
        End If
End Select

'Collision Detection with No Walls
If No_Walls = True Then
    If Direction = 4 And Snake_Head.X <= 0 Then
        Snake_Head.X = Columns
    ElseIf Direction = 2 And Snake_Head.X >= Columns Then
        Snake_Head.X = 1
    ElseIf Direction = 1 And Snake_Head.Y <= 0 Then
        Snake_Head.Y = Rows
    ElseIf Direction = 3 And Snake_Head.Y >= Rows Then
        Snake_Head.Y = 1
    End If
    If Direction2 = 4 And Snake_Head2.X <= 0 Then
        Snake_Head2.X = Columns
    ElseIf Direction2 = 2 And Snake_Head2.X >= Columns Then
        Snake_Head2.X = 1
    ElseIf Direction2 = 1 And Snake_Head2.Y <= 0 Then
        Snake_Head2.Y = Rows
    ElseIf Direction2 = 3 And Snake_Head2.Y >= Rows Then
        Snake_Head2.Y = 1
    End If
End If

'Draw Snakes
If Snake.UBound > 0 Then
    Snake(Snake_Head.Index).Left = Snake_Head.X * 120
    Snake(Snake_Head.Index).Top = Snake_Head.Y * 120
    Snake(Snake_Head.Index).Picture = imgHead(Direction).Picture
    Snake(Snake_Head.NextNode.Index).Picture = imgBody.Picture
    Snake(Snake_Tail.Index).Picture = imgTail(Direction).Picture
End If
If Snake2.UBound > 0 Then
    Snake2(Snake_Head2.Index).Left = Snake_Head2.X * 120
    Snake2(Snake_Head2.Index).Top = Snake_Head2.Y * 120
    Snake2(Snake_Head2.Index).Picture = imgHead2(Direction2).Picture
    Snake2(Snake_Head2.NextNode.Index).Picture = imgBody2.Picture
    Snake2(Snake_Tail2.Index).Picture = imgTail2(Direction2).Picture
End If

'Artifical Intelligence
If AI_On = True Then
    Select Case GameLibrary.Direction(Snake2(Snake_Head2.Index), Food)
        Case Is = 1
            If Direction2 <> 3 Then
                Direction2 = 1
            Else
                Direction2 = 2
            End If
        Case Is <= 8
            If Direction2 <> 4 Then
                Direction2 = 2
            Else
                Direction2 = 3
            End If
        Case Is = 9
            If Direction2 <> 1 Then
                Direction2 = 3
            Else
                Direction2 = 4
            End If
        Case Else
            If Direction2 <> 2 Then
                Direction2 = 4
            Else
                Direction2 = 1
            End If
    End Select
End If
End Sub

Public Sub Insert_Node_At_Front(New_X As Integer, New_Y As Integer)
Dim temp_node As Node
Set temp_node = New Node

If Is_Empty = True Then
    Snake_Body_Count = Snake_Body_Count + 1
    Load Snake(Snake_Body_Count)
    Snake(Snake_Body_Count).Width = 120
    Snake(Snake_Body_Count).Height = 120
    Snake(Snake_Body_Count).Visible = True
    Set Snake_Head = New Node
    Set Snake_Tail = New Node
    Snake_Head.Index = Snake_Body_Count
    Snake_Head.X = New_X
    Snake_Head.Y = New_Y
    Set Snake_Tail = Snake_Head
    Snake(Snake_Head.Index).Left = Snake_Head.X * 120
    Snake(Snake_Head.Index).Top = Snake_Head.Y * 120
    Snake(Snake_Head.Index).Picture = imgHead(Direction).Picture
Else
    Snake_Body_Count = Snake_Body_Count + 1
    Load Snake(Snake_Body_Count)
    Snake(Snake_Body_Count).Width = 120
    Snake(Snake_Body_Count).Height = 120
    Snake(Snake_Body_Count).Visible = True
    Set temp_node = Snake_Head
    Set Snake_Head = New Node
    Snake_Head.Index = Snake_Body_Count
    Snake_Head.X = New_X
    Snake_Head.Y = New_Y
    Snake_Head.NextNode = temp_node
    Snake(Snake_Head.Index).Left = Snake_Head.X * 120
    Snake(Snake_Head.Index).Top = Snake_Head.Y * 120
    Snake(Snake_Head.Index).Picture = imgHead(Direction).Picture
    Snake(Snake_Head.NextNode.Index).Picture = imgBody.Picture
    Snake(Snake_Tail.Index).Picture = imgTail(Direction).Picture
End If
End Sub

Public Function Remove_Node_At_Back() As Node
Dim temp_node As Node
Set Remove_Node_At_Back = New Node
If Is_Empty = True Then
    Set Remove_Node_At_Back = Nothing
    Exit Function
End If

If Snake_Head Is Snake_Tail Then
    Set Remove_Node_At_Back = Snake_Tail
    Set Snake_Head = Nothing
    Set Snake_Tail = Nothing
Else
    Set Remove_Node_At_Back = Snake_Tail
    Set temp_node = Snake_Head
    While Not temp_node.NextNode Is Snake_Tail
        Set temp_node = temp_node.NextNode
    Wend
    Set Snake_Tail = temp_node
    temp_node.NextNode = Nothing
End If
End Function

Public Function Is_Empty() As Boolean
If Snake_Head Is Nothing Then
    Is_Empty = True
Else
    Is_Empty = False
End If
End Function

Public Sub Insert_Node_At_Front2(New_X As Integer, New_Y As Integer)
Dim temp_node As Node
Set temp_node = New Node

If Is_Empty2 = True Then
    Snake_Body_Count2 = Snake_Body_Count2 + 1
    Load Snake2(Snake_Body_Count2)
    Snake2(Snake_Body_Count2).Width = 120
    Snake2(Snake_Body_Count2).Height = 120
    Snake2(Snake_Body_Count2).Visible = True
    Set Snake_Head2 = New Node
    Set Snake_Tail2 = New Node
    Snake_Head2.Index = Snake_Body_Count2
    Snake_Head2.X = New_X
    Snake_Head2.Y = New_Y
    Set Snake_Tail2 = Snake_Head2
    Snake2(Snake_Head2.Index).Left = Snake_Head2.X * 120
    Snake2(Snake_Head2.Index).Top = Snake_Head2.Y * 120
    Snake2(Snake_Head2.Index).Picture = imgHead2(Direction).Picture
Else
    Snake_Body_Count2 = Snake_Body_Count2 + 1
    Load Snake2(Snake_Body_Count2)
    Snake2(Snake_Body_Count2).Width = 120
    Snake2(Snake_Body_Count2).Height = 120
    Snake2(Snake_Body_Count2).Visible = True
    Set temp_node = Snake_Head2
    Set Snake_Head2 = New Node
    Snake_Head2.Index = Snake_Body_Count2
    Snake_Head2.X = New_X
    Snake_Head2.Y = New_Y
    Snake_Head2.NextNode = temp_node
    Snake2(Snake_Head2.Index).Left = Snake_Head2.X * 120
    Snake2(Snake_Head2.Index).Top = Snake_Head2.Y * 120
    Snake2(Snake_Head2.Index).Picture = imgHead2(Direction).Picture
    Snake2(Snake_Head2.NextNode.Index).Picture = imgBody2.Picture
    Snake2(Snake_Tail2.Index).Picture = imgTail2(Direction).Picture
End If
End Sub

Public Function Remove_Node_At_Back2() As Node
Dim temp_node As Node
Set Remove_Node_At_Back2 = New Node
If Is_Empty2 = True Then
    Set Remove_Node_At_Back2 = Nothing
    Exit Function
End If

If Snake_Head2 Is Snake_Tail2 Then
    Set Remove_Node_At_Back2 = Snake_Tail2
    Set Snake_Head2 = Nothing
    Set Snake_Tail2 = Nothing
Else
    Set Remove_Node_At_Back2 = Snake_Tail2
    Set temp_node = Snake_Head2
    While Not temp_node.NextNode Is Snake_Tail2
        Set temp_node = temp_node.NextNode
    Wend
    Set Snake_Tail2 = temp_node
    temp_node.NextNode = Nothing
End If
End Function

Public Function Is_Empty2() As Boolean
If Snake_Head2 Is Nothing Then
    Is_Empty2 = True
Else
    Is_Empty2 = False
End If
End Function

Public Sub Player1_Crashes()
Player1_Lives = Player1_Lives - 1
lblPlayer1Lives.Caption = "Player 1 Lives : " & Player1_Lives
If Player1_Lives = 0 And Player2_Lives > 0 Then
    tmrMoveSnake.Enabled = False
    If AI_On = True Then
        MsgBox "You Lose !", vbOKOnly + vbExclamation, "Snakes"
    Else
        MsgBox "Player 2 Wins !", vbOKOnly + vbExclamation, "Snakes"
    End If
    Unload frmSnakesTwoPlayer
    frmMain.Show
    Exit Sub
ElseIf Player1_Lives = 0 And Player2_Lives = 0 Then
    tmrMoveSnake.Enabled = False
    MsgBox "It's a Tie !", vbOKOnly + vbExclamation, "Snakes"
    Unload frmSnakesTwoPlayer
    frmMain.Show
    Exit Sub
Else
    Dim i As Integer
    For i = 1 To Snake_Body_Count
        Unload Snake(i)
    Next i
    Direction = 2
    Snake_Body_Count = 0
    lblPlayer1Score.Caption = "Player 1 Score : " & Player1_Score
    Set Snake_Head = Nothing
    Set Snake_Tail = Nothing
    If Snake_Head2.Y < Rows / 2 Then
        Insert_Node_At_Front 3, 56
        Insert_Node_At_Front 4, 56
        Insert_Node_At_Front 5, 56
        Insert_Node_At_Front 6, 56
        Insert_Node_At_Front 7, 56
        Insert_Node_At_Front 8, 56
    Else
        Insert_Node_At_Front 3, 3
        Insert_Node_At_Front 4, 3
        Insert_Node_At_Front 5, 3
        Insert_Node_At_Front 6, 3
        Insert_Node_At_Front 7, 3
        Insert_Node_At_Front 8, 3
    End If
End If
End Sub

Public Sub Player2_Crashes()
Player2_Lives = Player2_Lives - 1
lblPlayer2Lives.Caption = "Player 2 Lives : " & Player2_Lives
If Player2_Lives = 0 And Player1_Lives > 0 Then
    tmrMoveSnake.Enabled = False
    If AI_On = True Then
        MsgBox "You Win !", vbOKOnly + vbExclamation, "Snakes"
    Else
        MsgBox "Player 1 Wins !", vbOKOnly + vbExclamation, "Snakes"
    End If
    Unload frmSnakesTwoPlayer
    frmMain.Show
    Exit Sub
ElseIf Player2_Lives = 0 And Player1_Lives = 0 Then
    tmrMoveSnake.Enabled = False
    MsgBox "It's a TIE !", vbOKOnly + vbExclamation, "Snakes"
    Unload frmSnakesTwoPlayer
    frmMain.Show
    Exit Sub
Else
    Dim i As Integer
    For i = 1 To Snake_Body_Count2
        Unload Snake2(i)
    Next i
    Direction2 = 4
    Snake_Body_Count2 = 0
    lblPlayer2Score.Caption = "Player 2 Score : " & Player2_Score
    Set Snake_Head2 = Nothing
    Set Snake_Tail2 = Nothing
    If Snake_Head.Y < Rows / 2 Then
        Insert_Node_At_Front2 46, 56
        Insert_Node_At_Front2 45, 56
        Insert_Node_At_Front2 44, 56
        Insert_Node_At_Front2 43, 56
        Insert_Node_At_Front2 42, 56
        Insert_Node_At_Front2 41, 56
    Else
        Insert_Node_At_Front2 46, 3
        Insert_Node_At_Front2 45, 3
        Insert_Node_At_Front2 44, 3
        Insert_Node_At_Front2 43, 3
        Insert_Node_At_Front2 42, 3
        Insert_Node_At_Front2 41, 3
    End If
End If
End Sub

Private Sub tmrTimer_Timer()
If Is_Game_Paused = True Then Exit Sub
Timer.Width = Timer.Width - 5
If Timer.Width <= 10 Then
    Timer.Width = 0
    If Player1_Score > Player2_Score Then
        tmrMoveSnake.Enabled = False
        If AI_On = True Then
            MsgBox "Time is up - You Win !", vbOKOnly + vbExclamation, "Snakes"
        Else
            MsgBox "Time is up - Player 1 Wins !", vbOKOnly + vbExclamation, "Snakes"
        End If
        Unload frmSnakesTwoPlayer
        frmMain.Show
        Exit Sub
    ElseIf Player2_Score > Player1_Score Then
        tmrMoveSnake.Enabled = False
        If AI_On = True Then
            MsgBox "Time is up - You Lose !", vbOKOnly + vbExclamation, "Snakes"
        Else
            MsgBox "Time is up - Player 2 Wins !", vbOKOnly + vbExclamation, "Snakes"
        End If
        Unload frmSnakesTwoPlayer
        frmMain.Show
        Exit Sub
    Else
        tmrMoveSnake.Enabled = False
        MsgBox "Time is up - It's a TIE !", vbOKOnly + vbExclamation, "Snakes"
        Unload frmSnakesTwoPlayer
        frmMain.Show
        Exit Sub
    End If
End If
End Sub

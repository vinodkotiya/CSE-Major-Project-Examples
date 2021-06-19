VERSION 5.00
Begin VB.Form frmSnakes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snakes - One Player"
   ClientHeight    =   7395
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   5895
   Icon            =   "frmSnakes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrMoveSnake 
      Interval        =   10
      Left            =   3480
      Top             =   1080
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
      Left            =   4680
      TabIndex        =   1
      Top             =   7116
      Visible         =   0   'False
      Width           =   1116
   End
   Begin VB.Label lblScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Score : 0"
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
      Width           =   852
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   7092
      Left            =   0
      Top             =   0
      Width           =   132
   End
   Begin VB.Shape Shape2 
      FillStyle       =   0  'Solid
      Height          =   7092
      Left            =   5760
      Top             =   0
      Width           =   132
   End
   Begin VB.Shape Shape3 
      FillStyle       =   0  'Solid
      Height          =   132
      Left            =   120
      Top             =   0
      Width           =   5652
   End
   Begin VB.Shape Shape4 
      FillStyle       =   0  'Solid
      Height          =   132
      Left            =   120
      Top             =   6960
      Width           =   5652
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3480
      Picture         =   "frmSnakes.frx":030A
      Top             =   3000
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image nSnake 
      Height          =   120
      Index           =   6
      Left            =   2760
      Picture         =   "frmSnakes.frx":0AB8
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image Food 
      Appearance      =   0  'Flat
      Height          =   120
      Left            =   1560
      Picture         =   "frmSnakes.frx":0CAA
      Stretch         =   -1  'True
      Top             =   960
      Width           =   120
   End
   Begin VB.Image imgTail 
      Height          =   180
      Index           =   1
      Left            =   2760
      Picture         =   "frmSnakes.frx":1458
      Top             =   1680
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgTail 
      Height          =   180
      Index           =   3
      Left            =   2760
      Picture         =   "frmSnakes.frx":164A
      Top             =   2160
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgTail 
      Height          =   180
      Index           =   4
      Left            =   2520
      Picture         =   "frmSnakes.frx":183C
      Top             =   1920
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgTail 
      Height          =   180
      Index           =   2
      Left            =   3000
      Picture         =   "frmSnakes.frx":1A2E
      Top             =   1920
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgBody 
      Height          =   180
      Left            =   1320
      Picture         =   "frmSnakes.frx":1C20
      Top             =   1920
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgHead 
      Height          =   180
      Index           =   1
      Left            =   1920
      Picture         =   "frmSnakes.frx":1E12
      Top             =   1680
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgHead 
      Height          =   180
      Index           =   4
      Left            =   1680
      Picture         =   "frmSnakes.frx":2004
      Top             =   1920
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgHead 
      Height          =   180
      Index           =   2
      Left            =   2160
      Picture         =   "frmSnakes.frx":21F6
      Top             =   1920
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image imgHead 
      Height          =   180
      Index           =   3
      Left            =   1920
      Picture         =   "frmSnakes.frx":23E8
      Top             =   2040
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Image nSnake 
      Height          =   120
      Index           =   5
      Left            =   2640
      Picture         =   "frmSnakes.frx":25DA
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image nSnake 
      Height          =   120
      Index           =   4
      Left            =   2520
      Picture         =   "frmSnakes.frx":27CC
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image nSnake 
      Height          =   120
      Index           =   3
      Left            =   2400
      Picture         =   "frmSnakes.frx":29BE
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image nSnake 
      Height          =   120
      Index           =   2
      Left            =   2280
      Picture         =   "frmSnakes.frx":2BB0
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image nSnake 
      Height          =   120
      Index           =   1
      Left            =   2160
      Picture         =   "frmSnakes.frx":2DA2
      Stretch         =   -1  'True
      Top             =   3000
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Image Snake 
      Height          =   120
      Index           =   0
      Left            =   2280
      Picture         =   "frmSnakes.frx":2F94
      Stretch         =   -1  'True
      Top             =   1440
      Visible         =   0   'False
      Width           =   120
   End
End
Attribute VB_Name = "frmSnakes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' One player code module
Option Explicit
Private Direction As Integer
Private Snake_Head As Node
Private Snake_Tail As Node
Private Snake_Body_Count As Integer
Private Food_X As Integer
Private Food_Y As Integer
Private Player_Collect_Food As Integer
Const Columns = 48
Const Rows = 58
Private Is_Game_Paused As Boolean
Private Is_Ctrl_Down As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyP Then
    If Is_Game_Paused = False Then
        Is_Game_Paused = True
        lblGamePaused.Visible = True
    Else
        Is_Game_Paused = False
        lblGamePaused.Visible = False
    End If
ElseIf KeyCode = vbKeyEscape Then
        Unload frmSnakes
        frmMain.Show
End If
If Is_Game_Paused = True Then Exit Sub
Select Case KeyCode
    Case Is = vbKeyUp
        'If Direction <> 3 Then
        '    Direction = 1
        'End If
        If Snake_Head.X <> Snake_Head.NextNode.X And Snake_Head.Y - 1 <> Snake_Head.NextNode.Y Then
            Direction = 1
        End If
    Case Is = vbKeyRight
        'If Direction <> 4 Then
        '    Direction = 2
        'End If
        If Snake_Head.X + 1 <> Snake_Head.NextNode.X And Snake_Head.Y <> Snake_Head.NextNode.Y Then
            Direction = 2
        End If
    Case Is = vbKeyDown
        'If Direction <> 1 Then
        '    Direction = 3
        'End If
        If Snake_Head.X <> Snake_Head.NextNode.X And Snake_Head.Y + 1 <> Snake_Head.NextNode.Y Then
            Direction = 3
        End If
    Case Is = vbKeyLeft
        'If Direction <> 2 Then
        '    Direction = 4
        'End If
        If Snake_Head.X - 1 <> Snake_Head.NextNode.X And Snake_Head.Y <> Snake_Head.NextNode.Y Then
            Direction = 4
        End If
    Case Is = vbKeyF
        tmrMoveSnake.Interval = 1
    Case Is = vbKeyD
        tmrMoveSnake.Interval = 10
    Case Is = vbKeyS
        tmrMoveSnake.Interval = 100
    Case Is = vbKeyControl
        Is_Ctrl_Down = True
End Select
End Sub

Private Sub Form_Activate()
Snake_Body_Count = 0
Direction = 2
Player_Score = 0
Is_Game_Paused = False
Food_X = Int((Rnd * (Columns - 2)) + 1)
Food_Y = Int((Rnd * (Rows - 2)) + 1)
Food.Left = Food_X * 120
Food.Top = Food_Y * 120
tmrMoveSnake.Interval = 10
'Insert_Node_At_Front 3, 3
    Snake_Body_Count = Snake_Body_Count + 1
    Load Snake(Snake_Body_Count)
    Snake(Snake_Body_Count).Width = 120
    Snake(Snake_Body_Count).Height = 120
    Snake(Snake_Body_Count).Visible = True
    Set Snake_Head = New Node
    Set Snake_Tail = New Node
    Snake_Head.Index = Snake_Body_Count
    Snake_Head.X = 3
    Snake_Head.Y = 3
    Set Snake_Tail = Snake_Head
    Snake(Snake_Head.Index).Left = Snake_Head.X * 120
    Snake(Snake_Head.Index).Top = Snake_Head.Y * 120
    Snake(Snake_Head.Index).Picture = imgHead(Direction).Picture
Insert_Node_At_Front 4, 3
Insert_Node_At_Front 5, 3
Insert_Node_At_Front 6, 3
Insert_Node_At_Front 7, 3
Insert_Node_At_Front 8, 3
Insert_Node_At_Front 9, 3
Insert_Node_At_Front 10, 3
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyControl Then
    Is_Ctrl_Down = False
End If
End Sub

Private Sub tmrMoveSnake_Timer()
If Is_Game_Paused = True Then Exit Sub

Dim temp_node As Node
Select Case Direction
    Case Is = 1
        If Player_Collect_Food <= 0 Then
            Set temp_node = Snake_Head
            Set Snake_Head = Remove_Node_At_Back
            Snake_Head.NextNode = temp_node
            Snake_Head.Y = Snake_Head.NextNode.Y - 1
            Snake_Head.X = Snake_Head.NextNode.X
        Else
            Player_Collect_Food = Player_Collect_Food - 1
            Insert_Node_At_Front Snake_Head.X, Snake_Head.Y - 1
        End If
    Case Is = 2
        If Player_Collect_Food <= 0 Then
            Set temp_node = Snake_Head
            Set Snake_Head = Remove_Node_At_Back
            Snake_Head.NextNode = temp_node
            Snake_Head.X = Snake_Head.NextNode.X + 1
            Snake_Head.Y = Snake_Head.NextNode.Y
        Else
            Player_Collect_Food = Player_Collect_Food - 1
            Insert_Node_At_Front Snake_Head.X + 1, Snake_Head.Y
        End If
    Case Is = 3
        If Player_Collect_Food <= 0 Then
            Set temp_node = Snake_Head
            Set Snake_Head = Remove_Node_At_Back
            Snake_Head.NextNode = temp_node
            Snake_Head.Y = Snake_Head.NextNode.Y + 1
            Snake_Head.X = Snake_Head.NextNode.X
        Else
            Player_Collect_Food = Player_Collect_Food - 1
            Insert_Node_At_Front Snake_Head.X, Snake_Head.Y + 1
        End If
    Case Is = 4
        If Player_Collect_Food <= 0 Then
            Set temp_node = Snake_Head
            Set Snake_Head = Remove_Node_At_Back
            Snake_Head.NextNode = temp_node
            Snake_Head.X = Snake_Head.NextNode.X - 1
            Snake_Head.Y = Snake_Head.NextNode.Y
        Else
            Player_Collect_Food = Player_Collect_Food - 1
            Insert_Node_At_Front Snake_Head.X - 1, Snake_Head.Y
        End If
End Select

If Snake_Head.X <= 0 Or Snake_Head.Y <= 0 Or Snake_Head.X >= Columns Or Snake_Head.Y >= Rows Then
    tmrMoveSnake.Enabled = False
    frmHighScores.Show vbModal
    Unload frmSnakes
    frmMain.Show
    Dim i As Integer
    For i = Snake_Body_Count To 1
        Unload Snake(i)
    Next i
    Snake_Body_Count = 0
    Set Snake_Head = Nothing
    Set Snake_Tail = Nothing
    Exit Sub
End If

Snake(Snake_Head.Index).Left = Snake_Head.X * 120
Snake(Snake_Head.Index).Top = Snake_Head.Y * 120
Snake(Snake_Head.Index).Picture = imgHead(Direction).Picture
Snake(Snake_Head.NextNode.Index).Picture = imgBody.Picture
Snake(Snake_Tail.Index).Picture = imgTail(Direction).Picture

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
    Player_Score = Player_Score + Points
    lblScore.Caption = "Score : " & Player_Score
    Player_Collect_Food = 4
End If

'Collision Detection
If Snake_Body_Count > 6 Then
    Set temp_node = New Node
    Set temp_node = Snake_Head
    While Not temp_node.NextNode Is Snake_Tail
        Set temp_node = temp_node.NextNode
        If temp_node.X = Snake_Head.X And temp_node.Y = Snake_Head.Y Then
            Call Player_Crashes
            Exit Sub
        End If
    Wend
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

Public Sub Player_Crashes()
tmrMoveSnake.Enabled = False
frmHighScores.Show vbModal
Unload frmSnakes
frmMain.Show
Dim i As Integer
For i = Snake_Body_Count To 1
    Unload Snake(i)
Next i
Snake_Body_Count = 0
Set Snake_Head = Nothing
Set Snake_Tail = Nothing
End Sub


VERSION 5.00
Begin VB.Form frmSelectPlayer 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Opponent"
   ClientHeight    =   1725
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   6900
   Icon            =   "frmSelectPlayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   6900
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image picOpponent 
      Height          =   1440
      Index           =   3
      Left            =   2760
      Picture         =   "frmSelectPlayer.frx":030A
      Top             =   120
      Width           =   1680
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dexter"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   372
      Index           =   5
      Left            =   5580
      TabIndex        =   4
      Top             =   1280
      Width           =   1092
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Emilio"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   372
      Index           =   3
      Left            =   2920
      TabIndex        =   3
      Top             =   1280
      Width           =   1092
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vipul"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   372
      Index           =   4
      Left            =   4240
      TabIndex        =   2
      Top             =   1280
      Width           =   1092
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vikram"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   372
      Index           =   2
      Left            =   1600
      TabIndex        =   1
      Top             =   1280
      Width           =   1092
   End
   Begin VB.Label lblName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Sangram"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   372
      Index           =   1
      Left            =   300
      TabIndex        =   0
      Top             =   1280
      Width           =   1092
   End
   Begin VB.Image picOpponent 
      Height          =   1440
      Index           =   1
      Left            =   120
      Picture         =   "frmSelectPlayer.frx":814C
      Top             =   120
      Width           =   1680
   End
   Begin VB.Image picOpponent 
      Height          =   1440
      Index           =   5
      Left            =   5400
      Picture         =   "frmSelectPlayer.frx":FF8E
      Top             =   120
      Width           =   1680
   End
   Begin VB.Image picOpponent 
      Height          =   1440
      Index           =   4
      Left            =   4080
      Picture         =   "frmSelectPlayer.frx":17DD0
      Top             =   120
      Width           =   1680
   End
   Begin VB.Image picOpponent 
      Height          =   1440
      Index           =   2
      Left            =   1440
      Picture         =   "frmSelectPlayer.frx":1FC12
      Top             =   120
      Width           =   1680
   End
End
Attribute VB_Name = "frmSelectPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload frmSelectPlayer
    frmStartUp.Show
'ElseIf KeyCode = vbKey1 Then
'    Call picOpponent_MouseUp(1, 0, 0, 0, 0)
'ElseIf KeyCode = vbKey2 Then
'    Call picOpponent_MouseUp(2, 0, 0, 0, 0)
'ElseIf KeyCode = vbKey3 Then
'    Call picOpponent_MouseUp(3, 0, 0, 0, 0)
'ElseIf KeyCode = vbKey4 Then
'    Call picOpponent_MouseUp(4, 0, 0, 0, 0)
'ElseIf KeyCode = vbKey5 Then
'    Call picOpponent_MouseUp(5, 0, 0, 0, 0)
End If
End Sub

Private Sub picOpponent_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If picOpponent(Index).BorderStyle = 0 Then
    For i = 1 To 5
        picOpponent(i).BorderStyle = 0
    Next i
    picOpponent(Index).ZOrder
    picOpponent(Index).Height = 1512
    picOpponent(Index).BorderStyle = 1
End If
End Sub

Private Sub picOpponent_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Opponent = Index
frmPlayer1.lblOpponent.Caption = "Opponent : " & lblName(Index).Caption
frmPlayer1.OpponentPortrait.Picture = frmPlayer1.Portrait(Index).Picture
Unload frmSelectPlayer
frmControls.Show
End Sub

VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Snakes"
   ClientHeight    =   3030
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   2820
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   2820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   732
      Left            =   480
      TabIndex        =   3
      Top             =   1680
      Width           =   2052
      Begin VB.OptionButton Option4 
         Caption         =   "Play Against Computer"
         Height          =   252
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   1932
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Play Against Friend"
         Height          =   252
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1932
      End
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Two Players"
      Height          =   252
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Value           =   -1  'True
      Width           =   1332
   End
   Begin VB.OptionButton Option1 
      Caption         =   "One Player"
      Height          =   252
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1332
   End
   Begin VB.CommandButton cmdStartGame 
      Caption         =   "Start Game"
      Height          =   372
      Left            =   840
      TabIndex        =   0
      Top             =   2520
      Width           =   1212
   End
   Begin VB.Image Tail 
      Appearance      =   0  'Flat
      Height          =   372
      Index           =   1
      Left            =   -120
      Picture         =   "frmMain.frx":030A
      Stretch         =   -1  'True
      Top             =   480
      Width           =   396
   End
   Begin VB.Image Image4 
      Appearance      =   0  'Flat
      Height          =   372
      Left            =   600
      Picture         =   "frmMain.frx":0614
      Stretch         =   -1  'True
      Top             =   480
      Width           =   396
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      Height          =   372
      Left            =   960
      Picture         =   "frmMain.frx":091E
      Stretch         =   -1  'True
      Top             =   480
      Width           =   396
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   372
      Left            =   1320
      Picture         =   "frmMain.frx":0C28
      Stretch         =   -1  'True
      Top             =   480
      Width           =   396
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   372
      Left            =   1680
      Picture         =   "frmMain.frx":0F32
      Stretch         =   -1  'True
      Top             =   480
      Width           =   396
   End
   Begin VB.Image Head 
      Appearance      =   0  'Flat
      Height          =   360
      Index           =   1
      Left            =   2400
      Picture         =   "frmMain.frx":123C
      Stretch         =   -1  'True
      Top             =   492
      Width           =   384
   End
   Begin VB.Image BodyH 
      Appearance      =   0  'Flat
      Height          =   372
      Left            =   2040
      Picture         =   "frmMain.frx":1546
      Stretch         =   -1  'True
      Top             =   480
      Width           =   396
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Snakes"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   720
      TabIndex        =   6
      Top             =   0
      Width           =   1320
   End
   Begin VB.Image Image5 
      Appearance      =   0  'Flat
      Height          =   372
      Left            =   240
      Picture         =   "frmMain.frx":1850
      Stretch         =   -1  'True
      Top             =   480
      Width           =   396
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStartGame_Click()
If Option1.Value = True Then
    frmSnakes.Show
    Unload frmMain
Else
    If Option4.Value = True Then
        AI_On = True
        frmControls1.Show vbModal
    Else
        AI_On = False
        frmControls2.Show vbModal
    End If
End If
End Sub

Private Sub cmdStartGame_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    End
End If
End Sub

Private Sub Form_Activate()
If One_Player = False Then
    Option1.Value = False
    Option2.Value = True
Else
    Option1.Value = True
    Option2.Value = False
End If
If Main_Initialized = False Then
    Main_Initialized = True
    AI_On = True
    Option3.Value = False
    Option4.Value = True
Else
    If AI_On = True Then
        Option3.Value = False
        Option4.Value = True
    Else
        Option3.Value = True
        Option4.Value = False
    End If
End If
End Sub

Private Sub Option1_Click()
One_Player = True
Option3.Enabled = False
Option4.Enabled = False
End Sub

Private Sub Option1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call cmdStartGame_Click
ElseIf KeyCode = vbKeyEscape Then
    End
End If
End Sub

Private Sub Option2_Click()
One_Player = False
Option3.Enabled = True
Option4.Enabled = True
End Sub

Private Sub Option2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call cmdStartGame_Click
ElseIf KeyCode = vbKeyEscape Then
    End
End If
End Sub

Private Sub Option3_Click()
AI_On = False
End Sub

Private Sub Option3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call cmdStartGame_Click
ElseIf KeyCode = vbKeyEscape Then
    End
End If
End Sub

Private Sub Option4_Click()
AI_On = True
End Sub

Private Sub Option4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call cmdStartGame_Click
ElseIf KeyCode = vbKeyEscape Then
    End
End If
End Sub

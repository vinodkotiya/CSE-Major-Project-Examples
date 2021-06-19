VERSION 5.00
Begin VB.Form frmStartUp 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "One-On-One Soccer"
   ClientHeight    =   3195
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   3165
   Icon            =   "frmStartUp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3165
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   1800
   End
   Begin VB.CommandButton cmd2Player 
      Caption         =   "Two Player Game"
      Height          =   324
      Left            =   840
      TabIndex        =   1
      Top             =   2400
      Width           =   1452
   End
   Begin VB.CommandButton cmd1Player 
      Caption         =   "One Player Game"
      Height          =   324
      Left            =   840
      TabIndex        =   0
      Top             =   1920
      Width           =   1452
   End
   Begin VB.Image imgEnd 
      Height          =   360
      Left            =   2400
      Picture         =   "frmStartUp.frx":030A
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   768
   End
   Begin VB.Image Image2 
      Height          =   492
      Left            =   1320
      Picture         =   "frmStartUp.frx":16DE
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   492
   End
   Begin VB.Image Ball 
      Height          =   480
      Left            =   0
      Picture         =   "frmStartUp.frx":19E8
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Soccer"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   660
      Left            =   852
      TabIndex        =   3
      Top             =   480
      Width           =   1584
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "One-On-One"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   384
      Left            =   756
      TabIndex        =   2
      Top             =   120
      Width           =   1740
   End
End
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private x_vel As Integer
Private y_vel As Integer

Private Sub cmd1Player_Click()
frmStartUp.Hide
frmSelectPlayer.Show 1
End Sub

Private Sub cmd1Player_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    If Show_About_Window = True Then
        Unload frmStartUp
        frmAbout.Show
    Else
        End
    End If
End If
End Sub

Private Sub cmd2Player_Click()
frmControls2.Show 1
End Sub

Private Sub cmd2Player_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    If Show_About_Window = True Then
        Unload frmStartUp
        frmAbout.Show
    Else
        End
    End If
End If
End Sub

Private Sub Form_Load()
x_vel = 240
y_vel = -240
End Sub

Private Sub imgEnd_Click()
If Show_About_Window = True Then
    Unload frmStartUp
    frmAbout.Show
Else
    End
End If
End Sub

Private Sub Timer1_Timer()
Ball.Left = Ball.Left + x_vel
Ball.Top = Ball.Top + y_vel
If Ball.Left < 0 Or Ball.Left > frmStartUp.Width - 480 Then
    x_vel = -x_vel
End If
If Ball.Top < 0 Or Ball.Top > frmStartUp.Height - 720 Then
    y_vel = -y_vel
End If
End Sub

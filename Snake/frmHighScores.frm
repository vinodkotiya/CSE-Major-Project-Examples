VERSION 5.00
Begin VB.Form frmHighScores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High Scores"
   ClientHeight    =   2580
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   1440
   ControlBox      =   0   'False
   Icon            =   "frmHighScores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   1440
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picScores 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1932
      Left            =   600
      ScaleHeight     =   1935
      ScaleWidth      =   735
      TabIndex        =   1
      Top             =   120
      Width           =   732
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   972
   End
   Begin VB.Image Arrow 
      Height          =   252
      Left            =   120
      Picture         =   "frmHighScores.frx":030A
      Stretch         =   -1  'True
      Top             =   100
      Width           =   384
   End
End
Attribute VB_Name = "frmHighScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Score procedure
Option Explicit
Private High_Scores(1 To 10) As Integer
Dim i, j As Integer

Private Sub cmdOK_Click()
Unload frmHighScores
End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Then
    Unload frmHighScores
End If
End Sub

Private Sub Form_Activate()
Dim Buffer As String
Dim Counter As Integer

Counter = 0

Randomize

Open "C:\Windows\Snakes.dat" For Append As #1
Close #1

Open "C:\Windows\Snakes.dat" For Input As #1
    If Not EOF(1) Then
        For i = 1 To 9
            If EOF(1) Then
                Exit For
            End If
            Line Input #1, Buffer
            Counter = Counter + 1
            High_Scores(i) = Val(Buffer)
        Next i
        High_Scores(Counter + 1) = Player_Score
    Else
        High_Scores(1) = Player_Score
    End If
Close #1

Call Bubble_Sort_Scores(Counter + 1)

Open "C:\Windows\Snakes.dat" For Output As #1
    For i = 1 To Counter + 1
        Write #1, High_Scores(i)
        picScores.Print High_Scores(i)
    Next i
Close #1
End Sub


Public Sub Bubble_Sort_Scores(Number_of_Scores As Integer)
Dim Buffer As Integer

For i = 1 To Number_of_Scores
    For j = 1 To Number_of_Scores
        If High_Scores(i) >= High_Scores(j) Then
            Buffer = High_Scores(j)
            High_Scores(j) = High_Scores(i)
            High_Scores(i) = Buffer
        End If
    Next j
Next i

For i = 1 To Number_of_Scores
    If Player_Score = High_Scores(i) Then
        Arrow.Top = 100 + (190 * (i - 1))
        Exit For
    End If
Next i
End Sub

Private Sub picScores_GotFocus()
On Error Resume Next
cmdOK.SetFocus
End Sub

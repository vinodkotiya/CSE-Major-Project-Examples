VERSION 5.00
Begin VB.Form frm1Player 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "One Player Tic Tac Toe"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4200
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4200
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer4 
      Interval        =   1
      Left            =   5640
      Top             =   480
   End
   Begin VB.Timer Timer3 
      Left            =   5640
      Top             =   960
   End
   Begin VB.Timer Timer2 
      Left            =   5040
      Top             =   960
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5040
      Top             =   360
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Game"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   840
      Width           =   975
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      X1              =   3240
      X2              =   3240
      Y1              =   1440
      Y2              =   2520
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   16
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   15
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Losses"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   14
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Wins"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   13
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblnum 
      Caption         =   "1"
      Height          =   255
      Left            =   5640
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label turn 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   4920
      TabIndex        =   11
      Top             =   600
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      X1              =   840
      X2              =   840
      Y1              =   120
      Y2              =   2400
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      X1              =   1680
      X2              =   1680
      Y1              =   120
      Y2              =   2400
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      X1              =   120
      X2              =   2400
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      X1              =   120
      X2              =   2400
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label s 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   615
   End
   Begin VB.Label s 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   1
      Left            =   1080
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.Label s 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   2
      Left            =   1800
      TabIndex        =   8
      Top             =   120
      Width           =   615
   End
   Begin VB.Label s 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   615
   End
   Begin VB.Label s 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   4
      Left            =   1080
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.Label s 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   5
      Left            =   1800
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.Label s 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   6
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label s 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   7
      Left            =   1080
      TabIndex        =   3
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label s 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Index           =   8
      Left            =   1800
      TabIndex        =   2
      Top             =   1800
      Width           =   615
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      Index           =   0
      X1              =   120
      X2              =   2400
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      Index           =   1
      X1              =   120
      X2              =   2400
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      Index           =   2
      X1              =   120
      X2              =   2400
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      Index           =   3
      X1              =   360
      X2              =   360
      Y1              =   120
      Y2              =   2400
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      Index           =   4
      X1              =   1200
      X2              =   1200
      Y1              =   120
      Y2              =   2400
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      Index           =   5
      X1              =   1920
      X2              =   1920
      Y1              =   120
      Y2              =   2400
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      Index           =   6
      X1              =   240
      X2              =   2280
      Y1              =   240
      Y2              =   2280
   End
   Begin VB.Line Line5 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      Index           =   7
      X1              =   120
      X2              =   2280
      Y1              =   2280
      Y2              =   120
   End
End
Attribute VB_Name = "frm1Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iwon As Boolean
Dim cwon As Boolean
Dim tie As Boolean

Private Sub Command1_Click()
s(0).Caption = ""
For Index = 1 To 8
    num = num + 1
    s(num).Caption = ""
Next Index
Line5(0).Visible = False
For i = 1 To 7
    x = x + 1
    Line5(x).Visible = False
Next i
lblnum.Caption = 1
Timer1.Interval = 1
Timer2.Interval = 0
Timer3.Interval = 0
Timer4.Interval = 1
cwon = False
iwon = False
tie = False
turn.Caption = "X"
End Sub

Private Sub Command2_Click()
Unload Me
frmTTTMain.Show
End Sub

Private Sub Form_Load()
turn.Caption = "X"

Line5(0).Visible = False
For c = 1 To 7
    h = h + 1
    Line5(h).Visible = False
Next c
End Sub



Private Sub s_Click(y As Integer)
If turn.Caption = "X" And s(y).Caption = "" And cwon = False And iwon = False And tie = False Then
s(y).Caption = "X"
turn.Caption = "O"
Call check
lblnum = Val(lblnum.Caption) + 2
End If
End Sub

Private Sub Timer1_Timer()
'Checks to see if anyone has one

If s(0).Caption = "X" And s(1).Caption = "X" And s(2).Caption = "X" Then

iwon = True
cwon = False
tie = False
Line5(0).Visible = True
End If

If s(3).Caption = "X" And s(4).Caption = "X" And s(5).Caption = "X" Then

iwon = True
cwon = False
tie = False
Line5(1).Visible = True
End If

If s(6).Caption = "X" And s(7).Caption = "X" And s(8).Caption = "X" Then

iwon = True
cwon = False
tie = False
Line5(2).Visible = True

End If

If s(0).Caption = "X" And s(3).Caption = "X" And s(6).Caption = "X" Then

iwon = True
cwon = False
tie = False
Line5(3).Visible = True


End If

If s(1).Caption = "X" And s(4).Caption = "X" And s(7).Caption = "X" Then

iwon = True
cwon = False
tie = False
Line5(4).Visible = True

End If

If s(2).Caption = "X" And s(5).Caption = "X" And s(8).Caption = "X" Then

iwon = True
cwon = False
tie = False
Line5(5).Visible = True

End If

If s(0).Caption = "X" And s(4).Caption = "X" And s(8).Caption = "X" Then

iwon = True
cwon = False
tie = False
Line5(6).Visible = True

End If

If s(2).Caption = "X" And s(4).Caption = "X" And s(6).Caption = "X" Then

iwon = True
cwon = False
tie = False
Line5(7).Visible = True
End If

If s(0).Caption = "O" And s(1).Caption = "O" And s(2).Caption = "O" Then

iwon = False
cwon = True
tie = False
Line5(0).Visible = True

End If

If s(3).Caption = "O" And s(4).Caption = "O" And s(5).Caption = "O" Then

iwon = False
cwon = True
tie = False
Line5(1).Visible = True

End If

If s(6).Caption = "O" And s(7).Caption = "O" And s(8).Caption = "O" Then

iwon = False
cwon = True
tie = False
Line5(2).Visible = True

End If

If s(0).Caption = "O" And s(3).Caption = "O" And s(6).Caption = "O" Then

iwon = False
cwon = True
tie = False
Line5(3).Visible = True

End If

If s(1).Caption = "O" And s(4).Caption = "O" And s(7).Caption = "O" Then

iwon = False
cwon = True
tie = False
Line5(4).Visible = True
End If

If s(2).Caption = "O" And s(5).Caption = "O" And s(8).Caption = "O" Then

iwon = False
cwon = True
tie = False
Line5(5).Visible = True


End If

If s(0).Caption = "O" And s(4).Caption = "O" And s(8).Caption = "O" Then

iwon = False
cwon = True
tie = False
Line5(6).Visible = True

End If

If s(2).Caption = "O" And s(4).Caption = "O" And s(6).Caption = "O" Then

iwon = False
cwon = True
tie = False
Line5(7).Visible = True
End If

If s(0) <> "" And s(1) <> "" And s(2) <> "" And s(3) <> "" And s(4) <> "" And s(5) <> "" And s(6) <> "" And s(7) <> "" And s(8) <> "" And cwon = False And iwon = False Then
tie = True
iwon = False
cwon = False
End If
End Sub

Private Sub Timer2_Timer()

'this is where the computer "thinks"
If iwon = False And cwon = False And tie = False Then
If turn.Caption = "O" And s(0).Caption = "X" And s(1).Caption = "X" And s(2).Caption = "" Or s(0).Caption = "O" And s(1).Caption = "O" And s(2).Caption = "" Then
s(2).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
ElseIf turn.Caption = "O" And s(0).Caption = "X" And s(2).Caption = "X" And s(1).Caption = "" Or s(0).Caption = "O" And s(2).Caption = "O" And s(1).Caption = "" Then
s(1).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
ElseIf turn.Caption = "O" And s(1).Caption = "X" And s(2).Caption = "X" And s(0).Caption = "" Or s(1).Caption = "O" And s(2).Caption = "O" And s(0).Caption = "" Then
s(0).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0

ElseIf turn.Caption = "O" And s(3).Caption = "X" And s(4).Caption = "X" And s(5).Caption = "" Or s(3).Caption = "O" And s(4).Caption = "O" And s(5).Caption = "" Then
s(5).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
ElseIf turn.Caption = "O" And s(3).Caption = "X" And s(5).Caption = "X" And s(4).Caption = "" Or s(0).Caption = "O" And s(5).Caption = "O" And s(4).Caption = "" Then
s(4).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
ElseIf turn.Caption = "O" And s(4).Caption = "X" And s(5).Caption = "X" And s(3).Caption = "" Or s(4).Caption = "O" And s(5).Caption = "O" And s(3).Caption = "" Then
s(3).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
ElseIf turn.Caption = "O" And s(6).Caption = "X" And s(7).Caption = "X" And s(8).Caption = "" Or s(6).Caption = "O" And s(7).Caption = "O" And s(8).Caption = "" Then
s(8).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
ElseIf turn.Caption = "O" And s(6).Caption = "X" And s(8).Caption = "X" And s(7).Caption = "" Or s(6).Caption = "O" And s(8).Caption = "O" And s(7).Caption = "" Then
s(7).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
ElseIf turn.Caption = "O" And s(7).Caption = "X" And s(8).Caption = "X" And s(6).Caption = "" Or s(7).Caption = "O" And s(8).Caption = "O" And s(6).Caption = "" Then
s(6).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
ElseIf turn.Caption = "O" And s(0).Caption = "X" And s(3).Caption = "X" And s(6).Caption = "" Or s(0).Caption = "O" And s(3).Caption = "O" And s(6).Caption = "" Then
s(6).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
ElseIf turn.Caption = "O" And s(0).Caption = "X" And s(6).Caption = "X" And s(3).Caption = "" Or s(0).Caption = "O" And s(6).Caption = "O" And s(3).Caption = "" Then
s(3).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
ElseIf turn.Caption = "O" And s(3).Caption = "X" And s(6).Caption = "X" And s(0).Caption = "" Or s(3).Caption = "O" And s(6).Caption = "O" And s(0).Caption = "" Then
s(0).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
ElseIf turn.Caption = "O" And s(1).Caption = "X" And s(4).Caption = "X" And s(7).Caption = "" Or s(1).Caption = "O" And s(4).Caption = "O" And s(7).Caption = "" Then
s(7).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
ElseIf turn.Caption = "O" And s(1).Caption = "X" And s(7).Caption = "X" And s(4).Caption = "" Or s(1).Caption = "O" And s(7).Caption = "O" And s(4).Caption = "" Then
s(4).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
ElseIf turn.Caption = "O" And s(4).Caption = "X" And s(7).Caption = "X" And s(1).Caption = "" Or s(4).Caption = "O" And s(7).Caption = "O" And s(1).Caption = "" Then
s(1).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
ElseIf turn.Caption = "O" And s(2).Caption = "X" And s(5).Caption = "X" And s(8).Caption = "" Or s(2).Caption = "O" And s(5).Caption = "O" And s(8).Caption = "" Then
s(8).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
ElseIf turn.Caption = "O" And s(2).Caption = "X" And s(8).Caption = "X" And s(5).Caption = "" Or s(2).Caption = "O" And s(8).Caption = "O" And s(5).Caption = "" Then
s(5).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
ElseIf turn.Caption = "O" And s(5).Caption = "X" And s(8).Caption = "X" And s(2).Caption = "" Or s(5).Caption = "O" And s(8).Caption = "O" And s(2).Caption = "" Then
s(2).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
ElseIf turn.Caption = "O" And s(0).Caption = "X" And s(4).Caption = "X" And s(8).Caption = "" Or s(0).Caption = "O" And s(4).Caption = "O" And s(8).Caption = "" Then
s(8).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
ElseIf turn.Caption = "O" And s(0).Caption = "X" And s(8).Caption = "X" And s(4).Caption = "" Or s(0).Caption = "O" And s(8).Caption = "O" And s(4).Caption = "" Then
s(4).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
ElseIf turn.Caption = "O" And s(4).Caption = "X" And s(8).Caption = "X" And s(0).Caption = "" Or s(4).Caption = "O" And s(8).Caption = "O" And s(0).Caption = "" Then
s(0).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
ElseIf turn.Caption = "O" And s(2).Caption = "X" And s(4).Caption = "X" And s(6).Caption = "" Or s(2).Caption = "O" And s(4).Caption = "O" And s(6).Caption = "" Then
s(6).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
ElseIf turn.Caption = "O" And s(2).Caption = "X" And s(6).Caption = "X" And s(4).Caption = "" Or s(2).Caption = "O" And s(6).Caption = "O" And s(4).Caption = "" Then
s(4).Caption = "O"
turn.Caption = "X"
Timer2.Interval = 0
Else:
Timer2.Interval = 0
Timer3.Interval = 1
End If
End If
End Sub

Private Sub check()
If lblnum.Caption < 4 Then

Timer3.Interval = 1

End If
If lblnum.Caption >= 5 Then
Timer2.Interval = 1
End If
End Sub

Private Sub Timer3_Timer()
choose:
Randomize
x = Int(Rnd * 9) + 1
 
Select Case x

    Case 1
        If s(0).Caption = "" Then
            s(0).Caption = "O"
        turn.Caption = "X"
        Timer3.Interval = 0
        ElseIf s(0).Caption = "X" Or s(0).Caption = "O" Then
        GoTo choose
        End If
    Case 2
        If s(1).Caption = "" Then
           s(1).Caption = "O"
           turn.Caption = "X"
        Timer3.Interval = 0
        ElseIf s(1).Caption = "X" Or s(1).Caption = "O" Then
        GoTo choose
        End If
    Case 3
        If s(2).Caption = "" Then
           s(2).Caption = "O"
           turn.Caption = "X"
        Timer3.Interval = 0
        ElseIf s(2).Caption = "X" Or s(2).Caption = "O" Then
        GoTo choose
        End If
    Case 4
        If s(3).Caption = "" Then
           s(3).Caption = "O"
           turn.Caption = "X"
        Timer3.Interval = 0
        ElseIf s(3).Caption = "X" Or s(3).Caption = "O" Then
        GoTo choose
        End If
    Case 5
        If s(4).Caption = "" Then
           s(4).Caption = "O"
           turn.Caption = "X"
        Timer3.Interval = 0
        ElseIf s(4).Caption = "X" Or s(4).Caption = "O" Then
        GoTo choose
        End If
    Case 6
        If s(5).Caption = "" Then
           s(5).Caption = "O"
           turn.Caption = "X"
        Timer3.Interval = 0
        ElseIf s(5).Caption = "X" Or s(5).Caption = "O" Then
        GoTo choose
        End If
    Case 7
        If s(6).Caption = "" Then
           s(6).Caption = "O"
           turn.Caption = "X"
        Timer3.Interval = 0
        ElseIf s(6).Caption = "X" Or s(6).Caption = "O" Then
        GoTo choose
        End If
    Case 8
        If s(7).Caption = "" Then
           s(7).Caption = "O"
           turn.Caption = "X"
        Timer3.Interval = 0
        ElseIf s(7).Caption = "X" Or s(7).Caption = "O" Then
        GoTo choose
        End If
    Case 9
        If s(8).Caption = "" Then
           s(8).Caption = "O"
           turn.Caption = "X"
        Timer3.Interval = 0
        ElseIf s(8).Caption = "X" Or s(8).Caption = "O" Then
        GoTo choose
        End If

End Select

End Sub

Private Sub Timer4_Timer()
If cwon = True Then
Timer1.Interval = 0
Timer2.Interval = 0
Timer3.Interval = 0
Timer4.Interval = 0

lost = MsgBox("The computer beat you!", vbExclamation, "Tic Tac Toe")
Label1(3).Caption = Val(Label1(3)) + 1
turn.Caption = ""
End If

If iwon = True Then
Timer1.Interval = 0
Timer2.Interval = 0
Timer3.Interval = 0
Timer4.Interval = 0
win = MsgBox("You Beat the computer!", vbExclamation, "Tic Tac Toe")
Label1(2).Caption = Val(Label1(2)) + 1
turn.Caption = ""
End If

If tie = True Then
Timer1.Interval = 0
Timer2.Interval = 0
Timer3.Interval = 0
Timer4.Interval = 0
atie = MsgBox("There is no winner.  The game is a tie!", vbExclamation, "Tic Tac Toe")
turn.Caption = ""

End If

End Sub


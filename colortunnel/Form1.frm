VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "ColorTunnel"
   ClientHeight    =   4710
   ClientLeft      =   3900
   ClientTop       =   3540
   ClientWidth     =   6990
   DrawWidth       =   2
   ForeColor       =   &H80000018&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   466
   Begin VB.CommandButton Command8 
      Caption         =   "Exit"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Narrower"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Direction"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Wider"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Sparser"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Smaller"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Denser"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Bigger"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************
' Name: Mesmerizing gradient circle
' Description:Displays a Mesmerizing and Hypnotic Effect
' using an endless Loop of gradient colored circles that
' blend nicely
' By: Jose M. Lopez
'
'This code is copyrighted and has limited warranties.
'Please see http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=8618&lngWId=1 for details.
'
'Modified by FollowTheWatch55 on 7/4/04 - followthewatch55@yahoo.com
'**************************************

Dim r As Integer 'red
Dim g As Integer 'green
Dim b As Integer 'blue
Dim rr As Integer 'red placeholder
Dim gg As Integer 'green placeholder
Dim bb As Integer 'blue placeholder
Dim rd As Single 'radius
Dim rs As Integer 'red +/- sign
Dim gs As Integer 'green +/- sign
Dim bs As Integer 'blue +/- sign
Dim rrs As Integer 'red placeholder +/- sign
Dim ggs As Integer 'green placeholder +/- sign
Dim bbs As Integer 'blue placeholder +/- sign
Dim cbf As Integer 'circle backwards of forward
Dim CircleSize As Integer
Dim CircleRadius As Single
Dim x As Long
Dim y As Long

Private Sub Command1_Click()
Me.Cls
CircleSize = CircleSize + 1
End Sub

Private Sub Command2_Click()
Me.Cls
CircleRadius = CircleRadius - 0.01
End Sub

Private Sub Command3_Click()
Me.Cls
CircleSize = CircleSize - 1
End Sub

Private Sub Command4_Click()
Me.Cls
CircleRadius = CircleRadius + 0.01
End Sub

Private Sub Command5_Click()
If DrawWidth = 1 Then Exit Sub
Me.Cls
DrawWidth = DrawWidth - 1
End Sub

Private Sub Command6_Click()
If DrawWidth = 32767 Then Exit Sub
Me.Cls
DrawWidth = DrawWidth + 1
End Sub

Private Sub Command7_Click()
'Select which way circles go
    cbf = cbf + 1


    If cbf = 2 Then
        cbf = 0
        cstp = 1
    End If


    If cbf = 0 Then
        rtnCircleForward
    Else
        rtnCircleBackward
    End If
End Sub

Private Sub Command8_Click()
End
End Sub

Private Sub Form_Load()
    Form1.Left = (Screen.Width / 2) - (Form1.Width / 2)
    Form1.Top = (Screen.Height / 2) - (Form1.Height / 2)
    Form1.Caption = "ColorTunnel " + LTrim(Str$(App.Major)) + "." + LTrim(Str$(App.Minor))
    Form1.Show
    CircleSize = 100
    CircleRadius = 1
    cbf = 0 'initialize backwards of forwards var
    rtnCircleForward
End Sub

Private Sub rtnCircleForward()

    r = 0: g = 0: b = 0
    rs = 1: gs = 1: bs = 1 'initial sign positive
    x = Form1.ScaleWidth / 2: y = Form1.ScaleHeight / 2 'Center
lblLoop:


    For rd = 1 To CircleSize Step CircleRadius
        'Load Color Placeholders so that each time the For rd loop is reinitialized
        'the colors at beginning of the For rd are almost the same as the
        'previous For rd. They will be loaded back to the rgb with one increment
        'at the end of the For rd

        If rd = 1 Then
            rr = r: gg = g: bb = b: rrs = rs: ggs = gs: bbs = bs
        End If

        
        rtnColors 'increment colors
        


        Form1.Circle (x, y), rd, RGB(r, g, b) 'make one circle


            DoEvents 'Don't want To Get stuck here
            Next rd

            'load back rgb and then increment them
            r = rr: g = gg: b = bb: rs = rrs: gs = ggs: bs = bbs
            rtnColors
            GoTo lblLoop
        End Sub


Private Sub rtnCircleBackward()

    'Basically the same as above but in reverse
    r = 0: g = 0: b = 0
    rs = 1: gs = 1: bs = 1
    x = Form1.ScaleWidth / 2: y = Form1.ScaleHeight / 2
lblLoop:


    For rd = CircleSize To 1 Step -CircleRadius


        If rd = CircleSize Then
            rr = r: gg = g: bb = b: rrs = rs: ggs = gs: bbs = bs
        End If

        rtnColors
        


        Form1.Circle (x, y), rd, RGB(r, g, b)
        
            DoEvents
            Next rd

            r = rr: g = gg: b = bb: rs = rrs: gs = ggs: bs = bbs
            rtnColors
            GoTo lblLoop
        End Sub


Private Sub rtnColors()

    'increment colors
    r = r + (5 * rs)
    If r = 255 Then rs = -1 'reached max, go the other way
    If r = 5 Then rs = 1 'reached min, go the other way
    g = g + (2 * gs)
    If g = 254 Then gs = -1
    If g = 2 Then gs = 1
    b = b + (3 * bs)
    If b = 252 Then bs = -1
    If b = 3 Then bs = 1
End Sub

Private Sub Form_Resize()
Me.Cls
x = Form1.ScaleWidth / 2: y = Form1.ScaleHeight / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

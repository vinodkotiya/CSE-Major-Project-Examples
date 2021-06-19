VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tanks"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer timerreload 
      Index           =   8
      Interval        =   1
      Left            =   7680
      Top             =   240
   End
   Begin VB.Timer timerreload 
      Index           =   7
      Interval        =   1
      Left            =   7200
      Top             =   240
   End
   Begin VB.Timer timerreload 
      Index           =   6
      Interval        =   1
      Left            =   6720
      Top             =   240
   End
   Begin VB.Timer timerreload 
      Index           =   5
      Interval        =   1
      Left            =   6240
      Top             =   240
   End
   Begin VB.Frame framewin 
      Height          =   1335
      Left            =   3120
      TabIndex        =   5
      Top             =   3840
      Visible         =   0   'False
      Width           =   4815
      Begin VB.PictureBox picwinner 
         AutoRedraw      =   -1  'True
         Height          =   975
         Left            =   120
         ScaleHeight     =   915
         ScaleWidth      =   915
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Continue"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         TabIndex        =   6
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblwin 
         Alignment       =   2  'Center
         Caption         =   "Player 1 Wins"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Timer timeraichangetarget 
      Interval        =   3000
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer timerreload 
      Index           =   4
      Interval        =   1
      Left            =   5760
      Top             =   240
   End
   Begin VB.Timer timerreload 
      Index           =   3
      Interval        =   1
      Left            =   5280
      Top             =   240
   End
   Begin VB.Timer timerreload 
      Index           =   2
      Interval        =   1
      Left            =   4800
      Top             =   240
   End
   Begin VB.Timer timerreload 
      Index           =   1
      Interval        =   1
      Left            =   4320
      Top             =   240
   End
   Begin VB.PictureBox picshells 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picshellsm 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   3
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer timerturn 
      Interval        =   35
      Left            =   120
      Top             =   6720
   End
   Begin VB.Timer timerfps 
      Interval        =   500
      Left            =   120
      Top             =   6240
   End
   Begin VB.PictureBox picback 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3840
      Left            =   3360
      ScaleHeight     =   3840
      ScaleWidth      =   6000
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.Timer timerrefresh 
      Interval        =   1
      Left            =   3600
      Top             =   240
   End
   Begin VB.PictureBox pictanksm 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1695
      ScaleWidth      =   2175
      TabIndex        =   1
      Top             =   2520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.PictureBox pictanks 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1695
      ScaleWidth      =   2175
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim leavenow As Byte
Dim fps As Long
Dim sfps As Long


Private Sub Command1_Click()
    For a = 1 To 8
      lefton(a) = 0
      righton(a) = 0
      upon(a) = 0
      downon(a) = 0
      ps(a) = 0
      pbs(a) = 0
      pbdir(a) = 0
      pfire(a) = 0
    Next a
    For a = 0 To ns
      sdir(a) = 0
    Next a
    
    frmmenu.Visible = True
    frmmain.Visible = False
 
    Unload frmmain
    Exit Sub

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If framewin.Visible = True Then Exit Sub
If KeyCode = 27 Then leavenow = 1
If pdir(1) > 0 And pstate(1) = 2 Then
  If KeyCode = 37 Then lefton(1) = 1
  If KeyCode = 38 Then upon(1) = 1
  If KeyCode = 39 Then righton(1) = 1
  If KeyCode = 40 Then downon(1) = 1
  If KeyCode = 13 Then pfire(1) = 1
End If
If pdir(2) > 0 And pstate(2) = 2 Then
  If KeyCode = 65 Then lefton(2) = 1
  If KeyCode = 87 Then upon(2) = 1
  If KeyCode = 68 Then righton(2) = 1
  If KeyCode = 83 Then downon(2) = 1
  If KeyCode = 32 Then pfire(2) = 1
End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If framewin.Visible = True Then Exit Sub
If pdir(1) > 0 Then
  If KeyCode = 37 Then lefton(1) = 0
  If KeyCode = 38 Then upon(1) = 0
  If KeyCode = 39 Then righton(1) = 0
  If KeyCode = 40 Then downon(1) = 0
  If KeyCode = 13 Then pfire(1) = 0
End If
If pdir(2) > 0 Then
  If KeyCode = 65 Then lefton(2) = 0
  If KeyCode = 87 Then upon(2) = 0
  If KeyCode = 68 Then righton(2) = 0
  If KeyCode = 83 Then downon(2) = 0
  If KeyCode = 32 Then pfire(2) = 0
End If
End Sub

Private Sub Form_Load()
  frmmain.Left = (Screen.Width - frmmain.Width) / 2
  frmmain.Top = (Screen.Height - frmmain.Height) / 2
  framewin.Left = (frmmain.Width - framewin.Width) / 2
  framewin.Top = (frmmain.Height - framewin.Height) / 2
  
  pxpos(1) = 200
  pypos(1) = 100
  pxpos(2) = 600
  pypos(2) = 101
  pxpos(3) = 201
  pypos(3) = 500
  pxpos(4) = 601
  pypos(4) = 501
  pxpos(5) = 300
  pypos(5) = 100
  pxpos(6) = 500
  pypos(6) = 101
  pxpos(7) = 301
  pypos(7) = 500
  pxpos(8) = 501
  pypos(8) = 501
  leavenow = 0
  fps = 0
  picback.Width = 800 * 15
  picback.Height = 600 * 15
  pictanks.Picture = LoadPicture("tankso.bmp")
  pictanksm.Picture = LoadPicture("tanks.bmp")
  picshells.Picture = LoadPicture("shell.bmp")
  picshellsm.Picture = LoadPicture("shellm.bmp")
  If dback = 1 Then picback.Picture = LoadPicture("back.bmp")
  For a = 0 To ns
    sdir(a) = 0
  Next a
  For a = 1 To 8
    preloaded(a) = 1
    ps(a) = 0
    timerreload(a).Interval = reloadspeed
  Next a
  
  Call changetarget

End Sub

Private Sub timeraichangetarget_Timer()
If framewin.Visible = True Then Exit Sub
Call changetarget
End Sub

Private Sub timerfps_Timer()
If framewin.Visible = True Then Exit Sub
frmmain.Caption = "Tanks Game " & "- Frames per Second - " & fps * 2
If gofast = 0 Then
  If fps * 2 > 70 Then
    slowdown = slowdown + 200
  Else
    If fps * 2 < 60 Then
      If slowdown > 10 Then slowdown = slowdown - 200
    End If
  End If
Else
  slowdown = 10
End If
fps = 0

End Sub

Private Sub timerrefresh_Timer()
'*** This is the Main Loop in the game ***
If framewin.Visible = True Then Exit Sub 'Leave if game has ended
Do 'start main loop
  fps = fps + 1 'used to calculate Frames per second
  
  '**** Draws Tanks on screen ****
  For a = 1 To 8
    If pdir(a) > 0 Then '0 if not in game
      success = BitBlt(picback.hDC, pxpos(a) - 29, pypos(a) - 29, 58, 58, pictanks.hDC, (pdir(a) - 1) * 58, (a - 1) * 58, SRCAND)
      success = BitBlt(picback.hDC, pxpos(a) - 29, pypos(a) - 29, 58, 58, pictanksm.hDC, (pdir(a) - 1) * 58, (a - 1) * 58, SRCPAINT)
    End If
  Next a
 
  '**** Ends if all but one tank has been destroyed
  If pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then
    lblwin.Caption = "It was a Draw"
    framewin.Visible = True
    picwinner.SetFocus
    Exit Sub
  ElseIf pdir(1) >= 1 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then
    lblwin.Caption = "Player 1 Wins"
    success = BitBlt(picwinner.hDC, 0, 0, picwinner.Width / 15, picwinner.Height / 15, pictanks.hDC, 0, 0, SRCCOPY)
    framewin.Visible = True
    picwinner.SetFocus
    Exit Sub
  ElseIf pdir(1) = 0 And pdir(2) >= 1 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then
    lblwin.Caption = "Player 2 Wins"
    success = BitBlt(picwinner.hDC, 0, 0, picwinner.Width / 15, picwinner.Height / 15, pictanks.hDC, 0, 58, SRCCOPY)
    framewin.Visible = True
    picwinner.SetFocus
    Exit Sub
  ElseIf pdir(1) = 0 And pdir(2) = 0 And pdir(3) >= 1 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then
    lblwin.Caption = "Player 3 Wins"
    success = BitBlt(picwinner.hDC, 0, 0, picwinner.Width / 15, picwinner.Height / 15, pictanks.hDC, 0, 116, SRCCOPY)
    framewin.Visible = True
    picwinner.SetFocus
    Exit Sub
  ElseIf pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) >= 1 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then
    lblwin.Caption = "Player 4 Wins"
    success = BitBlt(picwinner.hDC, 0, 0, picwinner.Width / 15, picwinner.Height / 15, pictanks.hDC, 0, 174, SRCCOPY)
    framewin.Visible = True
    picwinner.SetFocus
    Exit Sub
  ElseIf pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) >= 1 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then
    lblwin.Caption = "Player 5 Wins"
    success = BitBlt(picwinner.hDC, 0, 0, picwinner.Width / 15, picwinner.Height / 15, pictanks.hDC, 0, 232, SRCCOPY)
    framewin.Visible = True
    picwinner.SetFocus
    Exit Sub
  ElseIf pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) >= 1 And pdir(7) = 0 And pdir(8) = 0 Then
    lblwin.Caption = "Player 6 Wins"
    success = BitBlt(picwinner.hDC, 0, 0, picwinner.Width / 15, picwinner.Height / 15, pictanks.hDC, 0, 290, SRCCOPY)
    framewin.Visible = True
    picwinner.SetFocus
    Exit Sub
  ElseIf pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) >= 1 And pdir(8) = 0 Then
    lblwin.Caption = "Player 7 Wins"
    success = BitBlt(picwinner.hDC, 0, 0, picwinner.Width / 15, picwinner.Height / 15, pictanks.hDC, 0, 348, SRCCOPY)
    framewin.Visible = True
    picwinner.SetFocus
    Exit Sub
  ElseIf pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) >= 1 Then
    lblwin.Caption = "Player 8 Wins"
    success = BitBlt(picwinner.hDC, 0, 0, picwinner.Width / 15, picwinner.Height / 15, pictanks.hDC, 0, 406, SRCCOPY)
    framewin.Visible = True
    picwinner.SetFocus
    Exit Sub
  End If
   
  '**** Speeds up or slows down tanks ****
  For a = 1 To 8
    If upon(a) = 1 And downon(a) = 0 Then
      If ps(a) < 101 Then ps(a) = ps(a) + 1
    ElseIf downon(a) = 1 And upon(a) = 0 Then
      If ps(a) > -51 Then ps(a) = ps(a) - 1
    End If
  Next a
  
  '**** Calls Sub in module to move tanks ****
  Call movetanks
  
  '**** Detects collisions with edges of screen ****
  For a = 1 To 8
    If pypos(a) < 29 Or pxpos(a) < 29 Or pypos(a) > 600 - 29 Or pxpos(a) > 800 - 29 Then
       If pypos(a) < 29 Then pypos(a) = 29
       If pxpos(a) < 29 Then pxpos(a) = 29
       If pypos(a) > 600 - 29 Then pypos(a) = 600 - 29
       If pxpos(a) > 800 - 29 Then pxpos(a) = 800 - 29
       pbs(a) = -ps(a)
       ps(a) = 0
    End If
  Next a
  
  '**** Detect collisions with other tanks ****
  For b = 1 To 8
    For a = 1 To 8
      If a <> b And pdir(a) > 0 And pdir(b) > 0 Then
        If pxpos(b) > pxpos(a) - 40 And pxpos(b) < pxpos(a) + 40 And pypos(b) > pypos(a) - 40 And pypos(b) < pypos(a) + 40 Then
      '    pbdir(b) = pdir(b)
          pbs(b) = -ps(b)
          ps(b) = 0
          pypos(b) = ptypos(b)
          pxpos(b) = ptxpos(b)
        Else
          ptxpos(b) = pxpos(b)
          ptypos(b) = pypos(b)
        End If
      End If
    Next a
  Next b
  
  '**** Find spare shell entry and registers it ****
  For b = 1 To 8
    If pfire(b) = 1 And preloaded(b) = 1 Then
      For a = 0 To ns
        If sdir(a) = 0 Then
          sdir(a) = pdir(b)
          sxpos(a) = pxpos(b)
          sypos(a) = pypos(b)
          ss(a) = ps(b) + 180
          sown(a) = b
          preloaded(b) = 0
          timerreload(b).Enabled = True
          GoTo foundone
        End If
      Next a
    End If
foundone:
  Next b
      
  '**** draw shells to screen and deallocate those off screen,
  '     Also detects hits on tanks
  For a = 0 To ns
    If sdir(a) > 0 Then
      Call moveshells(a) 'moves shells
       For b = 1 To 8
        If b <> sown(a) And pdir(b) > 0 Then
          If sxpos(a) > pxpos(b) - 22 And sxpos(a) < pxpos(b) + 22 And sypos(a) > pypos(b) - 22 And sypos(a) < pypos(b) + 22 Then
            sdir(a) = 0
            ph(b) = ph(b) - 1
            If ph(b) = 0 Then
              pdir(b) = 0
              ps(b) = 0
            End If
          End If
        End If
      Next b
      
      success = BitBlt(picback.hDC, sxpos(a) - 2, sypos(a) - 2, 4, 4, picshells.hDC, 0, 0, SRCAND)
      success = BitBlt(picback.hDC, sxpos(a) - 2, sypos(a) - 2, 4, 4, picshellsm.hDC, 0, 0, SRCPAINT)
      'deallocate if off screen
      If sxpos(a) < 0 Or sxpos(a) > 800 Or sypos(a) < 0 Or sypos(a) > 600 Then sdir(a) = 0
    End If
  Next a
  'slow done after bounce
  For a = 1 To 8
    If pbs(a) > 0 Then pbs(a) = pbs(a) - 1
    If pbs(a) < 0 Then pbs(a) = pbs(a) + 1
  Next a
  'copy all to display
  
  success = BitBlt(frmmain.hDC, 0, 0, 800, 600, picback.hDC, 0, 0, SRCCOPY)
  
  picback.Cls 'clear for next frame
   
  Call aicontrol 'AI Control
  
  For a = 1 To slowdown
    DoEvents 'yield to system
  Next

  If leavenow = 1 Then 'leave if requested by user (ESC key)
    frmmenu.Visible = True
    frmmain.Visible = False
   
    Unload frmmain
    Exit Sub
  End If
Loop
End Sub

Private Sub timerreload_Timer(Index As Integer)
preloaded(Index) = 1
timerreload(Index).Enabled = False

End Sub

Private Sub timerturn_Timer()
If framewin.Visible = True Then Exit Sub
  For a = 1 To 8
    If lefton(a) = 1 And righton(a) = 0 And pdir(a) > 0 Then
      pdir(a) = pdir(a) - 1
      If pdir(a) = 0 Then pdir(a) = 36
    ElseIf righton(a) = 1 And lefton(a) = 0 And pdir(a) > 0 Then
      pdir(a) = pdir(a) + 1
      If pdir(a) = 37 Then pdir(a) = 1
    End If
  Next a
  

End Sub

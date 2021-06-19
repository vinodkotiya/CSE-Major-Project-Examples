VERSION 5.00
Begin VB.UserControl UserControl1 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4635
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   FillStyle       =   0  'Solid
   ScaleHeight     =   309
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   456
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture10 
      AutoRedraw      =   -1  'True
      Height          =   1335
      Left            =   0
      Picture         =   "METEOR.ctx":0000
      ScaleHeight     =   1275
      ScaleWidth      =   6795
      TabIndex        =   10
      Top             =   3240
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.PictureBox Picture9 
      AutoRedraw      =   -1  'True
      Height          =   1335
      Left            =   0
      Picture         =   "METEOR.ctx":9BAA
      ScaleHeight     =   1275
      ScaleWidth      =   6795
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   6855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   3720
      Width           =   1335
   End
   Begin VB.PictureBox Picture8 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   4440
      Picture         =   "METEOR.ctx":13754
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture7 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   3480
      Picture         =   "METEOR.ctx":1479E
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture6 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   2520
      Picture         =   "METEOR.ctx":157E8
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   1560
      Picture         =   "METEOR.ctx":16832
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   4440
      Picture         =   "METEOR.ctx":1787C
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   3480
      Picture         =   "METEOR.ctx":18856
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   2520
      Picture         =   "METEOR.ctx":19830
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   855
      Left            =   1560
      Picture         =   "METEOR.ctx":1A87A
      ScaleHeight     =   53
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   53
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' ***   Main UserControl1 codes     ****
Sub EARTH()
     Dim SUCCESS As Integer
     SUCCESS = BitBlt(hDC, 0, 225, Picture9.Width, Picture9.Height, Picture9.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(hDC, 0, 225, Picture10.Width, Picture10.Height, Picture10.hDC, 0, 0, SRCPAINT)
End Sub

Sub PLANEX()
     Dim SUCCESS As Integer
     SUCCESS = BitBlt(hDC, XPLANE, YPLANE, Picture1.Width, Picture1.Height, Picture1.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(hDC, XPLANE, YPLANE, Picture2.Width, Picture2.Height, Picture2.hDC, 0, 0, SRCPAINT)
End Sub

Sub BULLETX()
     Dim SUCCESS As Integer
     YBULLET = YBULLET - 50
     If YBULLET < 0 Then
          If MOUSESW = 1 Then
               XBULLET = XPLANE
               YBULLET = YPLANE
          Else
               FIRESW = 0
          End If
     End If
     SUCCESS = BitBlt(hDC, XBULLET, YBULLET, Picture5.Width, Picture5.Height, Picture5.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(hDC, XBULLET, YBULLET, Picture6.Width, Picture6.Height, Picture6.hDC, 0, 0, SRCPAINT)
End Sub
Sub DELAYX()
     Dim DELAY As Integer
     For DELAY = 1 To 30000
     Next DELAY
End Sub
Sub EXPLODE(XPOS As Integer, YPOS As Integer)
     Dim SUCCESS As Integer
     SUCCESS = BitBlt(hDC, XPOS, YPOS, Picture7.Width, Picture7.Height, Picture7.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(hDC, XPOS, YPOS, Picture8.Width, Picture8.Height, Picture8.hDC, 0, 0, SRCPAINT)
End Sub
' Background balls movement
Sub BACKGROUND()
     Dim CTR As Integer
     BACKGROUNDCTR = BACKGROUNDCTR + 3
     If BACKGROUNDCTR > 300 Then BACKGROUNDCTR = 0
     For CTR = 1 To 30
          Circle (STAR(CTR), (CTR * 10) - 300 + BACKGROUNDCTR), 1
          Circle (STAR(CTR), CTR * 10 + BACKGROUNDCTR), 1
     Next CTR
End Sub

Sub INITIALIZEX()
     Timer1.Enabled = False
     Command1.Visible = True
     YMETEOR1 = 0
     YMETEOR2 = 0
     YMETEOR3 = 0
     YMETEOR4 = 0
     YMETEOR5 = 0
     GAMECTR = 0
     BACKGROUNDCTR = 0
     SCORECTR = 0
     LIFECTR = 3
End Sub

Sub STAGE1()
     YMETEOR1 = YMETEOR1 + 6
     COLLIDESW = COLLIDE(XPLANE, YPLANE, XMETEOR1, YMETEOR1, 75)
     HITSW = COLLIDE(XBULLET, YBULLET, XMETEOR1, YMETEOR1, 75)
     If YMETEOR1 > 250 Then
          Call EXPLODE(XMETEOR1, YMETEOR1)
          Call DELAYX
          LIFECTR = LIFECTR - 1
          XMETEOR1 = (Rnd * 400)
          YMETEOR1 = 0
     End If
     SUCCESS = BitBlt(hDC, XMETEOR1, YMETEOR1, Picture3.Width, Picture3.Height, Picture3.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(hDC, XMETEOR1, YMETEOR1, Picture4.Width, Picture4.Height, Picture4.hDC, 0, 0, SRCPAINT)
     If COLLIDESW = 1 Then
          Call EXPLODE(XPLANE, YPLANE)
          Call EXPLODE(XMETEOR1, YMETEOR1)
          Call DELAYX
          SCORECTR = SCORECTR + 1
          LIFECTR = LIFECTR - 1
          XMETEOR1 = (Rnd * 400)
          YMETEOR1 = 0
     End If
     If FIRESW = 1 Then
          If HITSW = 1 Then
               Call EXPLODE(XMETEOR1, YMETEOR1)
               Call DELAYX
               If MOUSESW = 1 Then
                    XBULLET = XPLANE
                    YBULLET = YPLANE
               Else
                    FIRESW = 0
               End If
               SCORECTR = SCORECTR + 1
               XMETEOR1 = (Rnd * 400)
               YMETEOR1 = 0
           End If
      End If
End Sub

Sub STAGE2()
     YMETEOR2 = YMETEOR2 + 6
     COLLIDESW = COLLIDE(XPLANE, YPLANE, XMETEOR2, YMETEOR2, 75)
     HITSW = COLLIDE(XBULLET, YBULLET, XMETEOR2, YMETEOR2, 75)
     If YMETEOR2 > 250 Then
          Call EXPLODE(XMETEOR2, YMETEOR2)
          Call DELAYX
          LIFECTR = LIFECTR - 1
          XMETEOR2 = (Rnd * 400)
          YMETEOR2 = 0
     End If
     SUCCESS = BitBlt(hDC, XMETEOR2, YMETEOR2, Picture3.Width, Picture3.Height, Picture3.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(hDC, XMETEOR2, YMETEOR2, Picture4.Width, Picture4.Height, Picture4.hDC, 0, 0, SRCPAINT)
     If COLLIDESW = 1 Then
          Call EXPLODE(XPLANE, YPLANE)
          Call EXPLODE(XMETEOR2, YMETEOR2)
          Call DELAYX
          SCORECTR = SCORECTR + 2
          LIFECTR = LIFECTR - 1
          XMETEOR2 = (Rnd * 400)
          YMETEOR2 = 0
     End If
     If FIRESW = 1 Then
          If HITSW = 1 Then
               Call EXPLODE(XMETEOR2, YMETEOR2)
               Call DELAYX
               If MOUSESW = 1 Then
                    XBULLET = XPLANE
                    YBULLET = YPLANE
               Else
                    FIRESW = 0
               End If
               SCORECTR = SCORECTR + 2
               XMETEOR2 = (Rnd * 400)
               YMETEOR2 = 0
           End If
      End If
End Sub
Sub STAGE3()
     YMETEOR3 = YMETEOR3 + 6
     COLLIDESW = COLLIDE(XPLANE, YPLANE, XMETEOR3, YMETEOR3, 75)
     HITSW = COLLIDE(XBULLET, YBULLET, XMETEOR3, YMETEOR3, 75)
     If YMETEOR3 > 250 Then
          Call EXPLODE(XMETEOR3, YMETEOR3)
          Call DELAYX
          LIFECTR = LIFECTR - 1
          XMETEOR3 = (Rnd * 400)
          YMETEOR3 = 0
     End If
     SUCCESS = BitBlt(hDC, XMETEOR3, YMETEOR3, Picture3.Width, Picture3.Height, Picture3.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(hDC, XMETEOR3, YMETEOR3, Picture4.Width, Picture4.Height, Picture4.hDC, 0, 0, SRCPAINT)
     If COLLIDESW = 1 Then
          Call EXPLODE(XPLANE, YPLANE)
          Call EXPLODE(XMETEOR3, YMETEOR3)
          Call DELAYX
          SCORECTR = SCORECTR + 3
          LIFECTR = LIFECTR - 1
          XMETEOR3 = (Rnd * 400)
          YMETEOR3 = 0
     End If
     If FIRESW = 1 Then
          If HITSW = 1 Then
               Call EXPLODE(XMETEOR3, YMETEOR3)
               Call DELAYX
               If MOUSESW = 1 Then
                    XBULLET = XPLANE
                    YBULLET = YPLANE
               Else
                    FIRESW = 0
               End If
               SCORECTR = SCORECTR + 3
               XMETEOR3 = (Rnd * 400)
               YMETEOR3 = 0
           End If
      End If
End Sub

Sub STAGE4()
     YMETEOR4 = YMETEOR4 + 6
     COLLIDESW = COLLIDE(XPLANE, YPLANE, XMETEOR4, YMETEOR4, 75)
     HITSW = COLLIDE(XBULLET, YBULLET, XMETEOR4, YMETEOR4, 75)
     If YMETEOR4 > 250 Then
          Call EXPLODE(XMETEOR4, YMETEOR4)
          Call DELAYX
          LIFECTR = LIFECTR - 1
          XMETEOR4 = (Rnd * 400)
          YMETEOR4 = 0
     End If
     SUCCESS = BitBlt(hDC, XMETEOR4, YMETEOR4, Picture3.Width, Picture3.Height, Picture3.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(hDC, XMETEOR4, YMETEOR4, Picture4.Width, Picture4.Height, Picture4.hDC, 0, 0, SRCPAINT)
     If COLLIDESW = 1 Then
          Call EXPLODE(XPLANE, YPLANE)
          Call EXPLODE(XMETEOR4, YMETEOR4)
          Call DELAYX
          SCORECTR = SCORECTR + 4
          LIFECTR = LIFECTR - 1
          XMETEOR4 = (Rnd * 400)
          YMETEOR4 = 0
     End If
     If FIRESW = 1 Then
           If HITSW = 1 Then
               Call EXPLODE(XMETEOR4, YMETEOR4)
               Call DELAYX
               If MOUSESW = 1 Then
                    XBULLET = XPLANE
                    YBULLET = YPLANE
               Else
                    FIRESW = 0
               End If
               SCORECTR = SCORECTR + 4
               XMETEOR4 = (Rnd * 400)
               YMETEOR4 = 0
           End If
      End If
End Sub

Sub STAGE5()
     YMETEOR5 = YMETEOR5 + 6
     COLLIDESW = COLLIDE(XPLANE, YPLANE, XMETEOR5, YMETEOR5, 75)
     HITSW = COLLIDE(XBULLET, YBULLET, XMETEOR5, YMETEOR5, 75)
     If YMETEOR5 > 250 Then
          Call EXPLODE(XMETEOR5, YMETEOR5)
          Call DELAYX
          LIFECTR = LIFECTR - 1
          XMETEOR5 = (Rnd * 400)
          YMETEOR5 = 0
     End If
     SUCCESS = BitBlt(hDC, XMETEOR5, YMETEOR5, Picture3.Width, Picture3.Height, Picture3.hDC, 0, 0, SRCAND)
     SUCCESS = BitBlt(hDC, XMETEOR5, YMETEOR5, Picture4.Width, Picture4.Height, Picture4.hDC, 0, 0, SRCPAINT)
     If COLLIDESW = 1 Then
          Call EXPLODE(XPLANE, YPLANE)
          Call EXPLODE(XMETEOR5, YMETEOR5)
          Call DELAYX
          SCORECTR = SCORECTR + 5
          LIFECTR = LIFECTR - 1
          XMETEOR5 = (Rnd * 400)
          YMETEOR5 = 0
     End If
     If FIRESW = 1 Then
          If HITSW = 1 Then
               Call EXPLODE(XMETEOR5, YMETEOR5)
               Call DELAYX
               If MOUSESW = 1 Then
                    XBULLET = XPLANE
                    YBULLET = YPLANE
               Else
                    FIRESW = 0
               End If
               SCORECTR = SCORECTR + 5
               XMETEOR5 = (Rnd * 400)
               YMETEOR5 = 0
           End If
     End If
End Sub

Private Sub Command1_Click()
     Command1.Visible = False
     Timer1.Enabled = True
End Sub
' Timer control project calling different STAGEs

Private Sub Timer1_Timer()
     Cls
     GAMECTR = GAMECTR + 1
     Call BACKGROUND
     Call EARTH
     If FIRESW = 1 Then Call BULLETX
     Call PLANEX
     Call STAGE1
     If GAMECTR > 250 Then Call STAGE2
     If GAMECTR > 500 Then Call STAGE3
     If GAMECTR > 750 Then Call STAGE4
     If GAMECTR > 1000 Then Call STAGE5
     CurrentX = 0
     CurrentY = 0
     Print "ActiveX Meteor Game by Gagan Sahoo"
     Print "SCORE - " + Str$(SCORECTR)
     Print "LIFE - " + Str$(LIFECTR)
     If GAMECTR > 10000 Or LIFECTR <= 0 Then
          Call INITIALIZEX
          CurrentX = 195
          CurrentY = 125
          Print "Game Over"
     End If
End Sub

' Star War projects initializes here to start
Private Sub UserControl_Initialize()
     Dim CTR As Integer
     Randomize Timer
     FillColor = QBColor(15)
     ForeColor = QBColor(15)
     BackColor = QBColor(0)
     Call INITIALIZEX
     For CTR = 1 To 30
          STAR(CTR) = (Rnd * 450)
     Next CTR
     Call BACKGROUND
     Call EARTH
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     FIRESW = 1
     MOUSESW = 1
     XBULLET = XPLANE
     YBULLET = YPLANE
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     XPLANE = X - 25
     YPLANE = Y - 25
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     MOUSESW = 0
End Sub

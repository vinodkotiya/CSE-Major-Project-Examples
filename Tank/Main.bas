Attribute VB_Name = "Main"
' Main module of the game
Global armour As Integer
Global gofast As Byte 'moderate speed or not
Global dback As Byte 'Show background picture
Global pmusic As Byte 'play music
Global reloadspeed As Integer
Global pstate(8) As Integer 'is player computer or human
Global a As Long 'User in For... Next loops etc
Global b As Long 'User in For... Next loops etc
Global ran As Single ' for randomize timer
Global ran2 As Single
Global lefton(8) As Byte
Global righton(8) As Byte
Global upon(8) As Byte
Global downon(8) As Byte
Global slowdown As Long ' Used to run game at correct speed on all computers
Global ns As Integer 'number of shells
Global pfire(8) As Byte 'Unit fire
Global pxpos(8) As Single 'Tank x coordinates
Global pypos(8) As Single 'Tank y coordinates
Global ptxpos(8) As Single 'Tank x coordinates (temp)
Global ptypos(8) As Single 'Tank y coordinates (temp)
Global pdir(8) As Integer 'Tank Direction
Global pbdir(8) As Integer 'Bounce Direction
Global preloaded(8) As Byte 'Has tank reloaded?
Global ps(8) As Single 'Tank Speed
Global pbs(8) As Single 'Bounce Speed
Global ph(8) As Integer 'Tank Health
Global pt(8) As Integer 'Currant target (for AI)
Global sdir(1500) As Integer 'Shell direction starts 0, techniqually you can have 100 shells on screen at once - hey major slowdown
Global ss(1500) As Integer  'Shell speed - combination of tank speed and exit velocity
Global sxpos(1500) As Single 'shell x coord
Global sypos(1500) As Single 'shell y coord
Global sown(1500) As Integer 'Who fired shell?
Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (ByVal lpszSoundName As Any, ByVal uFlags As Long) As Long
Global Const SND_ASYNC = &H1     ' Play asynchronously
Global Const SND_NODEFAULT = &H2 ' Don't use default sound
Global Const SND_MEMORY = &H4    ' lpszSoundName points to a memory file

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const BLACKNESS = &H42 ' (DWORD) dest = BLACK
Public Const DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)
Public Const MERGECOPY = &HC000CA       ' (DWORD) dest = (source AND pattern)
Public Const MERGEPAINT = &HBB0226      ' (DWORD) dest = (NOT source) OR dest
Public Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
Public Const NOTSRCERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)
Public Const PATCOPY = &HF00021 ' (DWORD) dest = pattern
Public Const PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
Public Const PATPAINT = &HFB0A09        ' (DWORD) dest = DPSnoo
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Const WHITENESS = &HFF0062       ' (DWORD) dest = WHITE
Public Sub aicontrol()
 For a = 1 To 8
  If pdir(a) > 0 And pstate(a) = 1 Then
    If pxpos(pt(a)) < pxpos(a) And pypos(pt(a)) < pypos(a) Then
      If pdir(a) <= 32 And pdir(a) >= 14 Then
        lefton(a) = 0
        righton(a) = 1
      ElseIf pdir(a) < 14 Or pdir(a) > 33 Then
        righton(a) = 0
        lefton(a) = 1
      Else
        pfire(a) = 1
        righton(a) = 0
        lefton(a) = 0
      End If
    ElseIf pxpos(pt(a)) < pxpos(a) And pypos(pt(a)) > pypos(a) Then
      If pdir(a) <= 24 And pdir(a) >= 5 Then
        lefton(a) = 0
        righton(a) = 1
      ElseIf pdir(a) < 5 Or pdir(a) > 25 Then
        righton(a) = 0
        lefton(a) = 1
      Else
        pfire(a) = 1
        righton(a) = 0
        lefton(a) = 0
      End If
    ElseIf pxpos(pt(a)) > pxpos(a) And pypos(pt(a)) < pypos(a) Then
       If pdir(a) <= 5 Or pdir(a) >= 25 Then
        lefton(a) = 0
        righton(a) = 1
      ElseIf pdir(a) > 6 And pdir(a) < 25 Then
        righton(a) = 0
        lefton(a) = 1
      Else
        pfire(a) = 1
        righton(a) = 0
        lefton(a) = 0
      End If
    ElseIf pxpos(pt(a)) > pxpos(a) And pypos(pt(a)) > pypos(a) Then
       If pdir(a) < 15 Or pdir(a) >= 34 Then
        lefton(a) = 0
        righton(a) = 1
      ElseIf pdir(a) > 15 And pdir(a) < 34 Then
        righton(a) = 0
        lefton(a) = 1
      Else
        pfire(a) = 1
        righton(a) = 0
        lefton(a) = 0
      End If
    Else
      pfire(a) = 0
      righton(a) = 0
      lefton(a) = 0
    End If
    upon(a) = 1
  End If
Next a
End Sub

Public Sub changetarget()
  '**** Ends if all but one tank has been destroyed
  If pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then Exit Sub
  If pdir(1) >= 1 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then Exit Sub
  If pdir(1) = 0 And pdir(2) >= 1 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then Exit Sub
  If pdir(1) = 0 And pdir(2) = 0 And pdir(3) >= 1 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then Exit Sub
  If pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) >= 1 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then Exit Sub
  If pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) >= 1 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) = 0 Then Exit Sub
  If pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) >= 1 And pdir(7) = 0 And pdir(8) = 0 Then Exit Sub
  If pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) >= 1 And pdir(8) = 0 Then Exit Sub
  If pdir(1) = 0 And pdir(2) = 0 And pdir(3) = 0 And pdir(4) = 0 And pdir(5) = 0 And pdir(6) = 0 And pdir(7) = 0 And pdir(8) >= 1 Then Exit Sub
  

For a = 1 To 8 Step 1
  If pstate(a) = 1 And pdir(a) > 0 Then
doh:
    Randomize Timer
    ran = Rnd
     b = a + 1
     If b = 9 Then b = 1
    For ran2 = 0 To 0.875 Step 0.125
      If ran >= ran2 And ran < ran2 + 0.125 Then
        pt(a) = b
       End If
        b = b + 1
        If b = 9 Then b = 1
   
    Next ran2
  If pt(a) = a Then GoTo doh:
  If pdir(pt(a)) = 0 Then GoTo doh
    
  End If
Next a


End Sub
Public Sub movetanks()
  For a = 1 To 8
    If pdir(a) = 1 Then
      pypos(a) = pypos(a) - ps(a) / 20 - (pbs(a) / 20)
    End If
    If pdir(a) = 2 Then
      pypos(a) = pypos(a) - ps(a) / 22 - (pbs(a) / 22)
      pxpos(a) = pxpos(a) + ps(a) / 100 + (pbs(a) / 100)
    End If
    If pdir(a) = 3 Then
      pypos(a) = pypos(a) - ps(a) / 25 - (pbs(a) / 25)
      pxpos(a) = pxpos(a) + ps(a) / 60 + (pbs(a) / 60)
    End If
    If pdir(a) = 4 Then
      pypos(a) = pypos(a) - ps(a) / 27 - (pbs(a) / 27)
      pxpos(a) = pxpos(a) + ps(a) / 50 + (pbs(a) / 50)
    End If
    If pdir(a) = 5 Then
      pypos(a) = pypos(a) - ps(a) / 29 - (pbs(a) / 29)
      pxpos(a) = pxpos(a) + ps(a) / 35 + (pbs(a) / 35)
    End If
    If pdir(a) = 6 Then
      pypos(a) = pypos(a) - ps(a) / 35 - (pbs(a) / 35)
      pxpos(a) = pxpos(a) + ps(a) / 29 + (pbs(a) / 29)
    End If
    If pdir(a) = 7 Then
      pypos(a) = pypos(a) - ps(a) / 50 - (pbs(a) / 50)
      pxpos(a) = pxpos(a) + ps(a) / 27 + (pbs(a) / 27)
    End If
    If pdir(a) = 8 Then
      pypos(a) = pypos(a) - ps(a) / 60 - (pbs(a) / 60)
      pxpos(a) = pxpos(a) + ps(a) / 25 + (pbs(a) / 25)
    End If
    If pdir(a) = 9 Then
      pypos(a) = pypos(a) - ps(a) / 100 - (pbs(a) / 100)
      pxpos(a) = pxpos(a) + ps(a) / 22 + (pbs(a) / 22)
    End If
    If pdir(a) = 10 Then
      pxpos(a) = pxpos(a) + ps(a) / 20 + (pbs(a) / 20)
    End If
    If pdir(a) = 11 Then
      pypos(a) = pypos(a) + ps(a) / 100 + (pbs(a) / 100)
      pxpos(a) = pxpos(a) + ps(a) / 22 + (pbs(a) / 22)
    End If
    If pdir(a) = 12 Then
      pypos(a) = pypos(a) + ps(a) / 60 + (pbs(a) / 60)
      pxpos(a) = pxpos(a) + ps(a) / 25 + (pbs(a) / 25)
    End If
    If pdir(a) = 13 Then
      pypos(a) = pypos(a) + ps(a) / 50 + (pbs(a) / 50)
      pxpos(a) = pxpos(a) + ps(a) / 27 + (pbs(a) / 27)
    End If
    If pdir(a) = 14 Then
      pypos(a) = pypos(a) + ps(a) / 35 + (pbs(a) / 35)
      pxpos(a) = pxpos(a) + ps(a) / 29 + (pbs(a) / 29)
    End If
    If pdir(a) = 15 Then
      pypos(a) = pypos(a) + ps(a) / 29 + (pbs(a) / 29)
      pxpos(a) = pxpos(a) + ps(a) / 35 + (pbs(a) / 35)
    End If
    If pdir(a) = 16 Then
      pypos(a) = pypos(a) + ps(a) / 27 + (pbs(a) / 27)
      pxpos(a) = pxpos(a) + ps(a) / 50 + (pbs(a) / 50)
    End If
    If pdir(a) = 17 Then
      pypos(a) = pypos(a) + ps(a) / 25 + (pbs(a) / 25)
      pxpos(a) = pxpos(a) + ps(a) / 60 + (pbs(a) / 60)
    End If
    If pdir(a) = 18 Then
      pypos(a) = pypos(a) + ps(a) / 22 + (pbs(a) / 22)
      pxpos(a) = pxpos(a) + ps(a) / 100 + (pbs(a) / 100)
    End If
    If pdir(a) = 19 Then
      pypos(a) = pypos(a) + ps(a) / 20 + (pbs(a) / 20)
    End If
      If pdir(a) = 20 Then
      pypos(a) = pypos(a) + ps(a) / 22 + (pbs(a) / 22)
      pxpos(a) = pxpos(a) - ps(a) / 100 - (pbs(a) / 100)
    End If
    If pdir(a) = 21 Then
      pypos(a) = pypos(a) + ps(a) / 25 + (pbs(a) / 25)
      pxpos(a) = pxpos(a) - ps(a) / 60 - (pbs(a) / 60)
    End If
    If pdir(a) = 22 Then
      pypos(a) = pypos(a) + ps(a) / 27 + (pbs(a) / 27)
      pxpos(a) = pxpos(a) - ps(a) / 50 - (pbs(a) / 50)
    End If
    If pdir(a) = 23 Then
      pypos(a) = pypos(a) + ps(a) / 29 + (pbs(a) / 29)
      pxpos(a) = pxpos(a) - ps(a) / 35 - (pbs(a) / 35)
    End If
    If pdir(a) = 24 Then
      pypos(a) = pypos(a) + ps(a) / 35 + (pbs(a) / 35)
      pxpos(a) = pxpos(a) - ps(a) / 29 - (pbs(a) / 29)
    End If
    If pdir(a) = 25 Then
      pypos(a) = pypos(a) + ps(a) / 50 + (pbs(a) / 50)
      pxpos(a) = pxpos(a) - ps(a) / 27 - (pbs(a) / 27)
    End If
    If pdir(a) = 26 Then
      pypos(a) = pypos(a) + ps(a) / 60 + (pbs(a) / 60)
      pxpos(a) = pxpos(a) - ps(a) / 25 - (pbs(a) / 25)
    End If
    If pdir(a) = 27 Then
      pypos(a) = pypos(a) + ps(a) / 100 + (pbs(a) / 100)
      pxpos(a) = pxpos(a) - ps(a) / 22 - (pbs(a) / 22)
    End If
    If pdir(a) = 28 Then
      pxpos(a) = pxpos(a) - ps(a) / 20 - (pbs(a) / 20)
    End If
    If pdir(a) = 29 Then
      pypos(a) = pypos(a) - ps(a) / 100 - (pbs(a) / 100)
      pxpos(a) = pxpos(a) - ps(a) / 22 - (pbs(a) / 22)
    End If
    If pdir(a) = 30 Then
      pypos(a) = pypos(a) - ps(a) / 60 - (pbs(a) / 60)
      pxpos(a) = pxpos(a) - ps(a) / 25 - (pbs(a) / 25)
    End If
    If pdir(a) = 31 Then
      pypos(a) = pypos(a) - ps(a) / 50 - (pbs(a) / 50)
      pxpos(a) = pxpos(a) - ps(a) / 27 - (pbs(a) / 27)
    End If
    If pdir(a) = 32 Then
      pypos(a) = pypos(a) - ps(a) / 35 - (pbs(a) / 35)
      pxpos(a) = pxpos(a) - ps(a) / 29 - (pbs(a) / 29)
    End If
    If pdir(a) = 33 Then
      pypos(a) = pypos(a) - ps(a) / 29 - (pbs(a) / 29)
      pxpos(a) = pxpos(a) - ps(a) / 35 - (pbs(a) / 35)
    End If
    If pdir(a) = 34 Then
      pypos(a) = pypos(a) - ps(a) / 27 - (pbs(a) / 27)
      pxpos(a) = pxpos(a) - ps(a) / 50 - (pbs(a) / 50)
    End If
    If pdir(a) = 35 Then
      pypos(a) = pypos(a) - ps(a) / 25 - (pbs(a) / 25)
      pxpos(a) = pxpos(a) - ps(a) / 60 - (pbs(a) / 60)
    End If
    If pdir(a) = 36 Then
      pypos(a) = pypos(a) - ps(a) / 22 - (pbs(a) / 22)
      pxpos(a) = pxpos(a) - ps(a) / 100 - (pbs(a) / 100)
    End If
  Next a

End Sub
Public Sub moveshells(s As Long)
    If sdir(s) = 1 Then
      sypos(s) = sypos(s) - ss(s) / 20
    End If
    If sdir(s) = 2 Then
      sypos(s) = sypos(s) - ss(s) / 22
      sxpos(s) = sxpos(s) + ss(s) / 100
    End If
    If sdir(s) = 3 Then
      sypos(s) = sypos(s) - ss(s) / 25
      sxpos(s) = sxpos(s) + ss(s) / 60
    End If
    If sdir(s) = 4 Then
      sypos(s) = sypos(s) - ss(s) / 27
      sxpos(s) = sxpos(s) + ss(s) / 50
    End If
    If sdir(s) = 5 Then
      sypos(s) = sypos(s) - ss(s) / 29
      sxpos(s) = sxpos(s) + ss(s) / 35
    End If
    If sdir(s) = 6 Then
      sypos(s) = sypos(s) - ss(s) / 35
      sxpos(s) = sxpos(s) + ss(s) / 29
    End If
    If sdir(s) = 7 Then
      sypos(s) = sypos(s) - ss(s) / 50
      sxpos(s) = sxpos(s) + ss(s) / 27
    End If
    If sdir(s) = 8 Then
      sypos(s) = sypos(s) - ss(s) / 60
      sxpos(s) = sxpos(s) + ss(s) / 25
    End If
    If sdir(s) = 9 Then
      sypos(s) = sypos(s) - ss(s) / 100
      sxpos(s) = sxpos(s) + ss(s) / 22
    End If
    If sdir(s) = 10 Then
      sxpos(s) = sxpos(s) + ss(s) / 20
    End If
    If sdir(s) = 11 Then
      sypos(s) = sypos(s) + ss(s) / 100
      sxpos(s) = sxpos(s) + ss(s) / 22
    End If
    If sdir(s) = 12 Then
      sypos(s) = sypos(s) + ss(s) / 60
      sxpos(s) = sxpos(s) + ss(s) / 25
    End If
    If sdir(s) = 13 Then
      sypos(s) = sypos(s) + ss(s) / 50
      sxpos(s) = sxpos(s) + ss(s) / 27
    End If
    If sdir(s) = 14 Then
      sypos(s) = sypos(s) + ss(s) / 35
      sxpos(s) = sxpos(s) + ss(s) / 29
    End If
    If sdir(s) = 15 Then
      sypos(s) = sypos(s) + ss(s) / 29
      sxpos(s) = sxpos(s) + ss(s) / 35
    End If
    If sdir(s) = 16 Then
      sypos(s) = sypos(s) + ss(s) / 27
      sxpos(s) = sxpos(s) + ss(s) / 50
    End If
    If sdir(s) = 17 Then
      sypos(s) = sypos(s) + ss(s) / 25
      sxpos(s) = sxpos(s) + ss(s) / 60
    End If
    If sdir(s) = 18 Then
      sypos(s) = sypos(s) + ss(s) / 22
      sxpos(s) = sxpos(s) + ss(s) / 100
    End If
    If sdir(s) = 19 Then
      sypos(s) = sypos(s) + ss(s) / 20
    End If
      If sdir(s) = 20 Then
      sypos(s) = sypos(s) + ss(s) / 22
      sxpos(s) = sxpos(s) - ss(s) / 100
    End If
    If sdir(s) = 21 Then
      sypos(s) = sypos(s) + ss(s) / 25
      sxpos(s) = sxpos(s) - ss(s) / 60
    End If
    If sdir(s) = 22 Then
      sypos(s) = sypos(s) + ss(s) / 27
      sxpos(s) = sxpos(s) - ss(s) / 50
    End If
    If sdir(s) = 23 Then
      sypos(s) = sypos(s) + ss(s) / 29
      sxpos(s) = sxpos(s) - ss(s) / 35
    End If
    If sdir(s) = 24 Then
      sypos(s) = sypos(s) + ss(s) / 35
      sxpos(s) = sxpos(s) - ss(s) / 29
    End If
    If sdir(s) = 25 Then
      sypos(s) = sypos(s) + ss(s) / 50
      sxpos(s) = sxpos(s) - ss(s) / 27
    End If
    If sdir(s) = 26 Then
      sypos(s) = sypos(s) + ss(s) / 60
      sxpos(s) = sxpos(s) - ss(s) / 25
    End If
    If sdir(s) = 27 Then
      sypos(s) = sypos(s) + ss(s) / 100
      sxpos(s) = sxpos(s) - ss(s) / 22
    End If
    If sdir(s) = 28 Then
      sxpos(s) = sxpos(s) - ss(s) / 20
    End If
    If sdir(s) = 29 Then
      sypos(s) = sypos(s) - ss(s) / 100
      sxpos(s) = sxpos(s) - ss(s) / 22
    End If
    If sdir(s) = 30 Then
      sypos(s) = sypos(s) - ss(s) / 60
      sxpos(s) = sxpos(s) - ss(s) / 25
    End If
    If sdir(s) = 31 Then
      sypos(s) = sypos(s) - ss(s) / 50
      sxpos(s) = sxpos(s) - ss(s) / 27
    End If
    If sdir(s) = 32 Then
      sypos(s) = sypos(s) - ss(s) / 35
      sxpos(s) = sxpos(s) - ss(s) / 29
    End If
    If sdir(s) = 33 Then
      sypos(s) = sypos(s) - ss(s) / 29
      sxpos(s) = sxpos(s) - ss(s) / 35
    End If
    If sdir(s) = 34 Then
      sypos(s) = sypos(s) - ss(s) / 27
      sxpos(s) = sxpos(s) - ss(s) / 50
    End If
    If sdir(s) = 35 Then
      sypos(s) = sypos(s) - ss(s) / 25
      sxpos(s) = sxpos(s) - ss(s) / 60
    End If
    If sdir(s) = 36 Then
      sypos(s) = sypos(s) - ss(s) / 22
      sxpos(s) = sxpos(s) - ss(s) / 100
    End If
 
End Sub

Public Function ms(thiscontrol As Control, formwidth As Long)
  thiscontrol.Left = (formwidth - thiscontrol.Width) / 2
End Function

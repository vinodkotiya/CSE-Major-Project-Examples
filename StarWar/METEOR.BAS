Attribute VB_Name = "Module1"
' Declares global variables for Star War project
Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal XSrc As Integer, ByVal YSrc As Integer, ByVal dwRop As Long) As Integer
Global Const SRCCOPY = &HCC0020
Global Const SRCPAINT = &HEE0086
Global Const SRCAND = &H8800C6
Global STAR(30) As Integer
Global XPLANE As Integer
Global YPLANE As Integer
Global XMETEOR1 As Integer
Global YMETEOR1 As Integer
Global XMETEOR2 As Integer
Global YMETEOR2 As Integer
Global XMETEOR3 As Integer
Global YMETEOR3 As Integer
Global XMETEOR4 As Integer
Global YMETEOR4 As Integer
Global XMETEOR5 As Integer
Global YMETEOR5 As Integer
Global XBULLET As Integer
Global YBULLET As Integer
Global GAMECTR As Double
Global BACKGROUNDCTR As Integer
Global SCORECTR As Integer
Global LIFECTR As Integer
Global HITSW As Integer
Global COLLIDESW As Integer
Global FIRESW As Integer
Global MOUSESW As Integer



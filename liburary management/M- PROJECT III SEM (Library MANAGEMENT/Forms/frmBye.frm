VERSION 5.00
Begin VB.Form frmBye 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3240
      Top             =   360
   End
   Begin VB.Image Image1 
      Height          =   2250
      Left            =   0
      Picture         =   "frmBye.frx":0000
      Top             =   0
      Width           =   4500
   End
End
Attribute VB_Name = "frmBye"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim X As Integer

Private Sub Form_Initialize()
    X = 0
End Sub

Private Sub Form_Load()
    frmBye.Height = Image1.Height
    frmBye.Width = Image1.Width
End Sub

Private Sub Image1_Click()
    Unload Me
    End
End Sub

Private Sub Timer1_Timer()
    If X < 100 Then
        X = X + 20
    Else
        Unload Me
        End
    End If
End Sub

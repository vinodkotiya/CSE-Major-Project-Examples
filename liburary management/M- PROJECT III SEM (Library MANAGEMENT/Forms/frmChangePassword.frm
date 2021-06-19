VERSION 5.00
Begin VB.Form frmChangePassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "frmChangePassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Tag             =   "Login"
   Begin VB.Frame Frame1 
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   15
      TabIndex        =   5
      Top             =   -30
      Width           =   4530
      Begin VB.ComboBox Usr 
         Height          =   315
         ItemData        =   "frmChangePassword.frx":0442
         Left            =   1800
         List            =   "frmChangePassword.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox txt 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   2325
      End
      Begin VB.TextBox txt 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   960
         Width           =   2325
      End
      Begin VB.TextBox txt 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   600
         Width           =   2325
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   360
         Left            =   600
         TabIndex        =   3
         Tag             =   "OK"
         Top             =   1695
         Width           =   1140
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   360
         Left            =   2535
         TabIndex        =   4
         Tag             =   "Cancel"
         Top             =   1695
         Width           =   1140
      End
      Begin VB.Label lblLabels 
         Caption         =   "User Name"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Tag             =   "&Password:"
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label lblLabels 
         Caption         =   "Retype Password:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Tag             =   "&Password:"
         Top             =   1320
         Width           =   1320
      End
      Begin VB.Label lblLabels 
         Caption         =   "Old Password:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Tag             =   "&User Name:"
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         Caption         =   "New Password:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Tag             =   "&Password:"
         Top             =   960
         Width           =   1200
      End
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    On Error Resume Next
    Unload Me
End Sub
Private Sub cmdOK_Click()
    'On Error Resume Next
    Screen.MousePointer = 13
    '=======================================================
    ' checking whetherUser name and password are not empty
    '=======================================================
    Dim i
    For i = 0 To 2
        If txt(i) = "" Then
            Beep
            MsgBox "Please enter your " & lblLabels(i) & " .", vbExclamation, "Error"
            Screen.MousePointer = 0
            txt(i).SetFocus
            Exit Sub
        End If
    Next
        
    If txt(1).text <> txt(2).text Then
        Beep
        MsgBox "Password doesn't match.", vbExclamation, "Error"
        Screen.MousePointer = 0
        txt(1).text = ""
        txt(2).text = ""
        txt(1).SetFocus
        Exit Sub
    Else
        With ObjCon
            .Open FileDSN
                 query = "update Login set Password='" & LCase(txt(1).text) & "' where UserName='" & LCase(Usr.text) & "'"
                 .Execute (query)
                 Beep
                 MsgBox "Updated", vbInformation, "Info"
            .Close
        End With
        txt(0) = ""
        txt(1) = ""
        txt(2) = ""
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    txt(0).SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    tot (-1)
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Screen.MousePointer = 0
    Usr.text = Usr.List(0)
    tot (1)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        Call cmdOK_Click
    End If
End Sub

VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "~(*_*)~ "
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "MathTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2190
      Top             =   3735
   End
   Begin VB.CheckBox chkRandomSign 
      Caption         =   "Auto Change Signs"
      Height          =   285
      Left            =   135
      TabIndex        =   24
      Top             =   3795
      Width           =   1725
   End
   Begin VB.Frame Frame4 
      Height          =   3525
      Left            =   6360
      TabIndex        =   23
      Top             =   120
      Width           =   2610
      Begin VB.CommandButton cmdNumber 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Index           =   10
         Left            =   960
         TabIndex        =   15
         Top             =   2775
         Width           =   1560
      End
      Begin VB.CommandButton cmdNumber 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   9
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Width           =   720
      End
      Begin VB.CommandButton cmdNumber 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Index           =   8
         Left            =   1800
         TabIndex        =   13
         Top             =   1920
         Width           =   700
      End
      Begin VB.CommandButton cmdNumber 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Index           =   7
         Left            =   960
         TabIndex        =   12
         Top             =   1920
         Width           =   700
      End
      Begin VB.CommandButton cmdNumber 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   700
      End
      Begin VB.CommandButton cmdNumber 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Index           =   5
         Left            =   1800
         TabIndex        =   10
         Top             =   1095
         Width           =   700
      End
      Begin VB.CommandButton cmdNumber 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Index           =   4
         Left            =   960
         TabIndex        =   9
         Top             =   1080
         Width           =   700
      End
      Begin VB.CommandButton cmdNumber 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   700
      End
      Begin VB.CommandButton cmdNumber 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Index           =   2
         Left            =   1800
         TabIndex        =   7
         Top             =   240
         Width           =   700
      End
      Begin VB.CommandButton cmdNumber 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Index           =   1
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   700
      End
      Begin VB.CommandButton cmdNumber 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   700
      End
   End
   Begin VB.CommandButton cmdChangeSign 
      Caption         =   "Change Sign"
      Height          =   525
      Left            =   120
      TabIndex        =   1
      Top             =   3135
      Width           =   1400
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   105
      TabIndex        =   21
      Top             =   135
      Width           =   6075
      Begin VB.Label lblMessage 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Incorrect"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1305
         Left            =   150
         TabIndex        =   22
         Top             =   255
         Width           =   5745
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      Height          =   510
      Left            =   3240
      TabIndex        =   2
      Top             =   3135
      Width           =   1400
   End
   Begin VB.CommandButton cmdAnswer 
      Caption         =   "Answer"
      Height          =   510
      Left            =   1695
      TabIndex        =   3
      Top             =   3135
      Width           =   1400
   End
   Begin VB.CommandButton cmdCheckAnswer 
      Caption         =   "Check"
      Height          =   525
      Left            =   4800
      TabIndex        =   4
      Top             =   3135
      Width           =   1400
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   120
      TabIndex        =   16
      Top             =   1935
      Width           =   6075
      Begin VB.TextBox txtAnswer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   3960
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   270
         Width           =   1935
      End
      Begin VB.Image imgDiv 
         Height          =   450
         Left            =   1635
         Picture         =   "MathTest.frx":030A
         Top             =   315
         Width           =   420
      End
      Begin VB.Label lblEquals 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   3600
         TabIndex        =   20
         Top             =   270
         Width           =   465
      End
      Begin VB.Label lblSign 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   1620
         TabIndex        =   19
         Top             =   270
         Width           =   495
      End
      Begin VB.Label lblNumberTwo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   2040
         TabIndex        =   18
         Top             =   270
         Width           =   1590
      End
      Begin VB.Label lblNumberOne 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   180
         TabIndex        =   17
         Top             =   270
         Width           =   1440
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   Math Test
'------------------------------------------------------------------------------------------

Option Explicit
Dim vntTemp As Variant  'Multipurpose/Temporary Variable
Dim intCounter As Integer
Dim blnLR As Boolean
Private Sub Form_Load()
    txtAnswer.Text = ""
    txtAnswer.ForeColor = RGB(71, 86, 220)
    cmdNext_Click
    imgDiv.Visible = False
    intCounter = 0
    blnLR = True
End Sub

'Addition
Private Function Sum(vntA As Variant, vntB As Variant) As Integer
    Sum = CInt(Val(vntA)) + CInt(Val(vntB))
End Function

'Division
Private Function Remainder(vntA As Variant, vntB As Variant) As Double
    If CInt(Val(vntA)) <> 0 Then
        Remainder = CInt(Val(vntA)) / CInt(Val(vntB))
    End If
End Function

'Subtraction
Private Function Difference(vntA As Variant, vntB As Variant) As Integer
    Difference = CInt(Val(vntA)) - CInt(Val(vntB))
End Function

'Multiplication
Private Function Product(vntA As Variant, vntB As Variant) As Integer
    Product = CInt(Val(vntA)) * CInt(Val(vntB))
End Function

'Check Box/Auto Change Signs
Private Sub chkRandomSign_Click()
    If chkRandomSign.Value = 1 Then
        cmdChangeSign.Enabled = False
    Else
        cmdChangeSign.Enabled = True
    End If
End Sub

'Change Sign
Private Sub cmdChangeSign_Click()
    imgDiv.Visible = False
    Select Case (lblSign.Caption)
        Case "+"
            lblSign.Caption = "-"
        Case "-"
            lblSign.Caption = "x"
        Case "x"
            lblSign.Caption = "/"
            imgDiv.Visible = True
            If chkRandomSign.Value <> 1 Then
                Call cmdNext_Click
            End If
        Case "/"
            lblSign.Caption = "+"
    End Select
End Sub

'Number Buttons
Private Sub cmdNumber_Click(Index As Integer)
    If Index <> 10 Then
        txtAnswer.Text = txtAnswer.Text + Trim(Str(cmdNumber(Index).Caption))
    Else
        txtAnswer.Text = ""
        DisplayMessage "Math Test", 71, 86, 220
    End If
End Sub

'Generate Next Problem
Private Sub cmdNext_Click()
    Dim vntTemp As Integer
    EnableButtons
    If chkRandomSign.Value = 1 Then Call cmdChangeSign_Click
    DisplayMessage "Math Test", 71, 86, 220
    txtAnswer.Text = ""
StartLoop:
    For vntTemp = 0 To Val(Right$(Time$, 2))
        lblNumberOne.Caption = Right$(Rnd(300), 2)
        lblNumberTwo.Caption = Right$(Rnd(400), 2)
    Next
    If Val(lblNumberOne.Caption) < Val(lblNumberTwo.Caption) Then
        vntTemp = Val(lblNumberOne.Caption)
        lblNumberOne.Caption = lblNumberTwo.Caption
        lblNumberTwo.Caption = vntTemp
    End If
    If lblSign = "/" And Val(lblNumberOne.Caption) _
        Mod Val(lblNumberTwo.Caption) <> 0 Then GoTo StartLoop
    If Val(lblNumberTwo.Caption) < 3 Then GoTo StartLoop
    If lblNumberOne.Caption = lblNumberTwo.Caption Then GoTo StartLoop
End Sub

'Display Answer
Private Sub cmdAnswer_Click()
    DisableButtons
    DisplayMessage "Math Test", 7, 86, 220
    Select Case (lblSign.Caption)
        Case "+"
            txtAnswer = Sum(lblNumberOne.Caption, lblNumberTwo.Caption)
        Case "-"
            txtAnswer = Difference(lblNumberOne.Caption, lblNumberTwo.Caption)
        Case "x"
            txtAnswer = Product(lblNumberOne.Caption, lblNumberTwo.Caption)
        Case "/"
            txtAnswer = Remainder(lblNumberOne.Caption, lblNumberTwo.Caption)
    End Select
End Sub

'Check Answer
Private Sub cmdCheckAnswer_Click()
    DisplayMessage "Incorrect", 255, 0, 0
    Select Case (lblSign.Caption)
        Case "+"
            If Val(txtAnswer.Text) <> Sum(lblNumberOne.Caption, lblNumberTwo.Caption) Then Exit Sub
        Case "-"
            If Val(txtAnswer.Text) <> Difference(lblNumberOne.Caption, lblNumberTwo.Caption) Then Exit Sub
        Case "x"
            If Val(txtAnswer.Text) <> Product(lblNumberOne.Caption, lblNumberTwo.Caption) Then Exit Sub
        Case "/"
            If Val(Left$(txtAnswer.Text, 5)) <> Remainder(lblNumberOne.Caption, lblNumberTwo.Caption) Then Exit Sub
    End Select
    DisplayMessage "Correct", 0, 255, 0
End Sub

'Reset Message and Colors
Private Sub txtAnswer_GotFocus()
    txtAnswer.ForeColor = RGB(7, 86, 220)
    DisplayMessage "Math Test", 7, 86, 220
End Sub

'Check Enter
Private Sub txtAnswer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdCheckAnswer_Click
    End If
End Sub

'Display Messages
Private Sub DisplayMessage(strMessage As String, colorRed As Integer, colorGreen As Integer, colorBlue As Integer)
    lblMessage.ForeColor = RGB(colorRed, colorGreen, colorBlue)
    lblMessage.Caption = strMessage
End Sub

'Disable Command Buttons
Private Function DisableButtons()
    cmdCheckAnswer.Enabled = False
    For vntTemp = 0 To 10
        cmdNumber(vntTemp).Enabled = False
    Next vntTemp
    txtAnswer.Enabled = False
End Function
'Enable Command Buttons
Private Function EnableButtons()
    cmdCheckAnswer.Enabled = True
    For vntTemp = 0 To 10
        cmdNumber(vntTemp).Enabled = True
    Next vntTemp
    txtAnswer.Enabled = True
End Function

Private Sub Timer1_Timer()
    If blnLR Then
        Form1.Caption = " " + Form1.Caption
    Else
        Form1.Caption = Right$(Form1.Caption, Len(Form1.Caption) - 1)
    End If
    intCounter = intCounter + 1
    If intCounter > 15 Then
        intCounter = 0
        If blnLR = True Then
            blnLR = False
        Else
            blnLR = True
        End If
    End If
End Sub

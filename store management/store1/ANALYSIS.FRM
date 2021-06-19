VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form12 
   Caption         =   "Form12"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form12"
   MDIChild        =   -1  'True
   Picture         =   "Analysis.frx":0000
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mhAny 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   7435
      _Version        =   393216
      BackColor       =   12632256
      BackColorFixed  =   8421504
      BackColorBkg    =   14737632
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   5
      X1              =   0
      X2              =   11880
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   735
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub GridSet1()
mhAny.ColWidth(0) = 200
mhAny.ColWidth(1) = 900
mhAny.ColWidth(2) = 900
mhAny.ColWidth(3) = 1500
mhAny.ColWidth(4) = 2400
mhAny.ColWidth(5) = 800
mhAny.ColWidth(6) = 800
mhAny.ColWidth(7) = 1900
mhAny.ColWidth(8) = 1600
mhAny.ColWidth(9) = 1000
mhAny.TextMatrix(0, 1) = "Date"
mhAny.TextMatrix(0, 2) = "IndentNo"
mhAny.TextMatrix(0, 3) = "ItemName"
mhAny.TextMatrix(0, 4) = "Description"
mhAny.TextMatrix(0, 5) = "Quantity"
mhAny.TextMatrix(0, 6) = "EmpCode"
mhAny.TextMatrix(0, 7) = "Employee Name"
mhAny.TextMatrix(0, 8) = "Department"
mhAny.TextMatrix(0, 9) = "Consider"
End Sub

Public Sub GridSet2()
mhAny.ColWidth(0) = 300
mhAny.ColWidth(1) = 1200
mhAny.ColWidth(2) = 1200
mhAny.ColWidth(3) = 1500
mhAny.ColWidth(4) = 2000
mhAny.ColWidth(5) = 700
mhAny.ColWidth(6) = 700
mhAny.ColWidth(7) = 700
mhAny.ColWidth(8) = 700
mhAny.ColWidth(9) = 2000
mhAny.ColWidth(10) = 700
mhAny.ColWidth(11) = 1600
mhAny.ColWidth(12) = 1400
mhAny.TextMatrix(0, 1) = "IssueDate"
mhAny.TextMatrix(0, 2) = "ReqItmNo"
mhAny.TextMatrix(0, 3) = "ItemName"
mhAny.TextMatrix(0, 4) = "Breif"
mhAny.TextMatrix(0, 5) = "StkSno"
mhAny.TextMatrix(0, 6) = "IssueQty"
mhAny.TextMatrix(0, 7) = "Issued"
mhAny.TextMatrix(0, 8) = "Transfered"
mhAny.TextMatrix(0, 9) = "Remark"
mhAny.TextMatrix(0, 10) = "Ecode"
mhAny.TextMatrix(0, 11) = "Ename"
mhAny.TextMatrix(0, 12) = "Department"
End Sub

Public Sub GridSet3()
mhAny.ColWidth(0) = 300
mhAny.ColWidth(1) = 700
mhAny.ColWidth(2) = 900
mhAny.ColWidth(3) = 1700
mhAny.ColWidth(4) = 1700
mhAny.ColWidth(5) = 1700
mhAny.ColWidth(6) = 1200
mhAny.ColWidth(7) = 1000
mhAny.ColWidth(8) = 1000
mhAny.ColWidth(9) = 1000
mhAny.ColWidth(10) = 1000
mhAny.ColWidth(11) = 1000
mhAny.ColWidth(12) = 1450
mhAny.ColWidth(13) = 1450
mhAny.ColWidth(14) = 1000
End Sub

Private Sub Command1_Click()
Unload Me
End Sub


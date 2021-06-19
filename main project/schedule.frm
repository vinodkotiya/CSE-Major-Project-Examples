VERSION 5.00
Begin VB.Form frmsch 
   AutoRedraw      =   -1  'True
   Caption         =   "schedule"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13380
   LinkMode        =   1  'Source
   LinkTopic       =   "rail"
   MaxButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   13380
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text8 
      DataField       =   "Trains"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3480
      Width           =   4215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Home"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8400
      TabIndex        =   5
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Train Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      TabIndex        =   4
      Top             =   4680
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      DataField       =   "Train_no"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      MaxLength       =   5
      TabIndex        =   2
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Enter Train No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SCHEDULE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   3240
      TabIndex        =   0
      Top             =   1080
      Width           =   8535
   End
   Begin VB.Image Image2 
      Height          =   11415
      Left            =   -120
      Picture         =   "schedule.frx":0000
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   15480
   End
End
Attribute VB_Name = "frmsch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub frmsch_Load()
Text1.SetFocus
Text1.Text = " "
Text2.Text = " "
End Sub
Private Sub Command1_Click()
Dim rrsearch As New ADODB.Recordset
If Text1.Text = " " Then
MsgBox "blank text", vbDefaultButton2
Me.Show
Text1.Text = " "
frmsch.Text1.SetFocus
Else
rrsearch.Open "select * from down where train_no = " & Text1.Text & " ", cn, adOpenDynamic
If rrsearch.EOF = True Then
MsgBox "This Train No. is not present", vbApplicationModal
Text1.Text = " "
Text8.Text = " "
frmsch.Text1.SetFocus
Else
Unload Me
frmschoutput.Show
frmschoutput.Label9.Caption = rrsearch(1)
frmschoutput.Label11.Caption = rrsearch(0)
frmschoutput.Label13.Caption = rrsearch(3)
frmschoutput.Label15.Caption = rrsearch(4)
frmschoutput.Label19.Caption = rrsearch(5)
frmschoutput.Label20.Caption = rrsearch(6)
frmschoutput.Label21.Caption = rrsearch(2)
rrsearch.Close
End If
End If
End Sub

Private Sub Command2_Click()
Dim rrsearch1 As New ADODB.Recordset
If Text1.Text = " " Then
MsgBox "blank text", vbDefaultButton2
Me.Show
Text1.Text = " "
frmsch.Text1.SetFocus
Else
rrsearch1.Open "select * from down where train_no = " & Text1.Text & " ", cn, adOpenDynamic
If rrsearch1.EOF Then
MsgBox "No train name is find", vbApplicationModal
Me.Show
Text1.Text = " "
Text8.Text = " "
frmsch.Text1.SetFocus
Else
frmsch.Text8.Text = rrsearch1(1)
rrsearch1.Close
End If
End If
End Sub
Private Sub Command3_Click()
cn.Close
cn1.Close
Me.Hide
MDIForm1.Show
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command2.SetFocus
End If
End Sub


VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Main Form (Railway Enquiry)"
   ClientHeight    =   8700
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11670
   LinkTopic       =   "MDIForm1"
   MouseIcon       =   "MDIForm5.frx":0000
   Moveable        =   0   'False
   OLEDropMode     =   1  'Manual
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   10815
      Left            =   0
      ScaleHeight     =   10755
      ScaleWidth      =   11610
      TabIndex        =   0
      Top             =   0
      Width           =   11670
      Begin VB.PictureBox Adodc1 
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   9720
         ScaleHeight     =   315
         ScaleWidth      =   1155
         TabIndex        =   4
         Top             =   10440
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   14640
         Top             =   1560
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12600
         TabIndex        =   3
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12600
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12600
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   10935
         Left            =   -120
         Picture         =   "MDIForm5.frx":0442
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   15600
      End
   End
   Begin VB.Menu opt 
      Caption         =   "&Enquiry"
      Index           =   1
      NegotiatePosition=   1  'Left
      Begin VB.Menu far 
         Caption         =   "Fares"
         Index           =   2
         Begin VB.Menu sha 
            Caption         =   "Shatabdi Express Pairs of Station Wise Fares"
            Index           =   17
         End
         Begin VB.Menu raj 
            Caption         =   "Rajdhani Express Pairs of Station Wise Fares"
            Index           =   18
         End
         Begin VB.Menu jan 
            Caption         =   "Jan Shatabdi Express Pairs of Station Wise Fares"
            Index           =   19
         End
         Begin VB.Menu shakm 
            Caption         =   "Shatabdi Express K/m."
            Index           =   20
         End
         Begin VB.Menu rajkm 
            Caption         =   "Rajdhani Express K/m."
            Index           =   21
         End
         Begin VB.Menu mail 
            Caption         =   "Mail/Express K/m."
            Index           =   22
         End
         Begin VB.Menu jankm 
            Caption         =   "Jan Shatabdi Exp. K/m."
            Index           =   23
         End
      End
      Begin VB.Menu sch 
         Caption         =   "Schedule"
      End
      Begin VB.Menu res 
         Caption         =   "Reservation "
         Index           =   3
         Shortcut        =   ^R
      End
      Begin VB.Menu ta 
         Caption         =   "Trains Available"
         Index           =   17
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu rules 
      Caption         =   "&Rules"
      Index           =   7
      Begin VB.Menu rr 
         Caption         =   "Reservation Rules"
         Index           =   8
      End
      Begin VB.Menu rer 
         Caption         =   "Refunds Rules"
         Index           =   9
      End
      Begin VB.Menu bjr 
         Caption         =   "Break Journy Rules"
         Index           =   10
      End
      Begin VB.Menu lug 
         Caption         =   "Luggage"
         Index           =   11
      End
      Begin VB.Menu cnr 
         Caption         =   "Change in Name Rule"
         Index           =   14
      End
      Begin VB.Menu cjr 
         Caption         =   "Circular Journey Rules"
         Index           =   12
      End
      Begin VB.Menu scr 
         Caption         =   "Sr. Citizen Rules"
         Index           =   13
      End
   End
   Begin VB.Menu ext 
      Caption         =   "E&xit"
      Index           =   5
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ext_Click(Index As Integer)
End
End Sub

Private Sub MDIForm_Load()
Dim asd As String
Dim asd1 As String
asd = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=rail.mdb;Persist Security Info=False"
cn.Open (asd)
asd1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=pnr.mdb;Persist Security Info=False"
cn1.Open (asd1)
End Sub
Private Sub res_Click(Index As Integer)
Unload Me
frmpnr.Show
frmpnr.Text1.SetFocus
End Sub
Private Sub rr_Click(Index As Integer)
FileName = InputBox("D:\rules_reservation_files\resrules.htm", "open file")
Open "D:\rules_reservation_files\resrules.htm" For Input As #1
'Input #1, maxmumber
'ReDim dataset(maxnumber - 1)
For element = 0 To maxnumber - 1
'                            Print #1, dataset(element)
Next element
Close #1
'input "D:\rules_reservation_files\resrules.htm" For output  As #10
End Sub
Private Sub sch_Click()
Unload Me
frmsch.Show
frmsch.Text1.Text = " "
frmsch.Text1.SetFocus
frmsch.Text8.Text = ""
End Sub
Private Sub sha_Click(Index As Integer)
Unload Me
frmfare.Show
End Sub
Private Sub ta_Click(Index As Integer)
Unload Me
frminput.Show
frminput.Text1.SetFocus
End Sub
Private Sub Timer1_Timer()
Dim a As Date
Label1.Caption = Date
Label2.Caption = Format(a, "dddd")
Label3.Caption = Time
End Sub
        

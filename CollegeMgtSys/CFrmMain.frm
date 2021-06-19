VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm CFrmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "COLLEGE MANAGEMENT SYSTEM"
   ClientHeight    =   8310
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11880
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   480
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   847
      ButtonWidth     =   714
      ButtonHeight    =   688
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "one"
            Object.ToolTipText     =   "Enquiry"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "two"
            Object.ToolTipText     =   "New Adm"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "three"
            Object.ToolTipText     =   "New Course"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "space"
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "five"
            Object.ToolTipText     =   "New Faculty"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "six"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "seven"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "space2"
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nine"
            Object.ToolTipText     =   "Money Recieved"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ten"
            Object.ToolTipText     =   "Money Send"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "space1"
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "eleven"
            Object.ToolTipText     =   "Letter received"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "thirteen"
            Object.ToolTipText     =   "Letter Despatch"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "fourteen"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "fifteen"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFrmMain.frx":0000
            Key             =   "one"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFrmMain.frx":0452
            Key             =   "two"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFrmMain.frx":08A4
            Key             =   "three"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFrmMain.frx":0CF6
            Key             =   "four"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFrmMain.frx":1148
            Key             =   "five"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFrmMain.frx":159A
            Key             =   "six"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFrmMain.frx":19EC
            Key             =   "seven"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFrmMain.frx":1E3E
            Key             =   "eight"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFrmMain.frx":2158
            Key             =   "nine"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CFrmMain.frx":25AA
            Key             =   "ten"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7935
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   15319
            Picture         =   "CFrmMain.frx":29FC
            Text            =   "This Softwere is Developed By Vidya Bhushan Singh"
            TextSave        =   "This Softwere is Developed By Vidya Bhushan Singh"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Picture         =   "CFrmMain.frx":2E4E
            TextSave        =   "2/3/03"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Picture         =   "CFrmMain.frx":3172
            TextSave        =   "5:19 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MnuStudDetails 
      Caption         =   "&Student Details"
      Index           =   10
      Begin VB.Menu MnuStdDetails 
         Caption         =   "&Enquiry"
         Index           =   0
      End
      Begin VB.Menu MnuStdDetails 
         Caption         =   "New &Admission"
         Index           =   1
      End
      Begin VB.Menu MnuStdDetails 
         Caption         =   "New &Course"
         Index           =   2
      End
      Begin VB.Menu MnuStdDetails 
         Caption         =   "Hostel"
         Index           =   3
      End
      Begin VB.Menu MnuStdDetails 
         Caption         =   "Tc Status"
         Index           =   4
      End
      Begin VB.Menu MnuStdDetails 
         Caption         =   "Fee Status"
         Index           =   5
      End
      Begin VB.Menu MnuStdDetails 
         Caption         =   "Search"
         Index           =   6
      End
      Begin VB.Menu MnuStdDetails 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu MnuStdDetails 
         Caption         =   "E&xit"
         Index           =   8
      End
   End
   Begin VB.Menu MnuFatDetails 
      Caption         =   "&Faculty Details"
      Begin VB.Menu MnuFaculty 
         Caption         =   "New Faculty"
         Index           =   0
      End
      Begin VB.Menu MnuFaculty 
         Caption         =   "Salary Status"
         Index           =   1
      End
      Begin VB.Menu MnuFaculty 
         Caption         =   "Re Joining"
         Index           =   2
      End
      Begin VB.Menu MnuFaculty 
         Caption         =   "Edit User"
         Index           =   3
      End
   End
   Begin VB.Menu MnuTRans 
      Caption         =   "&Transactions"
      Begin VB.Menu MnuReceived 
         Caption         =   "Received"
      End
      Begin VB.Menu MnuDispatch 
         Caption         =   "Dispatched"
      End
      Begin VB.Menu MnuIncome 
         Caption         =   "Income"
      End
      Begin VB.Menu asdas 
         Caption         =   "asda"
      End
   End
   Begin VB.Menu MnuGoods 
      Caption         =   "Goods Status"
   End
   Begin VB.Menu MnuChartDis 
      Caption         =   "&Chart Display"
   End
   Begin VB.Menu MnuPrtReport 
      Caption         =   "&Print Report"
   End
End
Attribute VB_Name = "CFrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    
        
    ' Connection With Database
     ' SeqGen.ConnConnect
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    'Disconnection With Database
    SeqGen.DisConnect
End Sub



'calling of faculty forms from main menu

Private Sub MnuFaculty_Click(Index As Integer)
    Select Case Index
        Case 0: 'load new faculty
            Load FrmFatDetails
        Case 1: 'load salary form
        Case 2: 'load rejoin form
        Case 3: 'Edit user form
            Load FrmEditUser
    End Select
End Sub

'Calling of student forms from main form

Private Sub MnuStdDetails_Click(Index As Integer)
    Select Case Index 'select appropiate menu
        Case 0: 'Load enquiry form
            Load FrmEnq
        Case 1: 'Load New Adm Form
            Load FrmSDetails
        Case 2: 'load New Course form
            Load FrmCourse
        Case 3: 'load hostel form
        Case 4: 'load Tc Form
        Case 5: 'load fees Form
        Case 6: 'load search form
        Case 7: 'nothing
        Case 8: 'exit from project
            Unload CFrmMain
           
    End Select
End Sub

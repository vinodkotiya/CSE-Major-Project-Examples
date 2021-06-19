VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "LIBRARY MANAGEMENT"
   ClientHeight    =   6600
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7485
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   11
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Insert"
            Object.ToolTipText     =   "New Books"
            Object.Tag             =   ""
            ImageKey        =   "Insert"
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Issue"
            Object.ToolTipText     =   "Issue Books Students"
            Object.Tag             =   ""
            ImageKey        =   "Issue"
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Return"
            Object.ToolTipText     =   "Return Books"
            Object.Tag             =   ""
            ImageKey        =   "Return"
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "SearchBook"
            Object.ToolTipText     =   "Search Books"
            Object.Tag             =   ""
            ImageKey        =   "SBook"
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "RBook"
            Object.ToolTipText     =   "Report Books"
            Object.Tag             =   ""
            ImageKey        =   "RBooks"
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "RIssue"
            Object.ToolTipText     =   "Report Issue"
            Object.Tag             =   ""
            ImageKey        =   "RIssue"
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Status"
            Object.ToolTipText     =   "Book Status"
            Object.Tag             =   ""
            ImageKey        =   "Status"
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin ComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   6330
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   7646
            Text            =   "Press <TAB> To move between Fields"
            TextSave        =   "Press <TAB> To move between Fields"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            TextSave        =   "4/7/03"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            TextSave        =   "12:34 THAKUR"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2040
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0442
            Key             =   "Issue"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":075C
            Key             =   "Return"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0D90
            Key             =   "Insert"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":10AA
            Key             =   "SBook"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":13C4
            Key             =   "RBooks"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":16DE
            Key             =   "RIssue"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":19F8
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1D12
            Key             =   "Status"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileChangePassword 
         Caption         =   "C&hange Password"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileCloseAll 
         Caption         =   "C&lose All"
         Enabled         =   0   'False
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Enabled         =   0   'False
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnuInsertBooks 
         Caption         =   "&Books"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuInsertDisMis 
         Caption         =   "&Discarded and Missing"
      End
      Begin VB.Menu mnuInsertBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertStudents 
         Caption         =   "&Students"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuInsertBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertTeachers 
         Caption         =   "&Teachers"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditBooks 
         Caption         =   "&Books Details"
      End
      Begin VB.Menu mnuEditDisMis 
         Caption         =   "&Discarded and Missing"
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStudents 
         Caption         =   "&Students Details"
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditTeachersDetails 
         Caption         =   "&Teachers Details"
      End
   End
   Begin VB.Menu mnuIssue 
      Caption         =   "&Issue"
      Begin VB.Menu mnuIssueStudents 
         Caption         =   "&Students"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuIssueBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIssueTeachers 
         Caption         =   "&Teachers"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu mnuReturn 
      Caption         =   "&Return"
      Begin VB.Menu mnuReturnBooks 
         Caption         =   "&Return Books"
         Shortcut        =   {F2}
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchBooks 
         Caption         =   "&Book Search By"
         Begin VB.Menu mnuSearchBooksAccessionNo 
            Caption         =   "&AccessionNo"
         End
         Begin VB.Menu mnuSearchBookBar0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSearchBooksName 
            Caption         =   "&Name"
         End
         Begin VB.Menu mnuSearchBookBar1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSearchBookAuthors 
            Caption         =   "A&uthor"
         End
         Begin VB.Menu mnuSearchBookBar2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSearchPublisher 
            Caption         =   "&Publisher"
         End
      End
      Begin VB.Menu mnuSearchBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBookStatus 
         Caption         =   "B&ook Status"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSearchBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchStudents 
         Caption         =   "&Students By"
         Begin VB.Menu mnuSearchStudentsAlphabetical 
            Caption         =   "&Alphabetical Search"
         End
         Begin VB.Menu mnuSearchStudentsBar0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSearchCourse 
            Caption         =   "&Course"
         End
         Begin VB.Menu mnuSearchStudentsBar1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSearchStudentsAdvanced 
            Caption         =   "&LCardNo"
         End
      End
      Begin VB.Menu mnuSearchBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSearchTeachers 
         Caption         =   "&Teachers By"
         Begin VB.Menu mnuSearchTeachersAlphabetical 
            Caption         =   "&Alphabetical Search"
         End
         Begin VB.Menu mnuSearchTeachersBar0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSearchTeachersLCardNo 
            Caption         =   "&LCardNo"
         End
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub MDIForm_Load()
 
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
    If UCase(LoginPass) <> UCase("Admin") Then
        frmMain.mnuEditStudents.Enabled = False
        frmMain.mnuEditTeachersDetails.Enabled = False
        frmMain.mnuFileChangePassword.Enabled = True
        
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    
End Sub

Private Sub mnuBookStatus_Click()
    Call wait(1)
        Load frmStatusBooks
        Me.Arrange 0
        frmStatusBooks.Show
    Call wait(0)
End Sub


Private Sub mnuEditBooks_Click()
    Call wait(1)
        Load frmModifyBooks
        Me.Arrange 0
        frmModifyBooks.Show
    Call wait(0)
End Sub

Private Sub mnuEditDisMis_Click()
    Call wait(1)
        Load frmEditDisMis
        Me.Arrange 0
        frmEditDisMis.Show
    Call wait(0)
End Sub

Private Sub mnuEditStudents_Click()
    Call wait(1)
    Load frmModifyStudents
    Me.Arrange 0
    frmModifyStudents.Show
    Call wait(0)
End Sub

Private Sub mnuEditTeachersDetails_Click()
    Call wait(1)
    Load frmModifyTeachers
    Me.Arrange 0
    frmModifyTeachers.Show
    Call wait(0)
End Sub

Private Sub mnuFileChangePassword_Click()
    Call wait(1)
        Load frmChangePassword
        Me.Arrange 0
        frmChangePassword.Show
    Call wait(0)
End Sub

Private Sub mnuFileCloseAll_Click()
    While totalforms > 0
        Unload ActiveForm
    Wend
End Sub

'Private Sub mnuHelpAboutUs_Click()
 '   Load frmSplash
  '  frmSplash.ProgressBar1.Visible = True
   ' frmSplash.Show
'End Sub

Private Sub mnuInsertBooks_Click()
    Call wait(1)
        Load frmInsertBooks
        Me.Arrange 0
        frmInsertBooks.Show
    Call wait(0)
End Sub

Private Sub mnuInsertDisMis_Click()
    Call wait(1)
        Load frmInsertDiscarded
        Me.Arrange 0
        frmInsertDiscarded.Show
    Call wait(0)
End Sub

Private Sub mnuInsertStudents_Click()
    Call wait(1)
        Load frmInsertStudents
        Me.Arrange 0
        frmInsertStudents.Show
    Call wait(0)
End Sub

Private Sub mnuInsertTeachers_Click()
    Call wait(1)
        Load frmInsertTeachers
        Me.Arrange 0
        frmInsertTeachers.Show
    Call wait(0)
End Sub

Private Sub mnuIssueStudents_Click()
    Call wait(1)
        Load frmIssueBooks
        Me.Arrange 0
        frmIssueBooks.Show
    Call wait(0)
End Sub

Private Sub mnuIssueTeachers_Click()
    Call wait(1)
        Load frmIssueBooksTeachers
        Me.Arrange 0
        frmIssueBooksTeachers.Show
    Call wait(0)
End Sub

Private Sub mnuReportTeachers_Click()
    Load frmTeacherDialog
    frmTeacherDialog.Show vbModal
End Sub

Private Sub mnuReturnBooks_Click()
    Call wait(1)
        Load frmReturnBooksStudents
        Me.Arrange 0
        frmReturnBooksStudents.Show
    Call wait(0)
End Sub

Private Sub mnuSearchBookAuthors_Click()
    Call wait(1)
        Load frmSearchBooksAuthors
        Me.Arrange 0
        frmSearchBooksAuthors.Show
    Call wait(0)
End Sub

Private Sub mnuSearchBooksAccessionNo_Click()
    Call wait(1)
        Load frmSearchBooks
        Me.Arrange 0
        frmSearchBooks.Show
    Call wait(0)
End Sub

Private Sub mnuSearchBooksName_Click()
    Call wait(1)
        Load frmSearchBooksAlphabets
        Me.Arrange 0
        frmSearchBooksAlphabets.Show
    Call wait(0)
End Sub

Private Sub mnuSearchCourse_Click()
     Call wait(1)
        Load frmSearchStudentsCourse
        Me.Arrange 0
        frmSearchStudentsCourse.Show
    Call wait(0)
End Sub

Private Sub mnuSearchPublisher_Click()
    Call wait(1)
        Load frmSearchBooksPublishers
        Me.Arrange 0
        frmSearchBooksPublishers.Show
    Call wait(0)
End Sub



Private Sub mnuSearchStudentsAdvanced_Click()
    Call wait(1)
        Load frmSearchStudentsLCardNo
        Me.Arrange 0
        frmSearchStudentsLCardNo.Show
    Call wait(0)
End Sub

Private Sub mnuSearchStudentsAlphabetical_Click()
    Call wait(1)
        Load frmSearchStudents
        Me.Arrange 0
        frmSearchStudents.Show
    Call wait(0)
End Sub

Private Sub mnuSearchTeachersAlphabetical_Click()
    Call wait(1)
        Load frmSearchTeachers
        Me.Arrange 0
        frmSearchTeachers.Show
    Call wait(0)
End Sub

Private Sub mnuSearchTeachersLCardNo_Click()
    Call wait(1)
        Load frmSearchTeachersLCardNo
        Me.Arrange 0
        frmSearchTeachersLCardNo.Show
    Call wait(0)
End Sub

Private Sub mnuViewReturn_Click()
    
    Call wait(1)
        Load frmViewR
        Me.Arrange 0
        frmViewR.Show
    Call wait(0)
    
End Sub


Private Sub tbToolBar_ButtonClick(ByVal Button As ComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Insert"
            Call mnuInsertBooks_Click
        Case "Issue"
            Call mnuIssueStudents_Click
        Case "Return"
           Call mnuReturnBooks_Click
        Case "RBook"
       
        Case "RIssue"
       
        Case "Return"
           Call mnuReturnBooks_Click
        Case "SearchBook"
            Call mnuSearchBooksAccessionNo_Click
        Case "Status"
            Call mnuBookStatus_Click
        
    End Select
End Sub

Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub





Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    tbToolBar.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuFileExit_Click()
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub mnuFileClose_Click()
    Unload ActiveForm
End Sub

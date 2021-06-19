VERSION 5.00
Begin VB.Form frmIssueBooks 
   Caption         =   "Issue Books"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8790
   Icon            =   "frmIssueBooks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   8790
   Begin VB.Frame Frame1 
      Caption         =   "Issue Books"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.Frame Frame8 
         Height          =   2175
         Left            =   120
         TabIndex        =   19
         Top             =   2880
         Width           =   8535
         Begin VB.Frame Frame10 
            Caption         =   "Select Issue Details"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1935
            Left            =   120
            TabIndex        =   21
            Top             =   120
            Width           =   3975
            Begin VB.ComboBox Days 
               Height          =   315
               ItemData        =   "frmIssueBooks.frx":0442
               Left            =   2040
               List            =   "frmIssueBooks.frx":0455
               Style           =   2  'Dropdown List
               TabIndex        =   22
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "Return Date"
               Height          =   255
               Index           =   2
               Left            =   2040
               TabIndex        =   33
               Top             =   1440
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "Issue Date"
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "d/M/yy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   3
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   2040
               TabIndex        =   32
               Top             =   1080
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "Issue No"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   2040
               TabIndex        =   31
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label lblLabels 
               Caption         =   "Issue No:"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   26
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label lblLabels 
               Caption         =   "Issue Date (mm/dd/yy):"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   25
               Top             =   1080
               Width           =   1815
            End
            Begin VB.Label lblLabels 
               Caption         =   "Return Date (mm/dd/yy):"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   24
               Top             =   1440
               Width           =   1935
            End
            Begin VB.Label lblLabels 
               Caption         =   "No. of Days:"
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   23
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.Frame Frame9 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   4200
            TabIndex        =   20
            Top             =   600
            Width           =   4215
            Begin VB.CommandButton cmdRefresh 
               Caption         =   "&Refresh"
               Height          =   300
               Left            =   1560
               TabIndex        =   30
               Top             =   360
               Width           =   1095
            End
            Begin VB.CommandButton cmdExit 
               Caption         =   "&Exit"
               Height          =   300
               Left            =   6000
               TabIndex        =   29
               Top             =   840
               Width           =   1095
            End
            Begin VB.CommandButton cmdClose 
               Caption         =   "&Close"
               Height          =   300
               Left            =   2880
               TabIndex        =   28
               Top             =   360
               Width           =   1095
            End
            Begin VB.CommandButton cmdIssue 
               Caption         =   "&Issue"
               Enabled         =   0   'False
               Height          =   300
               Left            =   240
               TabIndex        =   27
               Top             =   360
               Width           =   1095
            End
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2655
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   4095
         Begin VB.Frame Frame6 
            Caption         =   "Book Details"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   120
            TabIndex        =   9
            Top             =   1200
            Width           =   3855
            Begin VB.Label lblBook 
               Caption         =   "Subject :-"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   14
               Top             =   960
               Width           =   3615
            End
            Begin VB.Label lblBook 
               Caption         =   "Author :-"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   13
               Top             =   720
               Width           =   3615
            End
            Begin VB.Label lblBook 
               Caption         =   "Title :-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   12
               Top             =   480
               Width           =   3615
            End
            Begin VB.Label lblBook 
               Caption         =   "Accession Number :-"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   11
               Top             =   240
               Width           =   3615
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Select Book Accession Number"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   3855
            Begin VB.ComboBox AccessionNo 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   360
               Width           =   2175
            End
            Begin VB.Label lbl 
               Caption         =   "AccessionNo"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   6
               Top             =   360
               Width           =   1095
            End
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2655
         Left            =   4320
         TabIndex        =   1
         Top             =   240
         Width           =   4335
         Begin VB.Frame Frame7 
            Caption         =   "Student Details"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   240
            TabIndex        =   10
            Top             =   1200
            Width           =   3975
            Begin VB.Label lblBook 
               Caption         =   "Status :-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   18
               Top             =   960
               Width           =   3735
            End
            Begin VB.Label lblBook 
               Caption         =   "Category :-"
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   17
               Top             =   720
               Width           =   3615
            End
            Begin VB.Label lblBook 
               Caption         =   "Batch :-"
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   16
               Top             =   480
               Width           =   3615
            End
            Begin VB.Label lblBook 
               Caption         =   "Name :-"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   15
               Top             =   240
               Width           =   3615
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Select Student's Library Card Number"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   3975
            Begin VB.ComboBox LCardNo 
               Height          =   315
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   360
               Width           =   2055
            End
            Begin VB.Label lbl 
               Caption         =   "LCardNo"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   3
               Top             =   360
               Width           =   735
            End
         End
      End
   End
End
Attribute VB_Name = "frmIssueBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BookC As Integer
Private Sub AccessionNo_Click()
    On Error Resume Next
    If AccessionNo.ListCount <= 0 Then
        AccessionNo.Enabled = False
        Exit Sub
    End If
    
    AccessionNo.Enabled = True
    
    Screen.MousePointer = 11
    With ObjCon
        
        .Open FileDSN
        
        '==================================================
        'Loading the records for the particular AccessionNo
        '==================================================
        
        query = "select Title, Author,Subject,Category from BookI where AccessionNo=" & CInt(AccessionNo.text)
        Set objrs = .Execute(query)
        
        lblBook(0).Caption = "AccessionNo : " & AccessionNo.text
        lblBook(1).Caption = "Title : " & GetString(objrs(0))
        lblBook(2).Caption = "Authors : " & GetString(objrs(1))
        lblBook(3).Caption = "Subject : " & GetString(objrs(2))
        
        If UCase(objrs(3)) <> UCase("Normal") Then
            Beep
            MsgBox "Reference Books or Periodicals can't be Issued.", vbExclamation, "Warning"
            cmdIssue.Enabled = False
        End If
        
        .Close
    End With
    Screen.MousePointer = 0
End Sub

Private Sub cmdClose_Click()
    Screen.MousePointer = 0
    Unload Me
End Sub
Private Sub AccessionNo_Change()
    Call AccessionNo_Click
End Sub
Private Sub cmdIssue_Click()
    On Error Resume Next
    Screen.MousePointer = 11
    Call Days_Click
    With ObjCon
        .Open FileDSN
        
        '==========================================
        'Inserting Issue Details in the Issue table
        '==========================================
        
        query = "insert into IssueDetails values(" & CInt(Label1(0).Caption) & "," & CInt(AccessionNo.text) & ", '" & LCardNo.text & "','" & CDate(Label1(1).Caption) & "','" & Label1(2).Caption & "')"
        .Execute (query)
        If UCase(Days.text) <> UCase("Semester") Then
            query = "update Students set BookCount=" & (BookC + 1) & " where LCardNo = '" & Trim(LCardNo.text) & "'"
            .Execute (query)
        End If
        
        query = "update Issue set Issued='Yes' where AccessionNo=" & CInt(AccessionNo.text)
        .Execute (query)
        
        Beep
        MsgBox "Book Issued", vbInformation, "Info"
        
        AccessionNo.RemoveItem (AccessionNo.ListIndex)
        AccessionNo.text = AccessionNo.List(0)
        
        .Close
        
    End With
        
    cmdIssue.Enabled = False
    
    Screen.MousePointer = 0
    
End Sub

Private Sub cmdRefresh_Click()
    On Error Resume Next
    Call Form_Load
End Sub

Private Sub Days_Click()
On Error Resume Next
    Screen.MousePointer = 11
    
    With ObjCon
        
        '======================
        'Fetching Issue Details
        '======================
        
        .Open FileDSN
            query = "select max(IssueNo) from IssueDetails"
            Set objrs = .Execute(query)
            
        If IsNull(objrs(0)) Then
            Label1(0).Caption = Val(1)
        Else
            Label1(0).Caption = Val(objrs(0)) + Val(1)
        End If
        
        Label1(1).Caption = Date
        
        If Trim(Days.text) <> "Semester" Then
            Label1(2).Caption = DateAdd("d", Days.text, Date)
        Else
            Label1(2).Caption = "Semester"
        End If
        
        Call LCardNo_Click
        cmdIssue.Enabled = False
        
        .Close
        
    End With
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Screen.MousePointer = 11
    
    With ObjCon
        
        .Open FileDSN
        
        '=====================================
        'Fetching AccessionNo from Issue Table
        '=====================================
        
        query = "select distinct(AccessionNO) from Issue where IssueL='Yes' and Issued='No' order by AccessionNo"
        Set objrs = .Execute(query)
        
        AccessionNo.Clear
        
        If Not objrs.EOF Then
            While Not objrs.EOF
                AccessionNo.AddItem objrs(0)
                objrs.MoveNext
            Wend
            
            AccessionNo.text = AccessionNo.List(0)
            
        Else
            AccessionNo.Enabled = False
            MsgBox "No Books for Issue available", vbExclamation, "Warning"
            .Close
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        .Close
    End With
    
    With ObjCon
    
        .Open FileDSN
        
        '====================================
        'Fetching LCardNo from Students Table
        '====================================
        
        query = "select distinct(LCardNo) from Students order by LCardNo"
        Set objrs = .Execute(query)
        
        LCardNo.Clear
        
        If Not objrs.EOF Then
            
            While Not objrs.EOF
                LCardNo.AddItem objrs(0)
                objrs.MoveNext
            Wend
            
            LCardNo.text = LCardNo.List(0)
         
        Else
        
            LCardNo.Enabled = False
            MsgBox "There are no studentsin the database", vbExclamation, "Warning"
            
        End If
        
        .Close
        
        Screen.MousePointer = 0
        Days.text = Days.List(0)
        
        Call tot(1)
        
    End With
End Sub

Private Sub Form_Resize()
    Me.Height = 5655
    Me.Width = 8910
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = 0
    Unload Me
    Call tot(-1)
End Sub

Private Sub LCardNo_Click()
    On Error Resume Next
    
    Screen.MousePointer = 11
    
    With ObjCon
        
        .Open FileDSN
        
        '==============================================
        'Loading the records for the particular LCardNo
        '==============================================
        
        query = "select Name, Batch, Category, BookCount, DOI,course from Students where LCardNo='" & Trim(LCardNo.text) & "'"
        Set objrs = .Execute(query)
        
        lblBook(4).Caption = "Name : " & GetString(objrs(0))
        lblBook(5).Caption = "Batch : " & GetString(objrs(1))
        lblBook(6).Caption = "Category : " & GetString(objrs(2))
        
        BookC = objrs(3)
        If Days.text = "Semester" Then
        
            cmdIssue.Enabled = True
            lblBook(7).Caption = "Status : Books can be Issued"
        Else
                
            If CInt(objrs(3)) >= 2 Then
                
                lblBook(7).Caption = "Status : Sorry! " & BookC & " Books already Issued"
                .Close
                cmdIssue.Enabled = False
                Screen.MousePointer = 0
                Exit Sub
            Else
                
                cmdIssue.Enabled = True
                lblBook(7).Caption = "Status : Books can be Issued"
                
            End If
        
                
        End If
        
        
        
        Dim str As String
        Dim d As Date
               
        d = DateAdd("yyyy", 2, objrs(4))
       
        If d < Date And UCase(objrs(5)) <> UCase("MCA") Then
            Beep
            MsgBox "Library card expired, Book can't be Issued.", vbExclamation, "Sorry"
            .Close
            cmdIssue.Enabled = False
            Screen.MousePointer = 0
            Exit Sub
        Else
            d = DateAdd("yyyy", 3, objrs(4))
            If d < Date And UCase(objrs(5)) = UCase("MCA") Then
                Beep
                MsgBox "Library card expired, Book can't be Issued.", vbExclamation, "Sorry"
                .Close
                cmdIssue.Enabled = False
                Screen.MousePointer = 0
                Exit Sub
            Else
                cmdIssue.Enabled = True
            End If
        End If
        
        .Close
        
        cmdIssue.Enabled = True
        
    End With
    
    Screen.MousePointer = 0
End Sub

VERSION 5.00
Begin VB.Form frmReturnBooksStudents 
   Caption         =   "Return Books"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8790
   Icon            =   "frmReturnBooksStudents.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   8790
   Begin VB.Frame Frame1 
      Caption         =   "Return Books "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.Frame Frame8 
         Height          =   1695
         Left            =   120
         TabIndex        =   13
         Top             =   2880
         Width           =   8535
         Begin VB.Frame Frame2 
            Caption         =   "Frame2"
            Height          =   1455
            Left            =   120
            TabIndex        =   25
            Top             =   120
            Width           =   3975
            Begin VB.ComboBox Fine 
               Height          =   315
               ItemData        =   "frmReturnBooksStudents.frx":0442
               Left            =   1800
               List            =   "frmReturnBooksStudents.frx":044C
               Style           =   2  'Dropdown List
               TabIndex        =   27
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label lblLabels 
               Caption         =   "Fine (@Rs 2/day):"
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   33
               Top             =   1080
               Width           =   1455
            End
            Begin VB.Label Label3 
               Caption         =   "Label3"
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
               Left            =   1800
               TabIndex        =   32
               Top             =   1080
               Width           =   2055
            End
            Begin VB.Label Label2 
               Caption         =   "Label2"
               Height          =   255
               Left            =   1800
               TabIndex        =   31
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label lblLabels 
               Caption         =   "Fine Status:"
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   28
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label lblLabels 
               Caption         =   "Return Date (mm/dd/yy):"
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Width           =   1575
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
            Height          =   1095
            Left            =   4200
            TabIndex        =   14
            Top             =   360
            Width           =   4215
            Begin VB.CommandButton cmdReturn 
               Caption         =   "&Return"
               Height          =   375
               Left            =   120
               TabIndex        =   29
               Top             =   240
               Width           =   3975
            End
            Begin VB.CommandButton cmdClose 
               Caption         =   "&Close"
               Height          =   300
               Left            =   3000
               TabIndex        =   17
               Top             =   720
               Width           =   1095
            End
            Begin VB.CommandButton cmd 
               Caption         =   "&Exit"
               Height          =   300
               Left            =   6000
               TabIndex        =   16
               Top             =   840
               Width           =   1095
            End
            Begin VB.CommandButton cmdRefresh 
               Caption         =   "&Refresh"
               Height          =   300
               Left            =   120
               TabIndex        =   15
               Top             =   720
               Width           =   1095
            End
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2655
         Left            =   120
         TabIndex        =   2
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
            TabIndex        =   6
            Top             =   1200
            Width           =   3855
            Begin VB.Label lblBook 
               Caption         =   "Subject :-"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   11
               Top             =   960
               Width           =   3615
            End
            Begin VB.Label lblBook 
               Caption         =   "Author :-"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   10
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
               TabIndex        =   9
               Top             =   480
               Width           =   3615
            End
            Begin VB.Label lblBook 
               Caption         =   "Accession Number :-"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   8
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
            TabIndex        =   3
            Top             =   240
            Width           =   3855
            Begin VB.ComboBox AccessionNo 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   360
               Width           =   2175
            End
            Begin VB.Label lbl 
               Caption         =   "AccessionNo"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   4
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
         Begin VB.Frame Frame10 
            Caption         =   "Issue Details"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   3975
            Begin VB.Label lblLabels 
               Caption         =   "OReturn Date (mm/dd/yy):"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   24
               Top             =   720
               Width           =   1935
            End
            Begin VB.Label lblLabels 
               Caption         =   "Issue Date (mm/dd/yy):"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   23
               Top             =   480
               Width           =   1815
            End
            Begin VB.Label lblLabels 
               Caption         =   "Issue No:"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   22
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "Issue No"
               Height          =   255
               Index           =   0
               Left            =   2040
               TabIndex        =   21
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "Issue Date"
               Height          =   255
               Index           =   1
               Left            =   2040
               TabIndex        =   20
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "Return Date"
               Height          =   255
               Index           =   2
               Left            =   2040
               TabIndex        =   19
               Top             =   720
               Width           =   1455
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Details :-"
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
            TabIndex        =   7
            Top             =   1200
            Width           =   3975
            Begin VB.Label lblBook 
               Caption         =   "LCardNo :-"
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
               Index           =   6
               Left            =   120
               TabIndex        =   34
               Top             =   360
               Width           =   975
            End
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
               Index           =   5
               Left            =   120
               TabIndex        =   30
               Top             =   720
               Width           =   3615
            End
            Begin VB.Label lblBook 
               Caption         =   "LCardNo "
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
               Index           =   4
               Left            =   1200
               TabIndex        =   12
               Top             =   360
               Width           =   1695
            End
         End
      End
   End
End
Attribute VB_Name = "frmReturnBooksStudents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim str As String
Dim f As Integer

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
        
        query = "select Title, Author,Subject from BookI where AccessionNo=" & CInt(AccessionNo.Text)
        Set objrs = .Execute(query)
        
        lblBook(0).Caption = "AccessionNo : " & AccessionNo.Text
        lblBook(1).Caption = "Title : " & GetString(objrs(0))
        lblBook(2).Caption = "Authors : " & GetString(objrs(1))
        lblBook(3).Caption = "Subject : " & GetString(objrs(2))
        
        .Close
    End With
    
    
    Screen.MousePointer = 11
    With ObjCon
        
        .Open FileDSN
        
        '=============================================
        'Loading Records for the LCardNo from Students
        '=============================================
        
        query = "select * from IssueDetails where AccessionNo=" & CInt(AccessionNo.Text)
        Set objrs = .Execute(query)
        
        Label1(0).Caption = objrs(0)
        Label1(1).Caption = objrs(3)
        Label1(2).Caption = objrs(4)
        
        lblBook(4).Caption = objrs(2)
        
        If Label1(2).Caption = "" Then
            str = "Teacher"
            lblBook(5).Caption = "Status : Teacher"
            Fine.Text = Fine.List(0)
            Fine.Enabled = False
            
        Else
            str = "Student"
            lblBook(5).Caption = "Status : Student"
            Fine.Enabled = True
            Fine.Text = Fine.List(0)
        End If
        
        
        .Close
    End With
    
    
    Label2.Caption = Date
    
    
    Screen.MousePointer = 0
End Sub

Private Sub cmdClose_Click()
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub cmdIssue_Click()
On Error Resume Next
    If str = "Teacher" Then
        Load frmIssueBooksTeachers
        
        frmIssueBooksTeachers.Show
        frmIssueBooksTeachers.SetFocus
    Else
        Load frmIssueBooks
       
        frmIssueBooks.Show
        frmIssueBooks.SetFocus
    End If
End Sub

Private Sub cmdRefresh_Click()
    On Error Resume Next
    Call Form_Load
End Sub


Private Sub cmdReturn_Click()
    On Error Resume Next
    
    With ObjCon
        
        .Open FileDSN
        
        '===============================================
        'Deleting Issue Enteries from issueDetails Table
        '===============================================
        
        query = "delete from IssueDetails where AccessionNo=" & CInt(AccessionNo.Text)
        .Execute (query)
        
        '=============================================
        'Putting return details in ReturnDetails Table
        '=============================================
        
        query = "insert into ReturnDetails values(" & AccessionNo.Text & ",'" & Trim(lblBook(4).Caption) & "','" & Trim(Label2.Caption) & "'," & f & ")"
        .Execute (query)
        
        '======================================================
        'Updating Issued Sttaus of the Book in the Issued Table
        '======================================================
        
        query = "update Issue set Issued='No' where AccessionNo=" & CInt(AccessionNo.Text)
        .Execute (query)
        
        If UCase(str) = UCase("Student") Then
        
            '====================================
            'Updating the BookCount of the Student
            '=====================================
            query = "select BookCount from Students where LCardNo='" & Trim(lblBook(4).Caption) & "'"
            Set objrs = .Execute(query)
            
            If objrs(0) > 0 Then
                Dim BookC As Integer
                BookC = Val(objrs(0))
                BookC = BookC - 1
                
                query = "update Students set BookCount=" & CInt(BookC)
                .Execute (query)
            End If
        End If
        
        Beep
        MsgBox "Book Returned", vbInformation, "Info"
        
        AccessionNo.RemoveItem (AccessionNo.ListIndex)
        AccessionNo.Text = AccessionNo.List(0)
        
        .Close
    End With
    Screen.MousePointer = 0
End Sub

Private Sub Fine_Click()
On Error Resume Next
    If Fine.Text = Trim("Cancel") Then
        f = 0
    Else
        Dim d As Integer
        d = DateDiff("d", CDate(Label1(2).Caption), Date)
        If d > 0 Then
            f = 2 * d
        Else
            f = 0
        End If
    End If
    
    Label3.Caption = "Rs " & f
End Sub

Private Sub Form_Load()

On Error Resume Next
    Screen.MousePointer = 11
    
    With ObjCon
        
        .Open FileDSN
        
        '============================================
        'Fetching AccessionNo from IssueDetails Table
        '============================================
        
        query = "select distinct(AccessionNO) from IssueDetails  order by AccessionNo"
        Set objrs = .Execute(query)
        
        AccessionNo.Clear
        
        If Not objrs.EOF Then
            While Not objrs.EOF
                AccessionNo.AddItem objrs(0)
                objrs.MoveNext
            Wend
            
            AccessionNo.Text = AccessionNo.List(0)
            cmdReturn.Enabled = True
        Else
            AccessionNo.Enabled = False
            MsgBox "No Books Issued", vbExclamation, "Warning"
            cmdReturn.Enabled = False
            .Close
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        .Close
        
    End With
    
    Call tot(1)
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    Me.Height = 5175
    Me.Width = 8910
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = 0
    Call tot(-1)
End Sub


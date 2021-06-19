VERSION 5.00
Begin VB.Form frmIssueBooksTeachers 
   Caption         =   "Issue Books Teachers"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8790
   Icon            =   "frmIssueBooksTeachers.frx":0000
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
            Height          =   1935
            Left            =   120
            TabIndex        =   21
            Top             =   120
            Width           =   3975
            Begin VB.Label Label1 
               Caption         =   "Issue Date"
               Height          =   255
               Index           =   1
               Left            =   1440
               TabIndex        =   29
               Top             =   720
               Width           =   2295
            End
            Begin VB.Label Label1 
               Caption         =   "IssueNo"
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
               Left            =   1440
               TabIndex        =   28
               Top             =   360
               Width           =   2295
            End
            Begin VB.Label lblLabels 
               Caption         =   "Issue No:"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   23
               Top             =   360
               Width           =   1455
            End
            Begin VB.Label lblLabels 
               Caption         =   "Issue Date:"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   22
               Top             =   720
               Width           =   1455
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
            Begin VB.CommandButton cmdClose 
               Caption         =   "&Close"
               Height          =   300
               Left            =   2880
               TabIndex        =   27
               Top             =   360
               Width           =   1095
            End
            Begin VB.CommandButton cmdExit 
               Caption         =   "&Exit"
               Height          =   300
               Left            =   6000
               TabIndex        =   26
               Top             =   840
               Width           =   1095
            End
            Begin VB.CommandButton cmdRefresh 
               Caption         =   "&Refresh"
               Height          =   300
               Left            =   1560
               TabIndex        =   25
               Top             =   360
               Width           =   1095
            End
            Begin VB.CommandButton cmdIssue 
               Caption         =   "&Issue"
               Height          =   300
               Left            =   120
               TabIndex        =   24
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
            Caption         =   "Teacher Details"
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
               Caption         =   "Subject :-"
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   17
               Top             =   720
               Width           =   3615
            End
            Begin VB.Label lblBook 
               Caption         =   "PhoneNo :-"
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   16
               Top             =   480
               Width           =   3615
            End
            Begin VB.Label lblBook 
               Caption         =   "Name :-"
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
               Left            =   120
               TabIndex        =   15
               Top             =   240
               Width           =   3615
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Select Teacher's Library Card Number"
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
Attribute VB_Name = "frmIssueBooksTeachers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AccessionNo_Change()
    Call AccessionNo_Click
End Sub

Private Sub Form_Resize()
    Me.Height = 5655
    Me.Width = 8910
End Sub

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
       
        cmdIssue.Enabled = True
         If UCase(objrs(3)) <> UCase("Normal") Then
            Beep
            MsgBox "Reference Books or Periodicals can't be Issued.", vbExclamation, "Warning"
            cmdIssue.Enabled = False
        End If
        Call Days_Click
        .Close
        
    End With
    
   
    Screen.MousePointer = 0
End Sub

Private Sub cmdClose_Click()
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub cmdIssue_Click()
    On Error Resume Next
    Screen.MousePointer = 11
    With ObjCon
        .Open FileDSN
        
        '==========================================
        'Inserting Issue Details in the Issue table
        '==========================================
        
        query = "insert into IssueDetails values(" & CInt(Label1(0).Caption) & "," & CInt(AccessionNo.text) & ", '" & LCardNo.text & "','" & CDate(Trim(Label1(1).Caption)) & "','')"
        .Execute (query)
        
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
        
        query = "select distinct(LCardNo) from Teachers order by LCardNo"
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
            MsgBox "There are no Teachers in the database", vbExclamation, "Warning"
            
        End If
        
        .Close
        Screen.MousePointer = 0
        Call tot(1)
    End With
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
        
        query = "select Name,Phone,Subject from Teachers where LCardNo='" & Trim(LCardNo.text) & "'"
        Set objrs = .Execute(query)
        
        
        lblBook(4).Caption = "Name : " & GetString(objrs(0))
        lblBook(5).Caption = "PhoneNo : " & GetString(objrs(1))
        lblBook(6).Caption = "Subject : " & GetString(objrs(2))
        lblBook(7).Caption = "Status : Book can be Issued"
        
        .Close
        
       ' cmdIssue.Enabled = True
        
    End With
    
    Screen.MousePointer = 0
End Sub

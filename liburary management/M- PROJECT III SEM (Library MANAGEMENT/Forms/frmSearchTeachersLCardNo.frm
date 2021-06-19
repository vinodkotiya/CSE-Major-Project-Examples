VERSION 5.00
Begin VB.Form frmSearchTeachersLCardNo 
   Caption         =   "Search Teachers :- Library Card No"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7395
   Icon            =   "frmSearchTeachersLCardNo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   7395
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.Frame Frame7 
         Caption         =   "Teacher's Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   3000
         TabIndex        =   5
         Top             =   240
         Width           =   4215
         Begin VB.CommandButton cmdExit 
            Caption         =   "&Close"
            Height          =   300
            Left            =   2880
            TabIndex        =   7
            Top             =   2280
            Width           =   1095
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "&Refresh"
            Height          =   300
            Left            =   1560
            TabIndex        =   6
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label lblBook 
            Caption         =   "Subject :"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   3615
         End
         Begin VB.Label lblBook 
            Caption         =   "Name :"
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
            Caption         =   "LCardNo :"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   3615
         End
         Begin VB.Label lblBook 
            Caption         =   "Address :"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   3615
         End
         Begin VB.Label lblBook 
            Caption         =   "Phone :-"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   9
            Top             =   1200
            Width           =   3615
         End
         Begin VB.Label lblBook 
            Caption         =   "Books Issued :-"
            Height          =   735
            Index           =   5
            Left            =   120
            TabIndex        =   8
            Top             =   1440
            Width           =   3975
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Select and Click"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2775
         Begin VB.CommandButton cmdSearch 
            Caption         =   "&Search"
            Enabled         =   0   'False
            Height          =   375
            Left            =   240
            TabIndex        =   3
            Top             =   2160
            Width           =   2055
         End
         Begin VB.ListBox List1 
            Height          =   1425
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label TotRecords 
            Caption         =   "Total Records Found :-"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   1800
            Width           =   2535
         End
      End
   End
End
Attribute VB_Name = "frmSearchTeachersLCardNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdExit_Click()
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub Command1_Click()
   
End Sub

Private Sub cmdRefresh_Click()
    Screen.MousePointer = 11
        Call Form_Load
    Screen.MousePointer = 0
End Sub

Private Sub cmdSearch_Click()
     Call List1_Click
End Sub

Private Sub Form_Resize()
    Me.Height = 3510
    Me.Width = 7515
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call tot(-1)
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Screen.MousePointer = 11
    
    With ObjCon
        
        .Open FileDSN
        
        '=========================================
        'Selecting distinct accessionNo from BookI
        '=========================================
        
        query = "select distinct(LCardNo) from Teachers order by LCardNo"
        Set objrs = .Execute(query)
        
        List1.Enabled = True
        List1.Clear
        
        If Not objrs.EOF Then
            
            While Not objrs.EOF
                List1.AddItem objrs(0)
                objrs.MoveNext
            Wend
            List1.text = List1.List(0)
            TotRecords.Caption = "Total Records Found : " & List1.ListCount
            cmdSearch.Enabled = True
        Else
            
            List1.AddItem "Empty"
            List1.Enabled = False
            
            TotRecords.Caption = "Total Records Found : " & List1.ListCount
            
        End If
        
       .Close
        
    End With
    
    Screen.MousePointer = 0
    
    Call tot(1)
End Sub

Private Sub List1_Click()
    
    On Error Resume Next
    
    Screen.MousePointer = 11
    
    '=======================================
    'Fetching records for particular LCardNo
    '=======================================
    
    With ObjCon
        
        .Open FileDSN
      
        query = "select Name,Subject,Address,Phone from Teachers where LCardNo='" & Trim(List1.text) & "'"
        Set objrs = .Execute(query)
        
        If Not objrs.EOF Then
            lblBook(0).Caption = "LCardNo : " & List1.text
            lblBook(1).Caption = "Name : " & objrs(0)
            lblBook(2).Caption = "Subject :" & objrs(1)
            lblBook(3).Caption = "Address : " & objrs(2)
            lblBook(4).Caption = "Phone : " & objrs(3)
        End If
        
        Dim objrs1 As Recordset
        
        query = "select AccessionNo from IssueDetails where LCardNo='" & Trim(List1.text) & "'"
        Set objrs1 = .Execute(query)
        
        If Not objrs1.EOF Then
            lblBook(5).Caption = "AccessionNo of Books Issued : "
            While Not objrs1.EOF
                lblBook(5).Caption = lblBook(5).Caption & objrs1(0) & " , "
                objrs1.MoveNext
            Wend
        Else
            lblBook(5).Caption = "Books Issued : None"
        End If
        
        Set objrs1 = Nothing
        .Close
        
    End With
    
    Screen.MousePointer = 0
End Sub

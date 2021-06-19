VERSION 5.00
Begin VB.Form frmSearchBooks 
   Caption         =   "Search Books :- AccessionNo"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7395
   Icon            =   "frmSearchBooks.frx":0000
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
         Height          =   2655
         Left            =   3000
         TabIndex        =   3
         Top             =   240
         Width           =   4215
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "&Refresh"
            Height          =   300
            Left            =   1560
            TabIndex        =   11
            Top             =   2280
            Width           =   1095
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "&Close"
            Height          =   300
            Left            =   2880
            TabIndex        =   8
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label lblBook 
            Caption         =   "Category :-"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   3615
         End
         Begin VB.Label lblBook 
            Caption         =   "AcessionNo :-"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   6
            Top             =   240
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
            Index           =   5
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   3615
         End
         Begin VB.Label lblBook 
            Caption         =   "Authors :-"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   4
            Top             =   720
            Width           =   3615
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
            TabIndex        =   9
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
            TabIndex        =   10
            Top             =   1800
            Width           =   2535
         End
      End
   End
End
Attribute VB_Name = "frmSearchBooks"
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
        
        query = "select distinct(AccessionNO) from BookI order by AccessionNo"
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
    
    '===========================================
    'Fetching records for particular AccessionNo
    '===========================================
    
    With ObjCon
        
        .Open FileDSN
      
        query = "select Title,Author,Category from BookI where AccessionNo=" & CInt(Trim(List1.List(List1.ListIndex)))
        Set objrs = .Execute(query)
        
        lblBook(4).Caption = "AcessionNo :-"
        lblBook(0).Caption = "Category :-"
        lblBook(5).Caption = "Title :-"
        lblBook(6).Caption = "Authors :-"
        
        lblBook(4).Caption = lblBook(4).Caption & "   " & List1.List(List1.ListIndex)
        
        If Not objrs.EOF Then
            lblBook(0).Caption = lblBook(0).Caption & "   " & GetString(objrs(2))
            lblBook(5).Caption = lblBook(5).Caption & "   " & GetString(objrs(0))
            lblBook(6).Caption = lblBook(6).Caption & "   " & GetString(objrs(1))
        End If
        
        .Close
        
    End With
    
    Screen.MousePointer = 0
End Sub

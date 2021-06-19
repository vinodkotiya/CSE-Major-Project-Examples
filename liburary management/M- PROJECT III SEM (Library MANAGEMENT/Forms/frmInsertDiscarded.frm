VERSION 5.00
Begin VB.Form frmInsertDiscarded 
   Caption         =   "Discarded & Missing Books"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   Icon            =   "frmInsertDiscarded.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4725
   Begin VB.Frame Frame1 
      Caption         =   "Insert Discarded and Missing Books"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.Frame Frame3 
         Caption         =   "Select "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4455
         Begin VB.Frame Frame5 
            Height          =   1095
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   4215
            Begin VB.OptionButton Option2 
               Caption         =   "Missing"
               Height          =   255
               Left            =   2040
               TabIndex        =   12
               Top             =   720
               Width           =   2055
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Discarded"
               Height          =   195
               Left            =   120
               TabIndex        =   11
               Top             =   720
               Width           =   1575
            End
            Begin VB.ComboBox AccessionNo 
               Height          =   315
               Left            =   1920
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   240
               Width           =   2175
            End
            Begin VB.Label Label1 
               Caption         =   "Accession No."
               Height          =   255
               Left            =   120
               TabIndex        =   10
               Top             =   240
               Width           =   1455
            End
         End
         Begin VB.Frame Frame4 
            Height          =   615
            Left            =   120
            TabIndex        =   3
            Top             =   1320
            Width           =   4215
            Begin VB.Label Label3 
               Caption         =   "Title of the Book"
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
               Left            =   1080
               TabIndex        =   5
               Top             =   240
               Width           =   3015
            End
            Begin VB.Label Label2 
               Caption         =   "Title :-"
               Height          =   255
               Left            =   120
               TabIndex        =   4
               Top             =   240
               Width           =   735
            End
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   2280
         Width           =   4455
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "&Refresh"
            Height          =   300
            Left            =   480
            TabIndex        =   13
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton cmdClose 
            Caption         =   "&Close"
            Height          =   300
            Left            =   2280
            TabIndex        =   7
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton cmdInsert 
            Caption         =   "&Add to Selected List"
            Enabled         =   0   'False
            Height          =   375
            Left            =   480
            TabIndex        =   6
            Top             =   240
            Width           =   3135
         End
      End
   End
End
Attribute VB_Name = "frmInsertDiscarded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AccessionNo_Click()
    
    On Error Resume Next
    Screen.MousePointer = 11
    
    With ObjCon
    
        '============================
        'Getting the title from BookI
        '============================
        
        .Open FileDSN
        
            query = "select Title from BookI where AccessionNo=" & CInt(AccessionNo.text)
            Set objrs = .Execute(query)
            
            Label3.Caption = GetString(objrs(0))
            
            Screen.MousePointer = 0
            cmdInsert.Enabled = True
            
    .Close
        
    End With
    
End Sub

Private Sub cmdClose_Click()
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub cmdInsert_Click()
    On Error Resume Next
    Screen.MousePointer = 11
    
    With ObjCon
        
        .Open FileDSN
        
        query = "Select AccessionNo from IssueDetails where AccessionNo=" & CInt(AccessionNo.text)
        Set objrs = .Execute(query)
        
        If Not objrs.EOF Then
            Beep
            MsgBox "This book is already Issued. Information can't be updated", vbExclamation, "Sorry"
            Screen.MousePointer = 0
            cmdInsert.Enabled = False
            .Close
            Exit Sub
        End If
        
        '===================================
        'Updating discarded or missing lists
        '===================================
        
        If Option1.value = True Then
            query = "update Issue set IssueL='No',discarded='Yes',missing='No' where AccessionNo=" & CInt(AccessionNo.text)
        Else
            If Option2.value = True Then
                query = "update Issue set IssueL='No',discarded='No',missing='Yes' where AccessionNo=" & CInt(AccessionNo.text)
            End If
        End If
        
        .Execute (query)
        AccessionNo.RemoveItem (AccessionNo.ListIndex)
        AccessionNo.text = AccessionNo.List(0)
        
        Beep
        MsgBox "Selected List updated", vbInformation, "Congratulation"
        
        .Close
    End With
    Screen.MousePointer = 0
    
End Sub

Private Sub cmdRefresh_Click()
    Call Form_Load
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Screen.MousePointer = 11
    
    With ObjCon
        
        .Open FileDSN
        
        '=========================================
        'Selecting distinct accessionNo from BookI
        '=========================================
        
        query = "select distinct(AccessionNO) from Issue where IssueL='Yes' order by AccessionNo"
        Set objrs = .Execute(query)
        
        AccessionNo.Enabled = True
        AccessionNo.Clear
        
        If Not objrs.EOF Then
            
            While Not objrs.EOF
                AccessionNo.AddItem objrs(0)
                objrs.MoveNext
            Wend
            
            AccessionNo.text = AccessionNo.List(0)
            
            cmdSearch.Enabled = True
        Else
            
            AccessionNo.AddItem "Empty"
            AccessionNo.Enabled = False
            
            TotRecords.Caption = "Total Records Found : " & AccessionNo.ListCount
            
        End If
        
       .Close
        
    End With
    
    Screen.MousePointer = 0
    
    Call tot(1)
End Sub

Private Sub Form_Resize()
    Me.Height = 3900
    Me.Width = 4845
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = 0
    Call tot(-1)
End Sub

Private Sub Option1_Click()
    cmdInsert.Enabled = True
End Sub

Private Sub Option2_Click()
    cmdInsert.Enabled = True
End Sub

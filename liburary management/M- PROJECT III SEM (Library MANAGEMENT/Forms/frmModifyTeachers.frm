VERSION 5.00
Begin VB.Form frmModifyTeachers 
   Caption         =   "Modify Teachers Details"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   Icon            =   "frmModifyTeachers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6915
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   300
         Left            =   3360
         TabIndex        =   20
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   4500
         TabIndex        =   19
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Update"
         Height          =   300
         Left            =   3360
         TabIndex        =   18
         Top             =   3480
         Width           =   3375
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   5640
         TabIndex        =   17
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Frame Frame4 
         Caption         =   "Select Library Card Number"
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
         TabIndex        =   14
         Top             =   240
         Width           =   6495
         Begin VB.ComboBox LCardNo 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lbl 
            Caption         =   "LCardNo"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Library Details"
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
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   3015
         Begin VB.TextBox text 
            Height          =   285
            Index           =   2
            Left            =   1200
            TabIndex        =   10
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox text 
            Height          =   285
            Index           =   1
            Left            =   1200
            TabIndex        =   9
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   1
            Left            =   4320
            TabIndex        =   3
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lbl 
            Caption         =   "Date of Issue"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   5
            Top             =   960
            Width           =   975
         End
         Begin VB.Label lbl 
            Caption         =   "Subject"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Personal Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   3360
         TabIndex        =   1
         Top             =   1200
         Width           =   3375
         Begin VB.TextBox text 
            Height          =   285
            Index           =   5
            Left            =   1320
            TabIndex        =   13
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox text 
            Height          =   645
            Index           =   4
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox text 
            Height          =   285
            Index           =   3
            Left            =   1320
            TabIndex        =   11
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lbl 
            Caption         =   "PhoneNo"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   8
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label lbl 
            Caption         =   "Address"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lbl 
            Caption         =   "Name"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "frmModifyTeachers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub cmdClose_Click()
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
    Screen.MousePointer = 11
    With ObjCon
        
        .Open FileDSN
        
        '=================================================
        'Checking whether Book is Issued to student or not
        '=================================================
        query = "select LCardNo from IssueDetails where LCardNo='" & LCardNo.text & "'"
        Set objrs = .Execute(query)
        
        If Not objrs.EOF Then
            
            Beep
            MsgBox "This Record can't be deleted, as Books are Issued on this Card.", vbExclamation, "Warning"
            cmdDelete.Enabled = False
        Else
            
            '===================================
            'Deleting Record from Students Table
            '===================================
            
            query = "delete from Teachers where LCardNo='" & LCardNo.text & "'"
            .Execute (query)
            
             Beep
            MsgBox "Record Deleted", vbInformation, "Info"
            
            LCardNo.RemoveItem (LCardNo.ListIndex)
            LCardNo.text = LCardNo.List(0)
            cmdDelete.Enabled = False
        End If
        
        .Close
        
    End With
    Screen.MousePointer = 0
End Sub

Private Sub cmdReset_Click()
    Dim i As Integer
    For i = 0 To text.Count - 1
        text(i) = ""
    Next
End Sub

Private Sub cmdModify_Click()
 Screen.MousePointer = 11
    
    '=====================
    'Validating User Input
    '=====================
    
    Dim i As Integer
    
    For i = 1 To text.Count - 1
        If text(i) = "" Then
            Beep
            MsgBox "This field can't be empty, Please insert Data in the " & lbl(i).Caption & " field", vbCritical, "Error"
            Screen.MousePointer = 0
            text(i).SetFocus
            Exit Sub
        End If
    Next
    
    If Len(text(2)) > 8 Or Len(text(2)) < 6 Then
        Beep
        MsgBox "Invalid date entry, Date should be entered in dd/mm/yy Format.", vbCritical, "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    If Len(text(4)) > 255 Then
        Beep
        MsgBox "Address field can contain a maximum of 255 characters.", vbCritical, "Error"
        text(4).SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    With ObjCon
    
        .Open FileDSN
                          
            '==================================
            'Updating Information into Teachers
            '==================================
            query = "update teachers set Subject='" & Trim(text(1)) & "',Name='" & Trim(text(3)) & "',Address='" & SetString(Trim(text(4))) & "',Phone='" & Trim(text(5)) & "',DOI='" & Trim(text(2)) & "' where LCardNo='" & LCardNo.text & "'"
            .Execute (query)
            
            Beep
            MsgBox "Record Updated", vbInformation, "Congratulations"
            
        
        .Close
    End With
    
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()

    On Error Resume Next
    
    Screen.MousePointer = 11
    
    '======================================
    'Getting distinct LCardNo from Students
    '======================================
    
    With ObjCon
    
        .Open FileDSN
        
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
            LCardNo.AddItem "Empty"
            LCardNo.Enabled = False
        End If
        
        .Close
        
    End With
    
    Category.text = Category.List(0)
    Call tot(1)
    
    Screen.MousePointer = 0

End Sub

Private Sub Form_Resize()
    Me.Height = 4785
    Me.Width = 7035
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = 0
    Call tot(-1)
    Unload Me
End Sub

Private Sub LCardNo_Click()
On Error Resume Next
    Screen.MousePointer = 11
        
        With ObjCon
        
            .Open FileDSN
            
            '===========================================
            'Fetching records for the particular LCardNo
            '===========================================
            
            query = "select * from Teachers where LCardNo='" & Trim(LCardNo.text) & "'"
            Set objrs = .Execute(query)
            
            If Not objrs.EOF Then
            
                text(1) = GetString(objrs(1))
                text(2) = GetString(objrs(5))
                text(3) = GetString(objrs(2))
                text(4) = GetString(objrs(3))
                text(5) = GetString(objrs(4))
                
                
            End If
            
            .Close
        
        End With
        
    Screen.MousePointer = 0
    cmdDelete.Enabled = True
End Sub

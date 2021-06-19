VERSION 5.00
Begin VB.Form frmModifyStudents 
   Caption         =   "Modify Students Details"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   Icon            =   "frmModifyStudents.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4455
   ScaleMode       =   0  'User
   ScaleWidth      =   7035
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6855
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   300
         Left            =   3360
         TabIndex        =   9
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   300
         Left            =   4500
         TabIndex        =   10
         Top             =   3480
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
         TabIndex        =   23
         Top             =   240
         Width           =   6495
         Begin VB.ComboBox LCardNo 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lbl 
            Caption         =   "LCardNo"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   24
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
         Height          =   2775
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   3015
         Begin VB.TextBox text 
            Height          =   285
            Index           =   3
            Left            =   1200
            TabIndex        =   4
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox text 
            Height          =   285
            Index           =   2
            Left            =   1200
            TabIndex        =   2
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox text 
            Height          =   285
            Index           =   1
            Left            =   1200
            TabIndex        =   1
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   1
            Left            =   4320
            TabIndex        =   18
            Top             =   480
            Width           =   1695
         End
         Begin VB.ComboBox Category 
            Height          =   315
            ItemData        =   "frmModifyStudents.frx":0442
            Left            =   1200
            List            =   "frmModifyStudents.frx":044F
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1200
            Width           =   1695
         End
         Begin VB.Label lbl 
            Caption         =   "Batch"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lbl 
            Caption         =   "Course"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lbl 
            Caption         =   "Category"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   20
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label lbl 
            Caption         =   "Date of Issue"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   19
            Top             =   1800
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   5640
         TabIndex        =   11
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Update Record"
         Height          =   300
         Left            =   3360
         TabIndex        =   8
         Top             =   3120
         Width           =   3375
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
         Height          =   1815
         Left            =   3360
         TabIndex        =   13
         Top             =   1080
         Width           =   3375
         Begin VB.TextBox text 
            Height          =   525
            Index           =   6
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox text 
            Height          =   285
            Index           =   5
            Left            =   1320
            TabIndex        =   6
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox text 
            Height          =   285
            Index           =   4
            Left            =   1320
            TabIndex        =   5
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lbl 
            Caption         =   "RollNo"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lbl 
            Caption         =   "Address"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   15
            Top             =   1200
            Width           =   735
         End
         Begin VB.Label lbl 
            Caption         =   "Name"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "frmModifyStudents"
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
            
            query = "delete from Students where LCardNo='" & LCardNo.text & "'"
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

Private Sub cmdModify_Click()
    'On error resume next
    Screen.MousePointer = 11
        '=====================
        'Validating user Input
        '=====================
        
        For i = 1 To text.Count - 1
        If text(i) = "" Then
            Beep
            MsgBox "This field can't be empty, Please enter data into " & lbl(i).Caption & " field", vbCritical, "Error"
            text(i).SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
        Next
        
        If Category.text = "" Then
            Beep
            MsgBox "This field can't be empty, Please select one option", vbCritical, "Error"
            Category.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        If text(2) = "" Then
            Beep
            MsgBox "This field can't be empty.", vbCritical, "Error"
            text(2).SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        If Len(text(3)) > 8 Or Len(text(3)) < 6 Then
            Beep
            MsgBox "Invalid date entry, Date should be entered in dd/mm/yy Format.", vbCritical, "Error"
            Category.SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        If Len(text(6)) > 255 Then
            Beep
            MsgBox "Address field can contain a maximum of 255 characters.", vbCritical, "Error"
            text(6).SetFocus
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        '========================
        'Updating the Information
        '========================
        
        With ObjCon
        
            .Open FileDSN
            
                query = "delete from Students where LCardNo='" & Trim(LCardNo.text) & "'"
                .Execute (query)
                
                query = "insert into students values('" & SetString(Trim(LCardNo.text)) & "','" & SetString(Trim(text(1))) & "','" & SetString(Trim(text(2))) & "','" & SetString(Trim(Category.text)) & "','" & SetString(Trim(text(4))) & "','" & SetString(Trim(text(5))) & "','" & SetString(Trim(text(6))) & "','" & Trim(text(3)) & "')"
                .Execute (query)
                
                Beep
                MsgBox "Records Updated", vbInformation, "Info"
                
            .Close
            
            Call cmdReset_Click
            LCardNo.text = LCardNo.List(0)
                
            
            
        End With
        
    Screen.MousePointer = 0
End Sub

Private Sub cmdReset_Click()
    Dim i As Integer
    For i = 1 To text.Count - 1
        text(i) = ""
    Next
End Sub

Private Sub cmdRefresh_Click()
    Call Form_Load
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    
    Screen.MousePointer = 11
    
    '======================================
    'Getting distinct LCardNo from Students
    '======================================
    
    With ObjCon
    
        .Open FileDSN
        
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
    Me.Height = 4455
    Me.Width = 7035
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call tot(-1)
End Sub


Private Sub LCardNo_Click()
    On Error Resume Next
    Screen.MousePointer = 11
        
        With ObjCon
        
            .Open FileDSN
            
            '===========================================
            'Fetching records for the particular LCardNo
            '===========================================
            
            query = "select * from Students where LCardNo='" & Trim(LCardNo.text) & "'"
            Set objrs = .Execute(query)
            
            If Not objrs.EOF Then
            
                text(1) = GetString(objrs(1))
                text(2) = GetString(objrs(2))
                Category.text = GetString(objrs(3))
                text(4) = GetString(objrs(4))
                text(5) = GetString(objrs(5))
                text(6) = GetString(objrs(6))
                text(3) = objrs(7)
                
            End If
            
            .Close
        
        End With
        
    Screen.MousePointer = 0
    cmdDelete.Enabled = True
End Sub

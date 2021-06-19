VERSION 5.00
Begin VB.Form frmInsertStudents 
   Caption         =   "Insert Students Details"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   Icon            =   "frmInsertStudents.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   6915
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Reset"
      Height          =   300
      Left            =   4500
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6855
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
         Height          =   3375
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   3135
         Begin VB.ComboBox Course 
            Height          =   315
            ItemData        =   "frmInsertStudents.frx":0442
            Left            =   1200
            List            =   "frmInsertStudents.frx":0461
            TabIndex        =   22
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox text 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "d/M/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   1200
            TabIndex        =   3
            Top             =   2400
            Width           =   1695
         End
         Begin VB.TextBox text 
            Height          =   285
            Index           =   1
            Left            =   1200
            TabIndex        =   1
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox text 
            Height          =   285
            Index           =   0
            Left            =   1200
            TabIndex        =   0
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox txt 
            Height          =   285
            Index           =   1
            Left            =   4320
            TabIndex        =   16
            Top             =   480
            Width           =   1695
         End
         Begin VB.ComboBox Category 
            Height          =   315
            ItemData        =   "frmInsertStudents.frx":04C2
            Left            =   1200
            List            =   "frmInsertStudents.frx":04CC
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1920
            Width           =   1695
         End
         Begin VB.Label lbl 
            Caption         =   "LCardNo"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lbl 
            Caption         =   "Batch"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lbl 
            Caption         =   "Course"
            Height          =   255
            Index           =   21
            Left            =   120
            TabIndex        =   19
            Top             =   1440
            Width           =   735
         End
         Begin VB.Label lbl 
            Caption         =   "Category"
            Height          =   255
            Index           =   32
            Left            =   120
            TabIndex        =   18
            Top             =   1920
            Width           =   735
         End
         Begin VB.Label lbl 
            Caption         =   "Date of Issue (dd/mm/yy)"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   17
            Top             =   2400
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   5640
         TabIndex        =   9
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   300
         Left            =   3360
         TabIndex        =   7
         Top             =   2760
         Width           =   1095
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
         TabIndex        =   11
         Top             =   360
         Width           =   3375
         Begin VB.TextBox text 
            Height          =   525
            Index           =   5
            Left            =   1320
            MultiLine       =   -1  'True
            TabIndex        =   6
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox text 
            Height          =   285
            Index           =   4
            Left            =   1320
            TabIndex        =   5
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox text 
            Height          =   285
            Index           =   3
            Left            =   1320
            TabIndex        =   4
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label lbl 
            Caption         =   "RollNo"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   735
         End
         Begin VB.Label lbl 
            Caption         =   "Address (upto 255 characters)"
            Height          =   615
            Index           =   6
            Left            =   120
            TabIndex        =   13
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label lbl 
            Caption         =   "Name"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "frmInsertStudents"
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

Private Sub cmdReset_Click()
    Dim i As Integer
    For i = 0 To text.Count - 1
        text(i) = ""
    Next
    
    text(6).text = ""
    
    text(2) = Day(Date) & "/" & Month(Date) & "/" & Right(CStr(Year(Date)), 2)
    
    
End Sub

Private Sub cmdSave_Click()
    Screen.MousePointer = 11
    
    '=====================
    'Validating user Input
    '=====================
    
    For i = 0 To text.Count - 1
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
    
    If Course.text = "" Then
        Beep
        MsgBox "This field can't be empty, Please select one option", vbCritical, "Error"
        Course.SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    If Len(text(2)) > 8 Or Len(text(2)) < 6 Then
        Beep
        MsgBox "Invalid date entry, Date should be entered in dd/mm/yy Format.", vbCritical, "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    If Len(text(5)) > 255 Then
        Beep
        MsgBox "Address field can contain a maximum of 255 characters.", vbCritical, "Error"
        text(5).SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    With ObjCon
    
        .Open FileDSN
        
        '=================================================
        'Checking whether LCardNo and RollNo exists or not
        '=================================================
        
        query = "Select LCardNo,RollNo from Students where LCardNo='" & Trim(text(0)) & "' or RollNo='" & Trim(text(3)) & "'"
        Set objrs = .Execute(query)
        
        If Not objrs.EOF Then
            Beep
            MsgBox "LCardNo or RollNo already exists, Enter another.", vbCritical, "Error"
            text(0).SetFocus
        Else
            '=======================================
            'Inseting students details into Students
            '=======================================
            
            query = "insert into students values('" & SetString(Trim(text(0))) & "','" & SetString(Trim(text(1))) & "','" & SetString(Trim(Course.text)) & "','" & SetString(Trim(Category.text)) & "','" & SetString(Trim(text(3))) & "','" & SetString(Trim(text(4))) & "','" & SetString(Trim(text(5))) & "','" & Trim(text(2)) & "',0)"
            .Execute (query)
            
            Beep
            MsgBox "Record Inserted", vbInformation, "Congratulations"
            
        End If
        
        .Close
        
    End With
    
    Screen.MousePointer = 0
    
End Sub

Private Sub Form_Load()

Screen.MousePointer = 11
    Course.text = Course.List(0)
    Category.text = Category.List(0)
    
    text(2) = Day(Date) & "/" & Month(Date) & "/" & Right(CStr(Year(Date)), 2)
    
    
   
    Call tot(1)
    
Screen.MousePointer = 0

End Sub

Private Sub Form_Resize()
    Me.Height = 4365
    Me.Width = 7035
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call tot(-1)
End Sub

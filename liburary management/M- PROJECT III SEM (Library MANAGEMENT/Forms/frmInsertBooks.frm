VERSION 5.00
Begin VB.Form frmInsertBooks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Books"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9285
   Icon            =   "frmInsertBooks.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   9285
   Begin VB.Frame Frame1 
      Caption         =   "Insert Book Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   9255
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   300
         Left            =   120
         TabIndex        =   19
         Top             =   5640
         Width           =   1095
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "&Reset"
         Height          =   300
         Left            =   5445
         TabIndex        =   21
         Top             =   5640
         Width           =   1095
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   300
         Left            =   7560
         TabIndex        =   22
         Top             =   5640
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2775
         TabIndex        =   20
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Frame Frame3 
         Height          =   5295
         Left            =   4920
         TabIndex        =   36
         Top             =   240
         Width           =   4215
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   6
            Left            =   2280
            TabIndex        =   15
            Top             =   2760
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   16
            Left            =   2280
            TabIndex        =   18
            Top             =   4200
            Width           =   1695
         End
         Begin VB.TextBox txtFields 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Index           =   15
            Left            =   2280
            TabIndex        =   17
            Top             =   3720
            Width           =   1695
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   5
            Left            =   2280
            TabIndex        =   14
            Top             =   2280
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   13
            Left            =   2280
            TabIndex        =   13
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   12
            Left            =   2280
            TabIndex        =   12
            Top             =   1320
            Width           =   1095
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   11
            Left            =   2280
            TabIndex        =   11
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtFields 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   2280
            TabIndex        =   10
            Top             =   360
            Width           =   1455
         End
         Begin VB.ComboBox txtCombo1 
            Height          =   315
            ItemData        =   "frmInsertBooks.frx":0442
            Left            =   2280
            List            =   "frmInsertBooks.frx":044F
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "BookNo:"
            Height          =   255
            Index           =   19
            Left            =   360
            TabIndex        =   45
            Top             =   2760
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Volume:"
            Height          =   255
            Index           =   17
            Left            =   360
            TabIndex        =   44
            Top             =   840
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Pages:"
            Height          =   255
            Index           =   10
            Left            =   360
            TabIndex        =   43
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Language:"
            Height          =   255
            Index           =   9
            Left            =   360
            TabIndex        =   42
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Date of Entry (dd/mm/yy):"
            Height          =   375
            Index           =   7
            Left            =   360
            TabIndex        =   41
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Classno:"
            Height          =   255
            Index           =   5
            Left            =   360
            TabIndex        =   40
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Category:"
            Height          =   255
            Index           =   4
            Left            =   360
            TabIndex        =   39
            Top             =   3240
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Billno:"
            Height          =   255
            Index           =   3
            Left            =   360
            TabIndex        =   38
            Top             =   4200
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Billdate (dd/mm/yy):"
            Height          =   255
            Index           =   26
            Left            =   360
            TabIndex        =   37
            Top             =   3720
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5295
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   4695
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   2
            Left            =   1800
            TabIndex        =   2
            Top             =   1320
            Width           =   2655
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   1
            Left            =   1800
            TabIndex        =   1
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox txtFields 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   1800
            TabIndex        =   25
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   8
            Left            =   1800
            TabIndex        =   3
            Top             =   1800
            Width           =   2655
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   10
            Left            =   1800
            TabIndex        =   6
            Top             =   3240
            Width           =   2655
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   9
            Left            =   1800
            TabIndex        =   5
            Top             =   2760
            Width           =   2655
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   3
            Left            =   1800
            TabIndex        =   4
            Top             =   2280
            Width           =   2655
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   17
            Left            =   1800
            TabIndex        =   7
            Top             =   3720
            Width           =   2655
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   14
            Left            =   1800
            TabIndex        =   8
            Top             =   4200
            Width           =   2655
         End
         Begin VB.TextBox txtFields 
            Height          =   285
            Index           =   4
            Left            =   1800
            TabIndex        =   9
            Top             =   4680
            Width           =   2655
         End
         Begin VB.Label lblLabels 
            Caption         =   "Title:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   35
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Authors (Separate by Commas)"
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   34
            Top             =   1320
            Width           =   1575
         End
         Begin VB.Label lblLabels 
            Caption         =   "Accession_no:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lblLabels 
            Caption         =   "Place_of_publication:"
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   32
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            Caption         =   "Publisher:"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   31
            Top             =   2280
            Width           =   855
         End
         Begin VB.Label lblLabels 
            Caption         =   "Edition:"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   30
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label lblLabels 
            Caption         =   "Cost:"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   29
            Top             =   3240
            Width           =   855
         End
         Begin VB.Label lblLabels 
            Caption         =   "Remarks:"
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   28
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Source:"
            Height          =   255
            Index           =   14
            Left            =   120
            TabIndex        =   27
            Top             =   4200
            Width           =   855
         End
         Begin VB.Label lblLabels 
            Caption         =   "Subject:"
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   26
            Top             =   4680
            Width           =   1215
         End
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   9285
      TabIndex        =   0
      Top             =   5715
      Width           =   9285
   End
End
Attribute VB_Name = "frmInsertBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
    On Error Resume Next
    
    With ObjCon
    
        Screen.MousePointer = 11
        '============================
        'Getting the Accession Number
        '============================
        .Open FileDSN
        query = "select max(AccessionNo) from BookI"
        Set objrs = .Execute(query)
        
        If IsNull(objrs(0)) Then
            txtFields(0) = Val(1)
        Else
            txtFields(0) = Val(objrs(0)) + Val(1)
        End If
        
        .Close
        
        txtCombo1.text = txtCombo1.List(2)
        txtFields(7) = Day(Date) & "/" & Month(Date) & "/" & Right(CStr(Year(Date)), 2)
        txtFields(15) = Day(Date) & "/" & Month(Date) & "/" & Right(CStr(Year(Date)), 2)
        
        cmdSave.Enabled = True
        Screen.MousePointer = 0
        
    End With
    
End Sub

Private Sub cmdClose_Click()
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub cmdReset_Click()
    
    Dim i As Integer
    For i = 0 To 17
        txtFields(i) = ""
    Next
    
End Sub

Private Sub cmdSave_Click()
    On Error Resume Next
    
    '=====================
    'Validating User Input
    '=====================
    
    If txtFields(0) = "" Then
        Beep
        MsgBox "Enter the title of the book.", vbCritical, "Error"
        Screen.MousePointer = 0
        txtFields(0).SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtFields(12)) Then
        Beep
        MsgBox "Pages can be numeric only.", vbCritical, "Error"
        txtFields(12).SetFocus
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    If Len(txtFields(7)) > 8 Or Len(txtFields(15)) > 8 Or Len(txtFields(7)) < 6 Or Len(txtFields(15)) < 6 Then
        Beep
        MsgBox "Invalid Date, BillDate and Date of Entry should be entered in the dd/mm/yy format.", vbCritical, "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    With ObjCon
        .Open FileDSN
        
        '===============================
        'Inserting data into BookI table
        '===============================
        
        query = "insert into BookI values(" & CInt(Trim(txtFields(0))) & ",'" & SetString(Trim(txtFields(1))) & "','" & SetString(Trim(txtFields(2))) & "','" & SetString(Trim(txtFields(3))) & "','" & Trim(txtCombo1.text) & "','" & SetString(Trim(txtFields(4))) & "','" & SetString(Trim(txtFields(5))) & "','" & SetString(Trim(txtFields(6))) & "')"
        .Execute (query)
        
        '================================
        'Inserting data into BookII table
        '================================
        
        query = "insert into BookII values(" & CInt(Trim(txtFields(0))) & ",'" & CDate(Trim(txtFields(7))) & "','" & SetString(Trim(txtFields(8))) & "','" & SetString(Trim(txtFields(9))) & "','" & SetString(Trim(txtFields(10).text)) & "','" & SetString(Trim(txtFields(11))) & "'," & CInt(Trim(txtFields(12))) & ",'" & SetString(Trim(txtFields(13))) & "','" & SetString(Trim(txtFields(14))) & "','" & CDate(Trim(txtFields(15))) & "','" & SetString(Trim(txtFields(16))) & "','" & SetString(Trim(txtFields(17))) & "')"
        .Execute (query)
        
        
        '===============================
        'Inserting data into Issue table
        '===============================
        Dim i As Integer
        Dim j As Integer
        
        i = 1
        j = 0
        query = "insert into Issue values(" & CInt(Trim(txtFields(0))) & ",'Yes','No','No','No')"
        .Execute (query)
        
        .Close
        
        Beep
        MsgBox "Record Inserted", vbInformation, "Congratulations!"
        Screen.MousePointer = 0
        cmdSave.Enabled = False
        Call cmdReset_Click
        Exit Sub
        
    End With
    
End Sub

Private Sub Form_Load()
    txtCombo1.text = txtCombo1.List(2)
    Call tot(1)
End Sub

Private Sub Form_Resize()
    Me.Height = 6510
    Me.Width = 9375
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call tot(-1)
End Sub

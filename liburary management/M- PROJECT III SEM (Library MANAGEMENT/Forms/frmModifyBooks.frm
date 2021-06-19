VERSION 5.00
Begin VB.Form frmModifyBooks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modify Book Details"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10455
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmModifyBooks.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   10455
   Begin VB.Frame Frame1 
      Caption         =   "Modify Book Details"
      Height          =   6015
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10335
      Begin VB.Frame Frame5 
         Caption         =   "Select the AccessionNo and Click"
         Height          =   1695
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   5055
         Begin VB.ListBox List2 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1230
            Left            =   2160
            TabIndex        =   40
            Top             =   360
            Width           =   2775
         End
         Begin VB.ListBox List1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1230
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   5280
         TabIndex        =   33
         Top             =   4920
         Width           =   4935
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Update"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   360
            TabIndex        =   36
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "&Refresh"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2040
            TabIndex        =   35
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "&Exit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3600
            TabIndex        =   34
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Left            =   5280
         TabIndex        =   18
         Top             =   240
         Width           =   4935
         Begin VB.TextBox txtFields 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   7
            Left            =   2160
            TabIndex        =   29
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   2160
            TabIndex        =   28
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox txtFields 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   8
            Left            =   2160
            TabIndex        =   27
            Top             =   600
            Width           =   2415
         End
         Begin VB.TextBox txtFields 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   11
            Left            =   2160
            TabIndex        =   26
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox txtFields 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   10
            Left            =   2160
            TabIndex        =   25
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtFields 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   12
            Left            =   2160
            TabIndex        =   24
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox txtFields 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   13
            Left            =   2160
            TabIndex        =   23
            Top             =   2400
            Width           =   2055
         End
         Begin VB.TextBox txtFields 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   15
            Left            =   2160
            TabIndex        =   22
            Top             =   3120
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   14
            Left            =   2160
            TabIndex        =   21
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox txtFields 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "d MMM yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   17
            Left            =   2160
            TabIndex        =   20
            Top             =   3840
            Width           =   1455
         End
         Begin VB.TextBox txtFields 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   16
            Left            =   2160
            TabIndex        =   19
            Top             =   3480
            Width           =   2415
         End
         Begin VB.Label lblLabels 
            Caption         =   "Source:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   14
            Left            =   240
            TabIndex        =   48
            Top             =   2760
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Language:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   47
            Top             =   2400
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Pages:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   46
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Caption         =   "Volume:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   17
            Left            =   240
            TabIndex        =   45
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            Caption         =   "Cost:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   240
            TabIndex        =   44
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label lblLabels 
            Caption         =   "Place_of_publication:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   43
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label lblLabels 
            Caption         =   "Date of Entry:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   240
            TabIndex        =   42
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Edition:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   240
            TabIndex        =   41
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label lblLabels 
            Caption         =   "Billno:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   32
            Top             =   3480
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "Billdate:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   31
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Remarks:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   13
            Left            =   240
            TabIndex        =   30
            Top             =   3840
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3975
         Left            =   120
         TabIndex        =   2
         Top             =   1920
         Width           =   5055
         Begin VB.TextBox txtFields 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   1920
            TabIndex        =   37
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox txtFields 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   4
            Left            =   1920
            TabIndex        =   9
            Top             =   2160
            Width           =   2295
         End
         Begin VB.TextBox txtFields 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   1920
            TabIndex        =   8
            Top             =   1440
            Width           =   2415
         End
         Begin VB.TextBox txtFields 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   1920
            TabIndex        =   7
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtFields 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1920
            TabIndex        =   6
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox txtFields 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   1920
            TabIndex        =   5
            Top             =   2880
            Width           =   1695
         End
         Begin VB.TextBox txtFields 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   5
            Left            =   1920
            TabIndex        =   4
            Top             =   2520
            Width           =   1335
         End
         Begin VB.ComboBox txtCombo1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmModifyBooks.frx":0442
            Left            =   1920
            List            =   "frmModifyBooks.frx":044F
            TabIndex        =   3
            Top             =   1800
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Subject:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   17
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "Publisher:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   16
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "Accession_No:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Title:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   240
            TabIndex        =   14
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "Author:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   13
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label lblLabels 
            Caption         =   "BookNo :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   18
            Left            =   240
            TabIndex        =   12
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label lblLabels 
            Caption         =   "Classno:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   11
            Top             =   2520
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Category:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   10
            Top             =   1800
            Width           =   1335
         End
      End
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   10455
      TabIndex        =   0
      Top             =   5700
      Width           =   10455
   End
End
Attribute VB_Name = "frmModifyBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdDelete_Click()
    Screen.MousePointer = 11
    With ObjCon
        .Open FileDSN
        query = "select AccessionNo from Issue where AccessionNo=" & CInt(Trim(txtFields(0))) & " and Issue =" & CBool("true") & " or Discarded=" & CBool("true") & " or Missing=" & CBool("true")
        Set objrs = .Execute(query)
        
        If Not objrs.EOF Then
            Beep
            MsgBox "This Book is Issued, or put in discarded, or missing list,so can't be deleted. Please check the status of the book, GOTO Search >> Book >> Status", vbExclamation, "Warning"
        Else
            
            Screen.MousePointer = 0
            
            '===============================
            'Deleting entry from Issue Table
            '===============================
            query = "delete from Issue where AccessionNo=" & CInt(Trim(txtFields(0)))
            .Execute (query)
            
            '================================
            'Deleting entry from BookII Table
            '================================
            query = "delete from BookII where AccessionNo=" & CInt(Trim(txtFields(0)))
            .Execute (query)
            
            '===============================
            'Deleting entry from BookI Table
            '===============================
            query = "delete from BookI where AccessionNo=" & CInt(Trim(txtFields(0)))
            .Execute (query)
            
            Beep
            MsgBox "Successfully deleted", vbInformation, "Info"
            
            List2.RemoveItem (List2.ListIndex)
            List1.RemoveItem (List1.ListIndex)
            
        End If
        .Close
    End With
    cmdDelete.Enabled = False
    
    Dim i As Integer
    For i = 0 To 17
        txtFields(i) = ""
    Next
    
    Screen.MousePointer = 0
End Sub

Private Sub cmdExit_Click()
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    Screen.MousePointer = 11
    Call Form_Load
    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
    On Error Resume Next
    
    Screen.MousePointer = 11
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
    
    If Len(txtFields(7)) > 8 Or Len(txtFields(15)) > 8 Then
        Beep
        MsgBox "Invalid Date, BillDate and Date of Entry should be entered in the dd/mm/yy format.", vbCritical, "Error"
        Screen.MousePointer = 0
        Exit Sub
    End If
    
    With ObjCon
        .Open FileDSN
        
        '===============================
        'Updating data into BookI table
        '===============================
          
        query = "Update BookI set Title ='" & SetString(Trim(txtFields(1))) & "',Author='" & SetString(Trim(txtFields(2))) & "',Publisher='" & SetString(Trim(txtFields(3))) & "',Category='" & Trim(txtCombo1.text) & "',Subject='" & SetString(Trim(txtFields(4))) & "',ClassNo='" & SetString(Trim(txtFields(5))) & "',BookNo='" & SetString(Trim(txtFields(6))) & "' where AccessionNo=" & CInt(Trim(txtFields(0)))
        .Execute (query)
        
        '================================
        'Updating data into BookII table
        '================================
        
        query = "Update BookII set DOE='" & CDate(Trim(txtFields(7))) & "',POP='" & SetString(Trim(txtFields(8))) & "',Edition='" & SetString(Trim(txtFields(9))) & "',Cost='" & SetString(Trim(txtFields(10).text)) & "',Volume='" & SetString(Trim(txtFields(11))) & "',Pages=" & CInt(Trim(txtFields(12))) & ",Language='" & SetString(Trim(txtFields(13))) & "',Source='" & SetString(Trim(txtFields(14))) & "',BillDate='" & CDate(Trim(txtFields(15))) & "',BillNo='" & SetString(Trim(txtFields(16))) & "',Remarks='" & SetString(Trim(txtFields(17))) & "' where AccessionNo=" & CInt(Trim(txtFields(0)))
        .Execute (query)
        
        .Close
    End With
    
    Beep
    MsgBox "Record Updated", vbInformation, "Info"
    
    cmdSave.Enabled = False
    Screen.MousePointer = 0
    
End Sub

Private Sub Form_Resize()
    Me.Height = 6495
    Me.Width = 10545
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Screen.MousePointer = 0
    
    '==========================================
    'Retrieving the Accession Number and Titles
    '==========================================
    With ObjCon
        .Open FileDSN
        query = "select distinct(AccessionNo),Title from BookI order by AccessionNo"
        Set objrs = .Execute(query)
        
    '================================
    'Inputting records into the list1
    '================================
    
    If Not objrs.EOF Then
        List1.Clear
        List2.Clear
        While Not objrs.EOF
            List1.AddItem objrs(0)
            List2.AddItem GetString(objrs(1))
            objrs.MoveNext
        Wend
    Else
        List1.Enabled = False
    End If
    
        .Close
    End With
    
    Call tot(1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call tot(-1)
End Sub

Private Sub List1_Click()
    On Error Resume Next
    List2.text = List2.List(List1.ListIndex)
    
    With ObjCon
        
        .Open FileDSN
        
        '==================================================
        'Loading the records for the particular AccessionNo
        '==================================================
        
        query = "select * from BookI where AccessionNo=" & CInt(List1.List(List1.ListIndex))
        Set objrs = .Execute(query)
        
        Dim i As Integer
        For i = 0 To 3
            txtFields(i) = GetString(objrs(i))
        Next
        
        txtCombo1.text = GetString(objrs(4))
        
        For i = 4 To 6
            txtFields(i) = GetString(objrs(i + 1))
        Next
        
        query = "select * from BookII where AccessionNo=" & CInt(List1.List(List1.ListIndex))
        Set objrs = .Execute(query)
        
        For i = 7 To 17
            txtFields(i) = GetString(objrs(i - 6))
        Next
        
        .Close
    End With
    
    cmdSave.Enabled = True
    cmdDelete.Enabled = True
    Screen.MousePointer = 0
    
End Sub


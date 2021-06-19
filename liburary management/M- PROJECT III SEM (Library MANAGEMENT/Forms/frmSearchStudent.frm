VERSION 5.00
Begin VB.Form frmSearchStudents 
   Caption         =   "Search Students :- Name"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7395
   Icon            =   "frmSearchStudent.frx":0000
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
         Caption         =   "LCardNos for the Name Selected"
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
         Left            =   3000
         TabIndex        =   5
         Top             =   960
         Width           =   4215
         Begin VB.CommandButton cmdExit 
            Caption         =   "&Close"
            Height          =   300
            Left            =   3000
            TabIndex        =   10
            Top             =   1560
            Width           =   1095
         End
         Begin VB.ListBox List2 
            Height          =   1230
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   3975
         End
         Begin VB.Label TotREc1 
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1560
            Width           =   3015
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Select an alphabet for Search and press Search"
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
         TabIndex        =   2
         Top             =   120
         Width           =   7095
         Begin VB.ComboBox Albhabets 
            Height          =   315
            ItemData        =   "frmSearchStudent.frx":0442
            Left            =   240
            List            =   "frmSearchStudent.frx":0444
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton cmdSearch 
            Caption         =   "&Search"
            Height          =   375
            Left            =   2160
            TabIndex        =   3
            Top             =   360
            Width           =   2055
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
         Height          =   1935
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   2775
         Begin VB.ListBox List1 
            Height          =   1230
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label TotRec 
            Caption         =   "Label1"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1560
            Width           =   2535
         End
      End
   End
End
Attribute VB_Name = "frmSearchStudents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    
End Sub

Private Sub Albhabets_Change()
    Call cmdSearch_Click
End Sub
Private Sub Albhabets_Click()
    Call cmdSearch_Click
End Sub

Private Sub cmdExit_Click()
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
    Screen.MousePointer = 11
    
    '=====================
    'Searching for records
    '=====================
    With ObjCon
        .Open FileDSN
        
            query = "select Name from Students where Name like '" & Trim(Albhabets.text) & "%' order by Name"
            Set objrs = .Execute(query)
            
            List1.Clear
            List1.Enabled = True
            If Not objrs.EOF Then
                While Not objrs.EOF
                    List1.AddItem objrs(0)
                    objrs.MoveNext
                Wend
                 List1.text = List1.List(0)
                TotRec.Caption = "Total Records Found : " & List1.ListCount
            Else
                List1.AddItem "Empty"
                List1.Enabled = False
                TotRec.Caption = "Total Records Found : " & List1.ListCount - 1
            End If
            
        .Close
    End With
    Screen.MousePointer = 0
End Sub

Private Sub Form_Initialize()
    Screen.MousePointer = 11
        
        
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
On Error Resume Next
    Screen.MousePointer = 11
    Dim i As Integer
    For i = 65 To 90
        Albhabets.AddItem Chr(i)
    Next
        
    Albhabets.text = Albhabets.List(0)
    
    Call tot(1)
    Screen.MousePointer = 0
End Sub

Private Sub Form_Resize()
    Me.Height = 3510
    Me.Width = 7515
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = 0
    Call tot(-1)
End Sub

Private Sub List1_Click()
On Error Resume Next
     Screen.MousePointer = 11
    
    '============================================
    'Searching for LCardNos for the selected Name
    '============================================
    
    With ObjCon
        .Open FileDSN
        
            query = "select LCardNo from Students where Name='" & Trim(GetString(List1.List(List1.ListIndex))) & "' order by LCardNo"
            Set objrs = .Execute(query)
            List2.Clear
            If Not objrs.EOF Then
                
                While Not objrs.EOF
                    List2.AddItem objrs(0)
                    objrs.MoveNext
                Wend
               List2.text = List2.List(List1.ListIndex)
                TotREc1 = "Total Records Found : " & List2.ListCount
            Else
                List2.AddItem "Empty"
                List2.Enabled = False
                TotREc1 = "Total Records Found : 0"
            End If
            
        .Close
        
        Screen.MousePointer = 0
    End With
End Sub

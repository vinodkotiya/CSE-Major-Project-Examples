VERSION 5.00
Begin VB.Form FrmFacltDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Faculty Entry Module"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   Icon            =   "FrmFatDetails.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   8730
   Begin VB.Frame Frame1 
      Caption         =   "Faculty Details"
      Height          =   6855
      Left            =   15
      TabIndex        =   24
      Top             =   -15
      Width           =   8700
      Begin VB.Frame Frame3 
         Caption         =   "Movement"
         Height          =   1875
         Left            =   6540
         TabIndex        =   39
         Top             =   4905
         Width           =   2085
         Begin VB.CommandButton CmdMovement 
            Height          =   360
            Index           =   3
            Left            =   1230
            Picture         =   "FrmFatDetails.frx":0442
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Go Next"
            Top             =   720
            Width           =   495
         End
         Begin VB.CommandButton CmdMovement 
            Height          =   465
            Index           =   2
            Left            =   840
            Picture         =   "FrmFatDetails.frx":0884
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Go Last"
            Top             =   1050
            Width           =   405
         End
         Begin VB.CommandButton CmdMovement 
            Height          =   360
            Index           =   1
            Left            =   345
            Picture         =   "FrmFatDetails.frx":0CC6
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Go Previous"
            Top             =   720
            Width           =   495
         End
         Begin VB.CommandButton CmdMovement 
            DisabledPicture =   "FrmFatDetails.frx":1108
            DownPicture     =   "FrmFatDetails.frx":154A
            Height          =   480
            Index           =   0
            Left            =   825
            Picture         =   "FrmFatDetails.frx":198C
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Go First"
            Top             =   270
            Width           =   405
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Operations"
         Height          =   3495
         Left            =   6525
         TabIndex        =   38
         Top             =   1335
         Width           =   2070
         Begin VB.CommandButton CmdOperation 
            Caption         =   "&Ok"
            Height          =   345
            Index           =   6
            Left            =   105
            TabIndex        =   18
            ToolTipText     =   "Save Record"
            Top             =   3030
            Width           =   915
         End
         Begin VB.CommandButton CmdOperation 
            Caption         =   "&Close"
            Height          =   345
            Index           =   7
            Left            =   1095
            TabIndex        =   19
            ToolTipText     =   "Close Window"
            Top             =   3030
            Width           =   915
         End
         Begin VB.CommandButton CmdOperation 
            Caption         =   "&Add"
            Height          =   345
            Index           =   0
            Left            =   105
            TabIndex        =   12
            ToolTipText     =   "Add Record"
            Top             =   210
            Width           =   915
         End
         Begin VB.CommandButton CmdOperation 
            Caption         =   "D&elete"
            Height          =   345
            Index           =   1
            Left            =   1095
            TabIndex        =   13
            ToolTipText     =   "Delete Record"
            Top             =   210
            Width           =   915
         End
         Begin VB.CommandButton CmdOperation 
            Caption         =   "&Modify"
            Height          =   345
            Index           =   2
            Left            =   105
            TabIndex        =   14
            ToolTipText     =   "Modify Record"
            Top             =   1095
            Width           =   915
         End
         Begin VB.CommandButton CmdOperation 
            Caption         =   "&Find"
            Height          =   345
            Index           =   3
            Left            =   1095
            TabIndex        =   15
            ToolTipText     =   "Find Record"
            Top             =   1095
            Width           =   915
         End
         Begin VB.CommandButton CmdOperation 
            Caption         =   "&Display"
            Height          =   345
            Index           =   4
            Left            =   105
            TabIndex        =   16
            ToolTipText     =   "Display Record"
            Top             =   2115
            Width           =   915
         End
         Begin VB.CommandButton CmdOperation 
            Caption         =   "Ca&ncel"
            Height          =   345
            Index           =   5
            Left            =   1110
            TabIndex        =   17
            ToolTipText     =   "Cancel Operation"
            Top             =   2115
            Width           =   915
         End
      End
      Begin VB.TextBox TxtFDetails 
         Height          =   660
         Index           =   10
         Left            =   1800
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   5355
         Width           =   4500
      End
      Begin VB.TextBox TxtFDetails 
         Height          =   705
         Index           =   11
         Left            =   1800
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   6075
         Width           =   4500
      End
      Begin VB.TextBox TxtFDetails 
         Height          =   660
         Index           =   7
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   7
         Top             =   3945
         Width           =   4500
      End
      Begin VB.TextBox TxtFDetails 
         Height          =   645
         Index           =   6
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   6
         Top             =   3255
         Width           =   4515
      End
      Begin VB.TextBox TxtFDetails 
         Height          =   300
         Index           =   3
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   3
         Top             =   2145
         Width           =   4530
      End
      Begin VB.TextBox TxtFDetails 
         Height          =   300
         Index           =   2
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   2
         Top             =   1785
         Width           =   4530
      End
      Begin VB.TextBox TxtFDetails 
         Height          =   300
         Index           =   1
         Left            =   1800
         MaxLength       =   40
         TabIndex        =   1
         Top             =   1410
         Width           =   4530
      End
      Begin VB.ComboBox CmbSex 
         Height          =   315
         Index           =   4
         ItemData        =   "FrmFatDetails.frx":1DCE
         Left            =   1800
         List            =   "FrmFatDetails.frx":1DD8
         TabIndex        =   4
         Text            =   "Male"
         Top             =   2535
         Width           =   1875
      End
      Begin VB.TextBox TxtFDetails 
         Height          =   300
         Index           =   9
         Left            =   1800
         MaxLength       =   11
         TabIndex        =   9
         Top             =   5025
         Width           =   3225
      End
      Begin VB.TextBox TxtFDetails 
         Height          =   300
         Index           =   8
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   8
         Top             =   4680
         Width           =   3225
      End
      Begin VB.TextBox TxtFDetails 
         Height          =   300
         Index           =   5
         Left            =   1800
         MaxLength       =   11
         TabIndex        =   5
         Top             =   2895
         Width           =   1860
      End
      Begin VB.TextBox TxtFDetails 
         Height          =   300
         Index           =   0
         Left            =   1800
         MaxLength       =   6
         TabIndex        =   0
         Top             =   1065
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Work Exp"
         Height          =   195
         Index           =   11
         Left            =   900
         TabIndex        =   37
         Top             =   6540
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Education"
         Height          =   195
         Index           =   10
         Left            =   885
         TabIndex        =   36
         Top             =   5805
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date Of Joining"
         Height          =   195
         Index           =   9
         Left            =   510
         TabIndex        =   35
         Top             =   5115
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Designation"
         Height          =   195
         Index           =   8
         Left            =   765
         TabIndex        =   34
         Top             =   4770
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Permanent Add"
         Height          =   195
         Index           =   7
         Left            =   510
         TabIndex        =   33
         Top             =   4395
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Local Add"
         Height          =   195
         Index           =   6
         Left            =   885
         TabIndex        =   32
         Top             =   3690
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Date Of Birth"
         Height          =   195
         Index           =   5
         Left            =   690
         TabIndex        =   31
         Top             =   3015
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Index           =   4
         Left            =   1335
         TabIndex        =   30
         Top             =   2640
         Width           =   270
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Father's Name"
         Height          =   195
         Index           =   3
         Left            =   585
         TabIndex        =   29
         Top             =   2235
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail ID"
         Height          =   195
         Index           =   2
         Left            =   960
         TabIndex        =   28
         Top             =   1875
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Faculty Name"
         Height          =   195
         Index           =   1
         Left            =   630
         TabIndex        =   27
         Top             =   1500
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Faculty Code"
         Height          =   195
         Index           =   0
         Left            =   675
         TabIndex        =   26
         Top             =   1140
         Width           =   930
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Faculty Details Module"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   360
         Left            =   2760
         TabIndex        =   25
         Top             =   240
         Width           =   3180
      End
   End
End
Attribute VB_Name = "FrmFacltDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Procedure For hide and show buttons

Private Sub HideShowButtons(Index As Integer)
    Dim i As Integer
    Select Case Index
        Case 0, 1, 2, 3, 4:  'Hide the buttons
        
            For i = 0 To 4
                CmdOperation(i).Enabled = False
            Next i
        Exit Sub
            
        Case 5, 6, 7 ' Show the buttons
            For i = 0 To 4
                CmdOperation(i).Enabled = True
            Next i
        Exit Sub
    End Select
End Sub

Private Sub CmdMovement_Click(Index As Integer)

End Sub

Private Sub CmdOperation_Click(Index As Integer)
    Select Case Index
        Case 0: 'Add
                HideShowButtons (Index)
        Case 1: 'Delete
                HideShowButtons (Index)
        Case 2: 'Modify
                HideShowButtons (Index)
        Case 3: 'Search
                HideShowButtons (Index)
        Case 4: 'Display
                HideShowButtons (Index)
        Case 5: 'Cancel
                HideShowButtons (Index)
        Case 6: 'Commit
                HideShowButtons (Index)
        Case 7: 'Close The Windows
                HideShowButtons (Index)
                Unload FrmFatDetails
    End Select
End Sub

Private Sub Form_Load()
    FrmFatDetails.Height = 7230
    FrmFatDetails.Width = 8820
    FrmFatDetails.Top = (CFrmMain.ScaleHeight - FrmFatDetails.ScaleHeight) / 2 - 150
    FrmFatDetails.Left = (CFrmMain.ScaleWidth - FrmFatDetails.ScaleWidth) / 2
End Sub

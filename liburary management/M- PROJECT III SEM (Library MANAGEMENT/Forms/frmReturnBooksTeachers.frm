VERSION 5.00
Begin VB.Form frmReturnBooksTeachers 
   Caption         =   "Return Books"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8790
   Icon            =   "frmReturnBooksTeachers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   8790
   Begin VB.Frame Frame1 
      Caption         =   "Return Books Teachers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      Begin VB.Frame Frame8 
         Height          =   1695
         Left            =   120
         TabIndex        =   16
         Top             =   2880
         Width           =   8535
         Begin VB.Frame Frame9 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   2040
            TabIndex        =   17
            Top             =   360
            Width           =   4215
            Begin VB.CommandButton cmdReturn 
               Caption         =   "&Return"
               Height          =   375
               Left            =   120
               TabIndex        =   27
               Top             =   240
               Width           =   3975
            End
            Begin VB.CommandButton cmdExit 
               Caption         =   "&Exit"
               Height          =   300
               Left            =   3000
               TabIndex        =   21
               Top             =   720
               Width           =   1095
            End
            Begin VB.CommandButton cmd 
               Caption         =   "&Exit"
               Height          =   300
               Left            =   6000
               TabIndex        =   20
               Top             =   840
               Width           =   1095
            End
            Begin VB.CommandButton cmdCancel 
               Caption         =   "&Cancel"
               Height          =   300
               Left            =   1560
               TabIndex        =   19
               Top             =   720
               Width           =   1095
            End
            Begin VB.CommandButton cmdIssue 
               Caption         =   "&ReIssue"
               Height          =   300
               Left            =   120
               TabIndex        =   18
               Top             =   720
               Width           =   1095
            End
         End
      End
      Begin VB.Frame Frame5 
         Height          =   2655
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4095
         Begin VB.Frame Frame6 
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
            Height          =   1335
            Left            =   120
            TabIndex        =   6
            Top             =   1200
            Width           =   3855
            Begin VB.Label lblBook 
               Caption         =   "Subject :-"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   11
               Top             =   960
               Width           =   3615
            End
            Begin VB.Label lblBook 
               Caption         =   "Author :-"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   10
               Top             =   720
               Width           =   3615
            End
            Begin VB.Label lblBook 
               Caption         =   "Title :-"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   9
               Top             =   480
               Width           =   3615
            End
            Begin VB.Label lblBook 
               Caption         =   "Accession Number :-"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   8
               Top             =   240
               Width           =   3615
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Select Book Accession Number"
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
            TabIndex        =   3
            Top             =   240
            Width           =   3855
            Begin VB.ComboBox AccessionNo 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   360
               Width           =   2175
            End
            Begin VB.Label lbl 
               Caption         =   "AccessionNo"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   4
               Top             =   360
               Width           =   1095
            End
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2655
         Left            =   4320
         TabIndex        =   1
         Top             =   240
         Width           =   4335
         Begin VB.Frame Frame10 
            Caption         =   "Issue Details"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   3975
            Begin VB.Label lblLabels 
               Caption         =   "Issue Date:"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   26
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label lblLabels 
               Caption         =   "Issue No:"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   25
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "Issue No"
               Height          =   255
               Index           =   0
               Left            =   2040
               TabIndex        =   24
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label1 
               Caption         =   "Issue Date"
               Height          =   255
               Index           =   1
               Left            =   2040
               TabIndex        =   23
               Top             =   480
               Width           =   1455
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Teacher Details"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   120
            TabIndex        =   7
            Top             =   1200
            Width           =   3975
            Begin VB.Label lblBook 
               Caption         =   "Status :-"
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   15
               Top             =   960
               Width           =   3735
            End
            Begin VB.Label lblBook 
               Caption         =   "Address :-"
               Height          =   255
               Index           =   6
               Left            =   120
               TabIndex        =   14
               Top             =   720
               Width           =   3615
            End
            Begin VB.Label lblBook 
               Caption         =   "Phone :-"
               Height          =   255
               Index           =   5
               Left            =   120
               TabIndex        =   13
               Top             =   480
               Width           =   3615
            End
            Begin VB.Label lblBook 
               Caption         =   "Name :-"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   3615
            End
         End
      End
   End
End
Attribute VB_Name = "frmReturnBooksTeachers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame11_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Form_Resize()
    Me.Height = 5175
    Me.Width = 8910
End Sub

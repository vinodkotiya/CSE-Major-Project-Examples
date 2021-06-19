VERSION 5.00
Begin VB.Form frmEditDiscarded 
   Caption         =   "Discarded Books"
   ClientHeight    =   3345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   Icon            =   "frmEditDiscarded.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   4725
   Begin VB.Frame Frame1 
      Caption         =   "Insert Discarded Books"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.Frame Frame3 
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
         Height          =   1815
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4455
         Begin VB.Frame Frame4 
            Height          =   975
            Left            =   120
            TabIndex        =   5
            Top             =   720
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
               TabIndex        =   7
               Top             =   360
               Width           =   3015
            End
            Begin VB.Label Label2 
               Caption         =   "Title :-"
               Height          =   255
               Left            =   120
               TabIndex        =   6
               Top             =   360
               Width           =   735
            End
         End
         Begin VB.ComboBox AccessionNo 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Accession No."
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   2040
         Width           =   4455
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            Height          =   300
            Left            =   600
            TabIndex        =   10
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton cmdExit 
            Caption         =   "&Exit"
            Height          =   300
            Left            =   2640
            TabIndex        =   9
            Top             =   720
            Width           =   1095
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete from Discarded List"
            Height          =   375
            Left            =   600
            TabIndex        =   8
            Top             =   240
            Width           =   3135
         End
      End
   End
End
Attribute VB_Name = "frmEditDiscarded"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmBookDialog 
   Caption         =   "Books Report"
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3870
   Icon            =   "frmBookDialog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1860
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Select an Option"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin Crystal.CrystalReport CR1 
         Left            =   240
         Top             =   1320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Close"
         Height          =   300
         Left            =   2640
         TabIndex        =   4
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdReport 
         Caption         =   "&Report"
         Height          =   300
         Left            =   1320
         TabIndex        =   3
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Other Book Details"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   3615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "All the Books"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmBookDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRefresh_Click()

End Sub

Private Sub cmdExit_Click()
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub cmdReport_Click()

On Error Resume Next

If Option1.value = True Then
    With CR1
        .ReportTitle = "Book Details as on :"
        .ReportFileName = App.Path & "\Reports\TotBookDetails.rpt"
        .DiscardSavedData = True
        .Action = True
    End With
End If

If Option2.value = True Then
      With CR1
        .ReportTitle = "Book Details as on :"
        .ReportFileName = App.Path & "\Reports\OtherBookDetails.rpt"
        .DiscardSavedData = True
        .Action = True
    End With
End If

End Sub


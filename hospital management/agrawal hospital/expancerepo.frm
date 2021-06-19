VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form expancerepo 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Expence Report"
   ClientHeight    =   4620
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5655
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFC0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4620
   ScaleWidth      =   5655
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00FFC0C0&
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   3360
      Width           =   5175
      Begin Crystal.CrystalReport Crpt 
         Left            =   4200
         Top             =   240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         PrintFileLinesPerPage=   60
      End
      Begin VB.CommandButton cmdback 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdok 
         BackColor       =   &H00FFC0C0&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Report"
      Height          =   1095
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   5175
      Begin VB.OptionButton obprint 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Print"
         Height          =   855
         Left            =   3240
         TabIndex        =   8
         Top             =   120
         Width           =   1575
      End
      Begin VB.OptionButton obpreview 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Peview"
         Height          =   615
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Date"
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   5175
      Begin VB.TextBox dtpend 
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox dtpstart 
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Labto 
         BackColor       =   &H00FFC0C0&
         Caption         =   "To"
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Labfrom 
         BackColor       =   &H00FFC0C0&
         Caption         =   "From"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Label Labheading 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Expence  Report"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "expancerepo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public z As Boolean

Private Sub Cmdback_Click()
Unload Me
main.Show

End Sub

Private Sub cmdok_Click()

If dtpstart <= dtpend Then
Dim str As String
     Crpt.Reset
     Crpt.ReportFileName = App.Path & "\EXPENCE.rpt"
     If obpreview = True Then
        Crpt.Destination = crptToWindow
     Else
        Crpt.Destination = crptToPrinter
        On Error GoTo errlable
errlable:               MsgBox "THERE IS NO DEFAULT PRINTER AVAILABLE ", vbCritical, "PRINTER ERROR"

       Exit Sub
     End If
       str = "{expense.paiddate} in date(" & Format(dtpstart, "yyyy,mm,dd") & " ) to date (" & Format(dtpend, "yyyy,mm,dd") & ")"
       Crpt.ReplaceSelectionFormula str
       Crpt.WindowState = crptMaximized
       Crpt.Action = 1
       

       Else
        MsgBox "Enter the correct date"
       Exit Sub
       End If







q = True
End Sub


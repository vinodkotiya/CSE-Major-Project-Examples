VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form formdate 
   BackColor       =   &H00C0E0FF&
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport crReport5 
      Left            =   360
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\bankproject\montrep.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      Connect         =   "DSN=bank"
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Report"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker dtto 
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MM/dd/yy"
      Format          =   24510467
      CurrentDate     =   37381
   End
   Begin MSComCtl2.DTPicker dtfrom 
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MM/dd/yy"
      Format          =   24510467
      CurrentDate     =   37381
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the period to display the report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   3795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To Date : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From Date :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "formdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    If dtto.Value < dtfrom.Value Then
       MsgBox "second date should be greater than first date.", vbCritical + vbOKOnly
       Exit Sub
    End If
      'DataReport_booking.Show
crReport5.Action = crRunReport
crReport5.RetrieveSQLQuery


Unload Me
End Sub

Private Sub Form_Load()
dtfrom.Value = DateAdd("d", -30, Date)
'MsgBox "From Date" & dtfrom.Value
dtto.Value = Date
'MsgBox "To date " & dtto.Value
End Sub

Private Sub Text1_change()
 Command1.Enabled = True
End Sub

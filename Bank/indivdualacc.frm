VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form frmIndivdualAcc 
   Caption         =   "Indivdual Account "
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   4740
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport crReport2 
      Left            =   3960
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\bankproject\indacc.rpt"
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtname 
      DataMember      =   "Command3"
      DataSource      =   "DataEnvironment1"
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdcancle 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdokindivdual 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox txtaccno 
      DataMember      =   "Command3"
      DataSource      =   "DataEnvironment1"
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblname 
      Caption         =   "Name"
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblaccountno 
      Caption         =   "Account No."
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frmIndivdualAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim Rs As New ADODB.Recordset

Private Sub cmdcancle_Click()
Unload Me
End Sub

Private Sub cmdokindivdual_Click()
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\bankproject\bank.mdb;Persist Security Info=False"
Set Rs.ActiveConnection = db
SQL = "Select * from initial where acc_no = '" & txtaccno.Text & "' and name = '" & txtname.Text & "'"
'MsgBox SQL
Rs.Open SQL

If Rs.EOF Or Rs.BOF Then
    MsgBox "Invalid Account No or Name.....", vbCritical
    txtaccno.SetFocus
Else
    'MsgBox "Success", vbInformation
    DataReport1.Show
    Unload Me
End If
Rs.Close
db.Close
Set db = Nothing
Set Rs = Nothing

End Sub

Private Sub Form_Load()
'frmIndivdualAcc.txtaccno.Text = "D0005"
End Sub


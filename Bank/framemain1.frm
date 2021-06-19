VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.1#0"; "CRYSTL32.OCX"
Begin VB.Form frmmain1 
   Caption         =   "Main Form"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport crReport4 
      Left            =   4320
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\bankproject\indacc.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      SelectionFormula=   "{tran.acc_no} = {?accno}"
      Connect         =   "DSN=bank"
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport crReport1 
      Left            =   720
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\bankproject\listacc.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileName   =   "list"
      PrintFileType   =   2
      Connect         =   "DSN=bank"
      SQLQuery        =   "select * from tran"
      PrinterStartPage=   1
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin Crystal.CrystalReport crReport3 
      Left            =   720
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      ReportFileName  =   "C:\bankproject\montreptrial.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      Connect         =   "DSN=bank"
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1200
      TabIndex        =   6
      Top             =   4440
      Width           =   3135
   End
   Begin VB.CommandButton CommandAccEdit 
      Caption         =   "Edit Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1200
      TabIndex        =   5
      Top             =   3720
      Width           =   3135
   End
   Begin VB.CommandButton CommandReportMon 
      Caption         =   " Monthly Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1200
      TabIndex        =   4
      Top             =   3000
      Width           =   3135
   End
   Begin VB.CommandButton CommandTranDaily 
      Caption         =   " Daily Transaction"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1200
      TabIndex        =   3
      Top             =   2280
      Width           =   3135
   End
   Begin VB.CommandButton CommandAccountInd 
      Caption         =   "Individual Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      Top             =   1560
      Width           =   3135
   End
   Begin VB.CommandButton CommandAccountList 
      Caption         =   " List Of Accounts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.CommandButton CommandOpen 
      Caption         =   "Open New Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmmain1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_temp As ADODB.Recordset
Dim db As New ADODB.Connection
Dim Rs As New ADODB.Recordset

Private Sub cmdexit_Click()
End
End Sub

Private Sub CommandAccEdit_Click(Index As Integer)
'Unload Me
frmedit.Show
End Sub

Private Sub CommandAccountInd_Click(Index As Integer)
'frmIndivdualAcc.Show

'crReport4.ParameterFields(0) = "{?accno}" & D0001
'crReport4.ParameterFields(1) = "{?accno1}" & D0002
crReport4.SelectionFormula = "{tran.acc_no} = {?accno}"
crReport4.SelectionFormula = "{tran.acc_no} = {?accno1}"
'crReport4.ParameterFields(0) = "{?accno1}';'D0002';'false"
crReport4.RetrieveSQLQuery
crReport4.Action = 1



End Sub

Private Sub CommandAccountList_Click(Index As Integer)
'DataReport2.Show
'Unload Me
  'crReport1.DiscardSavedData = 1
 
 
 
 'crReport1.PrinterSelect
  crReport1.RetrieveSQLQuery
  crReport1.Action = 1
End Sub

Private Sub CommandOpen_Click(Index As Integer)
'Unload Me
With newacc
.Show
.SetFocus
.txtAcc_No.Locked = True
End With
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\bankproject\bank.mdb;Persist Security Info=False"
Set Rs.ActiveConnection = db
Set rs_temp = New ADODB.Recordset
Rs.Open "Select acc_no from initial "

 If Rs.EOF And Rs.BOF Then
      newacc.txtAcc_No = "D0001"
 Else
      rs_temp.Open "select max(acc_no) from initial", db, adOpenDynamic, adLockOptimistic
      old_id = rs_temp(0)
      temp = Right(old_id, 4)
      temp = temp + 1
      new_id = "D" & Right("0000" & temp, 4)
      newacc.txtAcc_No = new_id
 End If
 db.Close
 'Rs.Close
 
 Unload Me
End Sub

Private Sub CommandReportMon_Click(Index As Integer)
'DataReport3.Show
'Unload Me

'With formdate
 ' .Show
  '.SetFocus  {tran.mon} = {?chmon}
  '.dtto = Date
  '.dtfrom = DateAdd("d", -30, Date)
  '.Label1(0).Caption = "Click To Show Monthly Report "
  '.Label1(1).Visible = False
  '.Label1(2).Visible = False
  '.dtto.Visible = False
  '.dtfrom.Visible = False
 'End With
 

'crReport3.SelectionFormula = "{tran.acc_no} = {?accno1}"
'crReport3.SelectionFormula = "{tran.mon} = {?chmon}"
crReport3.SelectionFormula = "{tran.mon}={?chmon} and {tran.acc_no}={?accno}"
crReport3.SelectionFormula = "{tran.mon}={?chmon1} and {tran.acc_no}={?accno1}"

crReport3.RetrieveSQLQuery
crReport3.Action = crRunReport
End Sub

Private Sub CommandTranDaily_Click(Index As Integer)
'Unload Me
frmAccCheck.Show
'frmTranDaily.Show
End Sub



VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTranDaily 
   Caption         =   "Daily Transction Form"
   ClientHeight    =   6600
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   ScaleHeight     =   6600
   ScaleWidth      =   8835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Commandback 
      Caption         =   "Back"
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
      Left            =   5760
      TabIndex        =   27
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdedit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdfirst 
      BackColor       =   &H00C0C0C0&
      Caption         =   " &First"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdnext 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdprevious 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Previous"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdlast 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Last"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4560
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd/mm/yy"
      Format          =   24444931
      CurrentDate     =   37567
      MaxDate         =   120896
      MinDate         =   33237
   End
   Begin VB.TextBox txtchno 
      Height          =   375
      Left            =   6600
      TabIndex        =   8
      Top             =   3600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox txtpar 
      Height          =   375
      Left            =   2640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   4440
      Width           =   4095
   End
   Begin MSAdodcLib.Adodc dailytran 
      Height          =   375
      Left            =   120
      Top             =   6240
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\bankproject\bank.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\bankproject\bank.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "initial"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Timer Timautofill 
      Left            =   4080
      Top             =   6240
   End
   Begin VB.ComboBox cmbtranmode 
      Height          =   315
      Left            =   4560
      TabIndex        =   5
      Text            =   "Cash"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ComboBox cmbtrantype 
      Height          =   315
      Left            =   4560
      TabIndex        =   1
      Text            =   "Deposit"
      Top             =   1080
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1800
      Top             =   6240
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\bankproject\bank.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\bankproject\bank.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "initial"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox txtaccno 
      BackColor       =   &H80000000&
      DataField       =   "acc_no"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton CommandCancle 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton CommandOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   9
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox txttranamount 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox txttranmode 
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txttrandate 
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txttrantype 
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label9 
      Caption         =   "Cheque No"
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
      Left            =   5160
      TabIndex        =   20
      Top             =   3600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Particulars"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Account No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   18
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "( By Cash / Cheque )"
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
      Left            =   480
      TabIndex        =   17
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "( Deposit / Withdraw )"
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
      Left            =   480
      TabIndex        =   16
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Amount"
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
      Left            =   480
      TabIndex        =   15
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Tranction Mode"
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
      Left            =   480
      TabIndex        =   14
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Tranction Date"
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
      Left            =   480
      TabIndex        =   13
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Tranction Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   12
      Top             =   1080
      Width           =   2055
   End
End
Attribute VB_Name = "frmTranDaily"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim rs_staff As ADODB.Recordset


Private Sub Commandback_Click()
Unload Me
End Sub

Private Sub CommandCancle_Click()
Unload Me
With frmAccCheck
.Show
.SetFocus
End With
End Sub
Private Sub CommandOk_Click()
Dim SQL
Dim SQL1
Dim SQL2
Dim depositamount As Long
Dim withdrawamount As Long
Dim bal As Long
Dim bal1 As Long
Dim tbal As Long
Dim pbal As Long
Dim m1
m1 = Month(txttrandate.Text)
'MsgBox "Month is " & m1
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\bankproject\bank.mdb;Persist Security Info=False"
Set Rs.ActiveConnection = db

Rs.Open "Select currentbalance from initial where acc_no= '" & txtaccno.Text & "'"
If frmTranDaily.txttrantype.Text = "Deposit" Then
depositamount = frmTranDaily.txttranamount.Text
Else
withdrawamount = frmTranDaily.txttranamount.Text
balch = Rs.Fields("currentbalance") - 500
'MsgBox "balance should be minimum Rs." & balch
If balch >= withdrawamount Then
'MsgBox " Goto Transaction"
Else
MsgBox " Your Current balance is  Rs." & Rs.Fields("currentBalance")
MsgBox " Balance should be minimum 500 Rs." & vbCritical
db.Close
Exit Sub
End If
End If

 'Rs.Open "Select currentbalance from initial where acc_no= '" & txtaccno.Text & "'"
'MsgBox " Old Balance is:" & Rs.Fields("balance")

Do Until Rs.EOF
pbal = Rs.Fields("currentbalance")
MsgBox "Previous balance:" & pbal
tbal = Rs.Fields("currentbalance") + depositamount - withdrawamount
    If txttrantype.Text = "Deposit" Then
     bal = Rs.Fields("currentbalance") + txttranamount.Text
      Else
      bal = Rs.Fields("currentbalance") - txttranamount.Text
      End If
    Rs.MoveNext
Loop

 Rs.Close
 If txtchno.Visible = True Then
 txtchno.Text = frmTranDaily.txtchno.Text
 Else
 txtchno.Text = " "
 End If
SQL = "insert into tran(acc_no,tran_type, tran_date,mon,tran_mode, " & _
      "chequeno,deposit_amount,withdraw_amount,particulars,balance,pbal,tbal)" & _
      " values('" & _
      txtaccno.Text & "','" & _
      txttrantype.Text & "', '" & _
      txttrandate.Text & "', '" & _
       m1 & "', '" & _
      txttranmode.Text & "', '" & _
      txtchno.Text & "','" & _
      depositamount & "', '" & _
      withdrawamount & "', '" & _
      txtpar.Text & "','" & _
      bal & "', '" & _
      pbal & "', '" & _
      tbal & "')"
       db.Execute SQL
             
 SQL1 = "Update initial Set currentbalance ='" & bal & "' Where acc_no = '" & txtaccno.Text & "' "
        db.Execute SQL1
       
        MsgBox "Tranction is Compleeted Successfully"
        'frmmain1.Show
        Unload Me
        With frmAccCheck
        .Show
        .SetFocus
        End With
        frmAccCheck.txtAcc_No.Text = frmTranDaily.txtaccno.Text
        frmAccCheck.txtAcc_No.SetFocus
        db.Close
        
    Set db = Nothing
End Sub



Private Sub Form_Load()
On Error Resume Next
'db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\bankproject\bank.mdb;Persist Security Info=False"
'Set Rs.ActiveConnection = db
'Rs.Open "select * from tran", db, adOpenDynamic, adLockOptimistic
frmTranDaily.txttrandate.Locked = True
frmTranDaily.txttrantype.Locked = True
frmTranDaily.txttranmode.Locked = True
frmTranDaily.cmbtranmode.AddItem "Cheque"
frmTranDaily.cmbtranmode.AddItem "Cash"
frmTranDaily.cmbtrantype.AddItem "Deposit"
frmTranDaily.cmbtrantype.AddItem "Withdraw"
End Sub


Private Sub cmbtrantype_Click()

    Timautofill.Enabled = False
    txttrantype.Text = cmbtrantype.List(cmbtrantype.ListIndex)
        On Error Resume Next
End Sub

Private Sub cmbtrantype_Change()
On Error Resume Next
AutoComplete.cmbtrantype

End Sub


Private Sub cmbtrantype_DropDown()

    Timautofill.Enabled = True
End Sub

Private Sub cmbtranmode_Click()

    Timautofill.Enabled = False
    txttranmode.Text = cmbtranmode.List(cmbtranmode.ListIndex)
    
    On Error Resume Next
End Sub

Private Sub cmbtranmode_Change()
On Error Resume Next
AutoComplete.cmbtranmode
End Sub


Private Sub cmbtranmode_DropDown()

    Timautofill.Enabled = True
End Sub

Public Sub max()
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\bankproject\bank.mdb;Persist Security Info=False"
Set Rs.ActiveConnection = db

Rs.Open "Select max(tran_date) from tran where acc_no= '" & txtaccno.Text & "'"
MsgBox " Latest Transaction Date is:" & Rs.Fields("tran_date")

    Rs.MoveNext

 Rs.Close
End Sub



Private Sub txttranamount_KeyPress(KeyAscii As Integer)
Const Numbers$ = "0123456789."
    If KeyAscii <> 8 Then
       If InStr(Numbers, Chr(KeyAscii)) = 0 Then
            MsgBox _
   "Only numbers allowed.", _
   vbOKOnly + vbInformation, _
   " "
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

Private Sub txttranmode_Change()
If txttranmode.Text = "Cheque" Then
    txtchno.Visible = True
    txtchno.Enabled = True
    Label9.Visible = True
    Else
    txtchno.Visible = False
    txtchno.Enabled = False
    Label9.Visible = False
    End If
End Sub


Private Sub DTPicker1_Change()
txttrandate.Text = DTPicker1.Value
End Sub


'**********************************************************************************************
'Modify existing record
'**********************************************************************************************
Private Sub cmdedit_Click()
On Error GoTo erredit
cmdsave.Enabled = True
cmdedit.Enabled = False
CommandOk.Enabled = False
  cmdfirst.Enabled = False
  cmdlast.Enabled = False
  cmdnext.Enabled = False
  cmdprevious.Enabled = False
 txtaccno.Locked = True
'txttrantype.Locked = True
txttranamount.Enabled = True
txtpar.Enabled = True
txttrantype.Enabled = True
txttranmode.Enabled = True
txtchno.Enabled = True
edt = True
Exit Sub
erredit:
MsgBox Err.Description
End Sub

'**********************************************************************************************
'Move to the first record
'**********************************************************************************************
Private Sub cmdfirst_Click()
On Error GoTo errfirst
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\bankproject\bank.mdb;Persist Security Info=False"
Set Rs.ActiveConnection = db
Rs.Open "select * from tran", db, adOpenDynamic, adLockOptimistic
Call move_in_records(Rs, "movefirst", cmdfirst, cmdnext, cmdprevious, cmdlast)
Call display
cmdedit.Enabled = True
'cmddelete.Enabled = True
Exit Sub
errfirst:
MsgBox Err.Description

End Sub

'**********************************************************************************************
'Move to the lest record
'**********************************************************************************************
Private Sub cmdlast_Click()
On Error GoTo errlast

Call move_in_records(Rs, "movelast", cmdfirst, cmdnext, cmdprevious, cmdlast)
Call display
cmdedit.Enabled = True
'cmddelete.Enabled = True
Exit Sub
errlast:
MsgBox Err.Description

End Sub

'**********************************************************************************************
'ove to the next record
'**********************************************************************************************
Private Sub cmdnext_Click()
On Error GoTo errnext

Call move_in_records(Rs, "movenext", cmdfirst, cmdnext, cmdprevious, cmdlast)
Call display
cmdedit.Enabled = True
'cmddelete.Enabled = True
Exit Sub
errnext:
MsgBox Err.Description
End Sub

'**********************************************************************************************
'Move to the previous record
'**********************************************************************************************
Private Sub cmdprevious_Click()
On Error GoTo errprevious

Call move_in_records(Rs, "moveprevious", cmdfirst, cmdnext, cmdprevious, cmdlast)
Call display
cmdedit.Enabled = True
'cmddelete.Enabled = True
Exit Sub
errprevious:
MsgBox Err.Description

End Sub

'**********************************************************************************************
'Save record
'**********************************************************************************************
Private Sub cmdsave_Click()
'Dim edt As Boolean
'On Error GoTo errsave
'Set rs_staff = New ADODB.Recordset
'rs_staff.Open "select * from initial", db, adOpenDynamic, adLockOptimistic
s = MsgBox("Do you want to save this record or not", vbYesNo)
If s = 6 Then
  If edt = True Then
     
     Rs.Fields("acc_no") = UCase(txtaccno.Text)
     Rs.Fields("tran_type") = UCase(txttrantype.Text)
     Rs.Fields("tran_date") = UCase(txttrandate.Text)
     Rs.Fields("particulars") = UCase(txtpar.Text)
     Rs.Fields("chequeno") = UCase(txtchno.Text)
     Rs.Fields("tran_mode") = UCase(txttranmode.Text)
     
     If txttrantype.Text = "DEPOSIT" Then
     Rs.Fields("deposit_amount") = UCase(txttranamount.Text)
     Else
     Rs.Fields("withdraw_amount") = UCase(txttranamount.Text)
     End If
     
     Rs.Update
  Else
    
     Rs.Fields("acc_no") = UCase(txtaccno.Text)
     Rs.Fields("tran_type") = UCase(txttrantype.Text)
     Rs.Fields("tran_date") = UCase(txttrandate.Text)
     Rs.Fields("particulars") = UCase(txtpar.Text)
     Rs.Fields("chequeno") = UCase(txtchno.Text)
     Rs.Fields("tran_mode") = UCase(txttranmode.Text)
     
     If txttrantype.Text = "DEPOSIT" Then
     Rs.Fields("deposit_amount") = UCase(txttranamount.Text)
     Else
     Rs.Fields("withdraw_amount") = UCase(txttranamount.Text)
     End If
     
     Rs.Update
  End If
End If

cmdnext.Enabled = True
cmdfirst.Enabled = True
cmdprevious.Enabled = True
cmdlast.Enabled = True
cmdsave.Enabled = False
'cmdadd.Enabled = True
cmdedit.Enabled = False
Exit Sub
errsave:
MsgBox Err.Description
End Sub

'**********************************************************************************************
Public Sub display()
txttranmode.Text = UCase(Rs.Fields("tran_mode"))
txtaccno.Text = UCase(Rs.Fields("acc_no"))
txttrantype.Text = UCase(Rs.Fields("tran_type"))
'MsgBox "trantype" & txttrantype.Text
If txttrantype.Text = "DEPOSIT" Then
txttranamount.Text = UCase(Rs.Fields("deposit_amount"))
Else
'txttranamount.Text = UCase(Rs.Fields("withdraw_amount"))
End If

txttrandate.Text = UCase(Rs.Fields("tran_date"))
txtchno.Text = UCase(Rs.Fields("chequeno"))

txtpar.Text = UCase(Rs.Fields("particulars"))

End Sub

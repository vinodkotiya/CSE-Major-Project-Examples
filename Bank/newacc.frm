VERSION 5.00
Begin VB.Form newacc 
   Caption         =   "New Account Open Form"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAcc_No 
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox txtName 
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox txtAddress 
      Height          =   405
      Left            =   2760
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2160
      Width           =   2535
   End
   Begin VB.TextBox txtBalance 
      Height          =   375
      Left            =   2760
      MaxLength       =   100
      TabIndex        =   2
      Top             =   3240
      Width           =   2415
   End
   Begin VB.CommandButton CommandOk 
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
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton CommandCancle 
      Caption         =   "Cancel"
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
      TabIndex        =   4
      Top             =   4080
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4440
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4440
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
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
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4080
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label LableAccNo 
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
      Left            =   840
      TabIndex        =   15
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label LabelName 
      Caption         =   "Name"
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
      Left            =   840
      TabIndex        =   14
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label LabelAdd 
      Caption         =   "Address"
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
      Left            =   840
      TabIndex        =   13
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label LabelBal 
      Caption         =   "Balance"
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
      Left            =   840
      TabIndex        =   12
      Top             =   3240
      Width           =   1935
   End
End
Attribute VB_Name = "newacc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim rs_staff As ADODB.Recordset
Private Sub CommandCancle_Click()
Unload Me
frmmain1.Show
frmmain1.SetFocus
End Sub

Private Sub Form_Load()
Dim rs_staff As ADODB.Recordset
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\bankproject\bank.mdb;Persist Security Info=False"

Set rs_staff = New ADODB.Recordset
rs_staff.Open "select * from initial", db, adOpenDynamic, adLockOptimistic
  
  cmdsave.Enabled = False
  'cmdSearch.Visible = False
  cmdedit.Enabled = False
  cmdfirst.Enabled = False
  cmdnext.Enabled = False
  cmdlast.Enabled = False
  cmdprevious.Enabled = False
  
End Sub

Private Sub CommandOk_Click()
Dim SQL
Dim SQL1
Dim today
Dim m1

today = Format(Now, "short date")
m1 = Month(today)
 
        If txtName.Text = "" Then
           MsgBox "Name Should Not be Blank ,Enter Your Name.", vbCritical + vbOKOnly
           txtName.SetFocus
           Exit Sub
        End If
        If txtAddress.Text = "" Then
           MsgBox "Address Should Not be Blank ,Enter Your Address.", vbCritical + vbOKOnly
           txtAddress.SetFocus
           Exit Sub
        End If
    balance = txtBalance.Text
     If balance < 500 Then
  MsgBox " Initial Deposit should Not be less than Rs.500"
  txtBalance.SetFocus
  Exit Sub
End If

'db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\bankproject\bank.mdb;Persist Security Info=False"
SQL = "insert into initial(acc_date,acc_no, name, address, " & _
      "balance,currentbalance)" & _
      " values('" & _
      today & "', '" & _
      txtAcc_No.Text & "', '" & _
      txtName.Text & "', '" & _
      txtAddress.Text & "', '" & _
      txtBalance.Text & "', '" & _
      txtBalance.Text & "')"
        db.Execute SQL
        
SQL1 = "insert into tran(acc_no,tran_date,mon, " & _
     "initial,balance,tbal)" & _
      " values('" & _
      txtAcc_No.Text & "', '" & _
      today & "', '" & _
      m1 & "', '" & _
      txtBalance.Text & "', '" & _
      txtBalance.Text & "', '" & _
      txtBalance.Text & "')"
        db.Execute SQL1
        MsgBox "New Account is Opened Successfully"
        frmmain1.Show
        Unload Me
        'db.Close
    Set db = Nothing
End Sub

'**********************************************************************************************
'Modify existing record
'**********************************************************************************************
Private Sub cmdedit_Click()
On Error GoTo erredit
cmdsave.Enabled = True
cmdedit.Enabled = False
CommandOk.Enabled = False
  
 txtAcc_No.Locked = True
 txtName.Enabled = True
 txtAddress.Enabled = True
 txtBalance.Locked = True

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
Call move_in_records(rs_staff, "movefirst", cmdfirst, cmdnext, cmdprevious, cmdlast)
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
Call move_in_records(rs_staff, "movelast", cmdfirst, cmdnext, cmdprevious, cmdlast)
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
Set rs_staff = New ADODB.Recordset
rs_staff.Open "select * from initial", db, adOpenDynamic, adLockOptimistic

On Error GoTo errnext
Call move_in_records(rs_staff, "movenext", cmdfirst, cmdnext, cmdprevious, cmdlast)
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


Call move_in_records(rs_staff, "moveprevious", cmdfirst, cmdnext, cmdprevious, cmdlast)
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
     
     rs_staff.Fields("acc_no") = UCase(txtAcc_No.Text)
     rs_staff.Fields("name") = UCase(txtName.Text)
     rs_staff.Fields("address") = UCase(txtAddress.Text)
     rs_staff.Fields("balance") = UCase(txtBalance.Text)
     
     rs_staff.Update
  Else
    
     rs_staff.Fields("acc_no") = UCase(txtAcc_No.Text)
     rs_staff.Fields("name") = UCase(txtName.Text)
     rs_staff.Fields("address") = UCase(txtAddress.Text)
     rs_staff.Fields("balance") = UCase(txtBalance.Text)
     rs_staff.Update
  End If
End If

cmdnext.Enabled = True
cmdfirst.Enabled = True
cmdprevious.Enabled = True
cmdlast.Enabled = True
cmdsave.Enabled = False
'cmdadd.Enabled = True
cmdedit.Enabled = True
Exit Sub
errsave:
MsgBox Err.Description
End Sub

'**********************************************************************************************
Public Sub display()

txtAcc_No.Text = UCase(rs_staff.Fields("acc_no"))
txtName.Text = UCase(rs_staff.Fields("name"))
txtAddress.Text = UCase(rs_staff.Fields("address"))
txtBalance.Text = UCase(rs_staff.Fields("balance"))

End Sub

Private Sub Form_Unload(Cancel As Integer)
db.Close
End Sub

Private Sub txtBalance_KeyPress(KeyAscii As Integer)
'Call AllowOnlyIntegers(KeyAscii)
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



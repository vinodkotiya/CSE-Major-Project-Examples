VERSION 5.00
Begin VB.Form frmmodtran 
   Caption         =   "Modify Transaction"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Commandcancel 
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
      Left            =   3000
      TabIndex        =   4
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CommandButton CommandmodOk 
      Caption         =   "Modify"
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
      Left            =   480
      TabIndex        =   3
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtdate 
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txtname 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtaccno 
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   5160
      X2              =   5160
      Y1              =   240
      Y2              =   3120
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   3120
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   240
      X2              =   5160
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   240
      X2              =   5160
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label3 
      Caption         =   "Date Of Transaction"
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
      Left            =   480
      TabIndex        =   7
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
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
      Left            =   480
      TabIndex        =   6
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Account No"
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
      Left            =   480
      TabIndex        =   5
      Top             =   600
      Width           =   1695
   End
End
Attribute VB_Name = "frmmodtran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Commandcancel_Click()
Unload Me
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
 txtname.Enabled = True
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
     rs_staff.Fields("name") = UCase(txtname.Text)
     rs_staff.Fields("address") = UCase(txtAddress.Text)
     rs_staff.Fields("balance") = UCase(txtBalance.Text)
     
     rs_staff.Update
  Else
    
     rs_staff.Fields("acc_no") = UCase(txtAcc_No.Text)
     rs_staff.Fields("name") = UCase(txtname.Text)
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
txtname.Text = UCase(rs_staff.Fields("name"))
txtAddress.Text = UCase(rs_staff.Fields("address"))
txtBalance.Text = UCase(rs_staff.Fields("balance"))

End Sub


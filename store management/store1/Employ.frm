VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form5 
   BackColor       =   &H80000012&
   Caption         =   "Form5"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H0080C0FF&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8280
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0080C0FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3480
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mhEmploy 
      Height          =   2895
      Left            =   720
      TabIndex        =   16
      Top             =   4800
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   5106
      _Version        =   393216
      BackColor       =   8438015
      Cols            =   6
      FixedCols       =   0
      BackColorFixed  =   8421631
      BackColorBkg    =   12640511
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H0080C0FF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H0080C0FF&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H0080C0FF&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H00FFFF80&
      Height          =   285
      Left            =   8040
      TabIndex        =   5
      Text            =   " "
      Top             =   2640
      Width           =   2415
   End
   Begin VB.TextBox txtDept 
      BackColor       =   &H00FFFF80&
      Height          =   285
      Left            =   8040
      TabIndex        =   3
      Text            =   " "
      Top             =   2040
      Width           =   2415
   End
   Begin VB.TextBox txtEmpname 
      BackColor       =   &H00FFFF80&
      Height          =   285
      Left            =   8040
      TabIndex        =   1
      Text            =   " "
      Top             =   1440
      Width           =   2415
   End
   Begin VB.TextBox txtIntercom 
      BackColor       =   &H00FFFF80&
      Height          =   285
      Left            =   2880
      TabIndex        =   4
      Text            =   " "
      Top             =   2640
      Width           =   2295
   End
   Begin VB.TextBox txtDsn 
      BackColor       =   &H00FFFF80&
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Text            =   " "
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtEmpcode 
      BackColor       =   &H00FFFF80&
      Height          =   285
      Left            =   2880
      TabIndex        =   0
      Text            =   " "
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   5
      X1              =   0
      X2              =   11880
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Record For Deletion Or Updation"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   3120
      TabIndex        =   17
      Top             =   4320
      Width           =   5655
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Entry Form "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   615
      Left            =   3840
      TabIndex        =   15
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "E- Mail "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   6000
      TabIndex        =   11
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   6000
      TabIndex        =   10
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   6000
      TabIndex        =   9
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Intercom No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   7
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Code"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAdd_Click()
sql = "insert into Employ values(" & Val(txtEmpcode) & ", '" & txtEmpname & "',"
sql = sql & " '" & txtDsn & "','" & txtDept & "'," & Val(txtIntercom) & ",'" & txtEmail & "')"
cn.Execute sql
cmdReset_Click
Call Setting
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
mhEmploy.RemoveItem (mhEmploy.RowSel)
cn.Execute ("delete from Employ where empcode = " & Val(txtEmpcode) & " ")
Call Setting
End Sub


Private Sub cmdReset_Click()
txtEmpcode = ""
txtEmpname = ""
txtDsn = ""
txtDept = ""
txtIntercom = ""
txtEmail = ""
End Sub

Private Sub cmdUpdate_Click()
mhEmploy.TextMatrix(mhEmploy.RowSel, 0) = txtEmpcode
mhEmploy.TextMatrix(mhEmploy.RowSel, 1) = txtEmpname
mhEmploy.TextMatrix(mhEmploy.RowSel, 2) = txtDsn
mhEmploy.TextMatrix(mhEmploy.RowSel, 3) = txtDept
mhEmploy.TextMatrix(mhEmploy.RowSel, 4) = txtIntercom
mhEmploy.TextMatrix(mhEmploy.RowSel, 5) = txtEmail
cn.Execute ("delete from Employ where empcode = " & Val(txtEmpcode) & " ")
cmdAdd_Click
Form_Load
End Sub

Private Sub Form_Load()
Call Connect
Call Setting
cmdAdd.Enabled = False
cmdDelete.Enabled = False
cmdUpdate.Enabled = False
End Sub

Private Sub mhEmploy_Click()
If MDIForm1.Flag1 = False Then
txtEmpcode = mhEmploy.TextMatrix(mhEmploy.Row, 0)
txtEmpname = mhEmploy.TextMatrix(mhEmploy.Row, 1)
txtDsn = mhEmploy.TextMatrix(mhEmploy.Row, 2)
txtDept = mhEmploy.TextMatrix(mhEmploy.Row, 3)
txtIntercom = mhEmploy.TextMatrix(mhEmploy.Row, 4)
txtEmail = mhEmploy.TextMatrix(mhEmploy.Row, 5)
cmdDelete.Enabled = True
cmdUpdate.Enabled = True
End If
End Sub

Private Sub txtDept_KeyPress(KeyAscii As Integer)
cmdAdd.Enabled = True
End Sub

Private Sub txtEmpcode_LostFocus()
If txtEmpcode = "" And ActiveControl <> cmdCancel Then
MsgBox "Please Enter Employee Code", vbInformation, "Employee Entry"
txtEmpcode.SetFocus
Else
Set rs1 = cn.Execute("select empname from Employ where empcode = " & Val(txtEmpcode) & " ")
If Not rs.EOF And MDIForm1.Eflag = True Then
MsgBox "Invalid Employee Code", vbInformation, "Employee Entry"
txtEmpcode.SetFocus
End If
End If
End Sub
Public Sub Setting()
mhEmploy.ColWidth(0) = 900
mhEmploy.ColWidth(1) = 2300
mhEmploy.ColWidth(2) = 1700
mhEmploy.ColWidth(3) = 1700
mhEmploy.ColWidth(4) = 1200
mhEmploy.ColWidth(5) = 2000
mhEmploy.TextMatrix(0, 0) = "EmpCode"
mhEmploy.TextMatrix(0, 1) = "Employee Name"
mhEmploy.TextMatrix(0, 2) = "Designation"
mhEmploy.TextMatrix(0, 3) = "Department"
mhEmploy.TextMatrix(0, 4) = "Intercom No."
mhEmploy.TextMatrix(0, 5) = "E-Mail"
Set rs = cn.Execute("select * from Employ order by empname")
Set mhEmploy.DataSource = rs
End Sub

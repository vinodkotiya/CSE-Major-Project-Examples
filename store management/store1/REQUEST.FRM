VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   Caption         =   "Requisition Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "Request.frx":0000
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSComCtl2.DTPicker dtpCdate 
      Height          =   375
      Left            =   7800
      TabIndex        =   23
      Top             =   1320
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   22609923
      CurrentDate     =   37407
   End
   Begin VB.TextBox txtRqty 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7800
      TabIndex        =   4
      Text            =   " "
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox txtBreif 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2880
      TabIndex        =   5
      Text            =   " "
      Top             =   3240
      Width           =   7455
   End
   Begin VB.TextBox txtRegno 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2880
      TabIndex        =   11
      Text            =   " "
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtEmpid 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2880
      TabIndex        =   0
      Text            =   " "
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtEname 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7800
      TabIndex        =   1
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox txtDept 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7800
      TabIndex        =   2
      Text            =   " "
      Top             =   2280
      Width           =   2535
   End
   Begin VB.ComboBox cmbItem 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2880
      TabIndex        =   3
      Text            =   " "
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0C0C0&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mhIndent 
      Height          =   2295
      Left            =   960
      TabIndex        =   18
      Top             =   5160
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   4048
      _Version        =   393216
      BackColor       =   16777152
      Rows            =   20
      Cols            =   5
      BackColorFixed  =   16776960
      BackColorSel    =   12632064
      BackColorBkg    =   16777088
      BackColorUnpopulated=   12632064
      WordWrap        =   -1  'True
      FillStyle       =   1
      MousePointer    =   4
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
      _Band(0).Cols   =   5
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT RECORD FOR DELETION (  IN CASE OF MISTAKE DURING ENTRY )"
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
      Height          =   495
      Left            =   1320
      TabIndex        =   22
      Top             =   4800
      Width           =   8895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   5
      X1              =   0
      X2              =   11880
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " REQUISITION FORM"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   615
      Left            =   2280
      TabIndex        =   21
      Top             =   120
      Width           =   6855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Requested Quantity"
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
      Left            =   5640
      TabIndex        =   20
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Breif Description"
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
      Left            =   960
      TabIndex        =   19
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Indent No."
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
      Left            =   960
      TabIndex        =   17
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C000&
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
      Height          =   495
      Left            =   960
      TabIndex        =   16
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C000&
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
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C000&
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
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   2280
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Item List"
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
      Left            =   960
      TabIndex        =   13
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "Current Date"
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
      Left            =   5640
      TabIndex        =   12
      Top             =   1320
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nextreg As Integer, sno As Integer, i As Integer, Tempsno As Integer, nextid As Integer
Dim Eflag As Boolean, Flag1 As Boolean, Flag2 As Boolean
Dim Msg As String, Ireg As String
Dim delrec As Integer
Dim rno As Integer

Private Sub cmbItem_LostFocus()
If Me.ActiveControl <> cmdExit And Me.ActiveControl <> mhIndent And Me.ActiveControl <> cmdReset Then
If Trim(cmbItem.Text) = "" Then
MsgBox "You Must Select Or Enter An Item Name", vbInformation, "Requisition"
cmbItem.SetFocus
Else
 Set rs = cn.Execute("select itemid from Item where itemname = '" & cmbItem & "' ")
 If rs.EOF Then
 Set rs1 = cn.Execute("select max(itemid) as id from Item")
 If IsNull(rs1.Fields("id")) = True Then
 nextid = 1
 Else
 nextid = rs1.Fields("id") + 1
 End If
 cn.Execute ("Insert into Item(itemid,itemname) values(" & nextid & ",'" & cmbItem & "') ")
 End If
 txtRqty.SetFocus
 cmdOk.Enabled = True
End If
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If mhIndent.Rows > 2 Then 'And Val(mhIndent.TextMatrix(mhIndent.RowSel, 1)) <> 0 Then
If Flag1 = True Then
mhIndent.RemoveItem (mhIndent.RowSel)
Flag2 = True
i = 1
While Val(mhIndent.TextMatrix(i, 1)) <> 0
mhIndent.TextMatrix(i, 1) = i
i = i + 1
Wend
sno = i
Else
MsgBox "Select Record From List For Deletion", vbInformation, "Requisition"
End If
End If
'End If
End Sub

Private Sub cmdOk_Click()
If sno > mhIndent.Rows - 1 Then
mhIndent.Rows = mhIndent.Rows + 5
End If
mhIndent.TextMatrix(sno, 1) = sno
mhIndent.TextMatrix(sno, 2) = cmbItem.Text
mhIndent.TextMatrix(sno, 3) = txtRqty
mhIndent.TextMatrix(sno, 4) = txtBreif
sno = sno + 1
Call Clear
cmdSend.Enabled = True
End Sub

Private Sub cmdReset_Click()
txtRegno = ""
'txtCdate = ""
txtEmpid = ""
txtEname = ""
txtDept = ""
'txtItem = ""
cmbItem.Clear
txtRqty = ""
txtBreif = ""
Form_Load
txtEmpid.SetFocus
End Sub

Private Sub cmdSend_Click()
'If mhIndent.Rows > 2 Then
Msg = MsgBox("Are You Sure About Entered Request", vbQuestion + vbYesNo, "Requisition")

If Msg = vbNo Then
Me.Show
End If

If Msg = vbYes Then
cn.Execute ("insert into Indent(indentno,request_date,empcode) values(" & Val(txtRegno) & ",#" & Format(dtpCdate.Value, "dd-MMM-yy") & "#," & Val(txtEmpid) & ")")
Set rs = cn.Execute("select max(rno) as nextrno from ItemRequest")

If IsNull(rs.Fields("nextrno")) = True Then
rno = 1
Else
rno = rs.Fields("nextrno") + 1
End If

i = 1
While Val(mhIndent.TextMatrix(i, 1)) <> 0
sql = "insert into ItemRequest(rno,indentno,itemname,breif_r,qty_request) values( " & rno & "," & Val(txtRegno) & ",'" & mhIndent.TextMatrix(i, 2) & "','" & mhIndent.TextMatrix(i, 4) & "'," & mhIndent.TextMatrix(i, 3) & "   ) "
cn.Execute (sql)
i = i + 1
rno = rno + 1
Wend
cmdReset_Click
mhIndent.Clear
End If
'End If
'Errorhandler:
'If Err.Number = -2147217900 Then
'Resume Next
'End If
End Sub

Private Sub Form_Load()
Call Connect
dtpCdate.Value = Format(Now, "dd-MMM-yy")
Flag1 = False
Flag2 = False
cmdOk.Enabled = False
Set rs = cn.Execute("select max(indentno) as reg from Indent")
If IsNull(rs.Fields("reg")) Then
nextreg = 1
Else
nextreg = rs.Fields("reg") + 1
End If
txtRegno = nextreg
sno = 1
mhIndent.ColWidth(0) = 200
mhIndent.ColWidth(1) = 600
mhIndent.ColWidth(2) = 1900
mhIndent.ColWidth(3) = 800
mhIndent.ColWidth(4) = 4450
mhIndent.TextMatrix(0, 1) = "Sr.No"
mhIndent.TextMatrix(0, 2) = "Item Name"
mhIndent.TextMatrix(0, 3) = "Quantity"
mhIndent.TextMatrix(0, 4) = "Breif Description"
End Sub


Private Sub mhIndent_Click()
Flag1 = True
cmdDelete.Refresh
End Sub

Private Sub txtBreif_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then
KeyAscii = 0
End If
End Sub

Private Sub txtDept_LostFocus()
txtDept = Initcap(txtDept)
If Eflag = False Then
cn.Execute ("Insert into Employ (empcode,empname,dept) values(" & Val(txtEmpid) & ",'" & txtEname & "','" & txtDept & "')")
Eflag = True
End If
End Sub
Private Sub txtEmpid_KeyPress(KeyAscii As Integer)
Call NumberOnly(KeyAscii)
End Sub

Private Sub txtEmpid_LostFocus()
Set rs = cn.Execute("select empname,dept from Employ where empcode = " & Val(txtEmpid) & " ")
If Not rs.EOF Then
txtEname = rs.Fields("empname")
txtDept = rs.Fields("dept")
cmbItem.SetFocus
Eflag = True
Else
txtEname.SetFocus
Eflag = False
End If
Set rs1 = cn.Execute("select * from Item ")
While Not rs1.EOF
cmbItem.AddItem rs1.Fields("itemname")
rs1.MoveNext
Wend
End Sub
Public Sub Clear()
txtRegno = ""
cmbItem.Text = ""
'txtItem = ""
txtRqty = ""
txtBreif = ""
txtRegno = nextreg
cmbItem.SetFocus
cmdOk.Enabled = False
End Sub
Private Sub txtEname_LostFocus()
txtEname = Initcap(txtEname)
End Sub
Private Sub txtRqty_KeyPress(KeyAscii As Integer)
Call NumberOnly(KeyAscii)
End Sub

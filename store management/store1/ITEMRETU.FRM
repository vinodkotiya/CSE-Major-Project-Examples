VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form16 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Form16"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form16"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPrint 
      Caption         =   "PRINT "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   24
      Top             =   3600
      Width           =   1695
   End
   Begin VB.TextBox txtSno 
      Height          =   375
      Left            =   2160
      TabIndex        =   23
      Text            =   " "
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "RESET"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   21
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   11
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "PREVIEW"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   20
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   19
      Top             =   4800
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
      Height          =   1695
      Left            =   120
      TabIndex        =   17
      Top             =   5400
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   2990
      _Version        =   393216
      Cols            =   6
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin VB.ComboBox cmbSno 
      Height          =   315
      Left            =   2160
      TabIndex        =   9
      Top             =   3600
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dtpRdate 
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   3000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   24641539
      CurrentDate     =   37406
   End
   Begin VB.TextBox txtRqty 
      Height          =   375
      Left            =   6960
      TabIndex        =   10
      Text            =   " "
      Top             =   3600
      Width           =   2295
   End
   Begin VB.ComboBox cmbItem 
      Height          =   315
      Left            =   6960
      TabIndex        =   8
      Top             =   3000
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "TRANSFERED ITEM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   8160
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "ISSUED ITEM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   5280
      TabIndex        =   0
      Top             =   1200
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.ComboBox cmbDept 
      Height          =   315
      Left            =   6960
      TabIndex        =   6
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox txtEmpid 
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Text            =   " "
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sr. No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   22
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      Index           =   1
      X1              =   0
      X2              =   11880
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Record Of Issued/Transfered Items"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   4920
      Width           =   4215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Return Quantity"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   4440
      TabIndex        =   16
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Return Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Indent No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   4440
      TabIndex        =   13
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      Index           =   0
      X1              =   0
      X2              =   11880
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Type Of Item ( Issued/Transfered)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   4
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM RETURN FORM"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Indname As String
Dim dept As String
Dim i As Integer
Dim Iqty As Integer
Dim empid As Integer
Dim Dflag As Boolean, Iflag As Boolean
Dim res As VbMsgBoxResult

Private Sub cmbDept_Click()
If Option1(1).Value = True Then
sql = "SELECT ItemRequest.itemname, ItemIssue.breif_i, Indent.empcode, Employ.dept "
sql = sql & "FROM ((ItemRequest INNER JOIN ItemIssue ON ItemRequest.rno = ItemIssue.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.transfer = True and Indent.empcode in(select empcode from Employ where dept = '" & cmbDept.Text & "') order by itemname ;"
Set rs1 = cn.Execute(sql)
cmbItem.Clear
If Not rs1.EOF Then
cmbItem.Text = rs1.Fields("itemname")
End If
While Not rs1.EOF
cmbItem.AddItem rs1.Fields("itemname")
rs1.MoveNext
Wend
Dflag = True
End If
End Sub

Private Sub cmbDept_LostFocus()
If Option1(1).Value = True And Dflag = False Then
sql = "SELECT ItemRequest.itemname, ItemIssue.breif_i, Indent.empcode, Employ.dept "
sql = sql & "FROM ((ItemRequest INNER JOIN ItemIssue ON ItemRequest.rno = ItemIssue.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.transfer = True and Indent.empcode in(select empcode from Employ where dept = '" & cmbDept.Text & "') order by itemname ;"
Set rs1 = cn.Execute(sql)
cmbItem.Clear
If Not rs1.EOF Then
cmbItem.Text = rs1.Fields("itemname")
End If
While Not rs1.EOF
cmbItem.AddItem rs1.Fields("itemname")
rs1.MoveNext
Wend
End If
End Sub

Private Sub cmbItem_Click()
If Option1(1).Value = True Then
sql = "SELECT ItemIssue.rno,ItemRequest.ItemName,ItemIssue.Breif_i ,ItemIssue.qty_issue,ItemIssue.stocksno,Indent.empcode,Employ.Dept"
sql = sql & " FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.transfer = True and ItemRequest.itemname = '" & cmbItem.Text & "' and Indent.empcode in(select empcode from Employ where dept = '" & cmbDept.Text & "') ;"
End If
If Option1(0).Value = True Then
sql = "SELECT ItemIssue.rno,ItemRequest.ItemName,ItemIssue.Breif_i ,ItemIssue.qty_issue,ItemIssue.stocksno,Indent.empcode,Employ.empname"
sql = sql & " FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.issued = True and ItemRequest.itemname = '" & cmbItem.Text & "' and Indent.empcode = " & Val(txtEmpid) & "  ;"
End If
Set rs1 = cn.Execute(sql)
cmbSno.Clear
If Not rs1.EOF Then
cmbSno.Text = rs1.Fields("rno")
End If
While Not rs1.EOF
cmbSno.AddItem rs1.Fields("rno")
rs1.MoveNext
Wend
Set Grid2.DataSource = rs1
Call GridSet
Iflag = True
End Sub

Private Sub cmbItem_LostFocus()
If Iflag = False Then
If Option1(1).Value = True Then
sql = "SELECT ItemIssue.rno,ItemRequest.ItemName,ItemIssue.Breif_i ,ItemIssue.qty_issue,ItemIssue.stocksno,Indent.empcode,Employ.Dept"
sql = sql & " FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.transfer = True and ItemRequest.itemname = '" & cmbItem.Text & "' and Indent.empcode in(select empcode from Employ where dept = '" & cmbDept.Text & "') ;"
End If
If Option1(0).Value = True Then
sql = "SELECT ItemIssue.rno,ItemRequest.ItemName,ItemIssue.Breif_i ,ItemIssue.qty_issue,ItemIssue.stocksno,Indent.empcode,Employ.empname"
sql = sql & " FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.issued = True and ItemRequest.itemname = '" & cmbItem.Text & "' and Indent.empcode = " & Val(txtEmpid) & "  ;"
End If
Set rs1 = cn.Execute(sql)
cmbSno.Clear
If Not rs1.EOF Then
cmbSno.Text = rs1.Fields("rno")
End If
While Not rs1.EOF
cmbSno.AddItem rs1.Fields("rno")
rs1.MoveNext
Wend
Set Grid2.DataSource = rs1
Call GridSet
End If
End Sub

Private Sub cmbSno_LostFocus()
Set rs = cn.Execute("select qty_issue,stocksno from ItemIssue where rno = " & Val(cmbSno.Text) & " ")
Iqty = rs.Fields("qty_issue")
txtSno = rs.Fields("stocksno")

'If Me.ActiveControl <> cmdReset And ActiveControl <> cmdCancel And ActiveControl <> Option1(0) And ActiveControl <> Option1(1) Then
'If Trim(cmbSno.Text) = "" Then
'MsgBox "Please Select Indent No", vbInformation, "ItemReturn"
'cmbSno.SetFocus
'End If
'End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
If Val(txtRqty) = 0 Then
MsgBox "Enter Return Quantity", vbInformation, "Item Return"
txtRqty.SetFocus
Else
 If Val(txtRqty) > Iqty Then
 MsgBox "Return Quantity Is More Than Issued Quantity", vbInformation, "Item Return"
 txtRqty.SetFocus
 Else
   dept = cmbDept.Text
   If Option1(0).Value = True Then
   cn.Execute ("update ItemIssue set qty_issue = " & Iqty - Val(txtRqty) & " , issued = false,remark = '" & txtRqty + "-" + cmbItem.Text + " " + "Returned" & "' where rno = " & Val(cmbSno.Text) & " ")
   Set rs = cn.Execute("select balance from ItemStock where stocksno = " & Val(txtSno) & " ")
   cn.Execute ("update Itemstock set balance = " & rs.Fields("balance") + Val(txtRqty) & " where stocksno = " & Val(txtSno) & " ")
   sql = "insert into ItemReturn(rno,return_date,return_qty,empcode,dept) values(" & Val(cmbSno.Text) & ",#" & Format(dtpRdate.Value, "dd-MMM-yy") & "#," & Val(txtRqty) & "," & Val(txtEmpid) & ",'" & cmbDept.Text & "' ) "
   cn.Execute (sql)
   empid = Val(txtEmpid)
   res = MsgBox("Quantity Returned " + vbCrLf + " Do You Want To Return More Item For EmpCode:" + "  " + txtEmpid, vbInformation + vbYesNo, "ItemReturn")
   If res = vbYes Then
   cmdReset_Click
   txtEmpid = empid
   txtEmpid_LostFocus
   Else
   cmdReset_Click
   End If
   End If
   

   If Option1(1).Value = True Then
   cn.Execute ("update ItemIssue set qty_issue =  " & Iqty - Val(txtRqty) & " , transfer = false,remark = '" & txtRqty + "Returned" & "' where rno = " & Val(cmbSno.Text) & " ")
   Set rs = cn.Execute("select balance from ItemStock where stocksno = " & Val(txtSno) & " ")
   cn.Execute ("update Itemstock set balance = " & rs.Fields("balance") + Val(txtRqty) & " where stocksno = " & Val(txtSno) & " ")
   sql = "insert into ItemReturn(rno,return_date,return_qty,dept) values(" & Val(cmbSno.Text) & ",#" & Format(dtpRdate.Value, "dd-MMM-yy") & "#," & Val(txtRqty) & ",'" & cmbDept.Text & "' ) "
   cn.Execute (sql)
   res = MsgBox("Quantity Returned " + vbCrLf + " Do You Want To Return More Item For " + "  " + cmbDept.Text + " " + "Department", vbInformation + vbYesNo, "ItemReturn")
   If res = vbYes Then
   cmdReset_Click
   cmbDept.Text = dept
   cmbDept_LostFocus
   Else
   cmdReset_Click
   End If
   End If
 End If
End If

End Sub

Private Sub cmdPrint_Click()
If Option1(0).Value = True Then
cn.Execute ("update ItemReturn set report = true where empcode = " & empid & " ")
End If
If Option1(1).Value = True Then
cn.Execute ("update ItemReturn set report = true where dept = '" & dept & "' ")
End If
End Sub

Private Sub cmdReport_Click()
Report8.Sections("section2").Controls("lblName1").Caption = Indname
Report8.Sections("section2").Controls("lblName2").Caption = Indname
Report8.Sections("section2").Controls("lblDept1").Caption = dept
Report8.Sections("section2").Controls("lblDept2").Caption = dept
Report8.Sections("section2").Controls("lblDate1").Caption = Format(dtpRdate.Value, "dd-MMM-yy")
Report8.Sections("section2").Controls("lblDate2").Caption = Format(dtpRdate.Value, "dd-MMM-yy")
sql = " SELECT ItemReturn.return_qty,ItemRequest.itemname,ItemIssue.breif_i "
sql = sql & "FROM (ItemRequest INNER JOIN ItemReturn ON ItemRequest.rno = ItemReturn.rno) INNER JOIN ItemIssue ON ItemRequest.rno = ItemIssue.rno where ItemReturn.report = false;"
Set rs = cn.Execute(sql)
Set Report8.DataSource = rs
Report8.Show
End Sub

Private Sub cmdReset_Click()
txtEmpid = ""
cmbDept.Clear
cmbItem.Clear
cmbSno.Clear
txtRqty = ""
txtSno = ""
Grid2.Clear
End Sub

Private Sub Form_Load()
Call Connect
dtpRdate.Value = Format(Now, "dd-MMM-yy")
End Sub


Public Sub GridSet()
Grid2.ColWidth(0) = 300
Grid2.ColWidth(1) = 700
Grid2.ColWidth(2) = 1200
Grid2.ColWidth(3) = 1600
Grid2.ColWidth(4) = 800
Grid2.ColWidth(5) = 700
Grid2.ColWidth(6) = 1200
Grid2.ColWidth(7) = 1600
Grid2.TextMatrix(0, 1) = "IndNo"
Grid2.TextMatrix(0, 2) = "Item"
Grid2.TextMatrix(0, 3) = "Breif"
Grid2.TextMatrix(0, 4) = "IssueQty"
Grid2.TextMatrix(0, 5) = "Sr.No."
Grid2.TextMatrix(0, 6) = "Ecode"
If Option1(1).Value = True Then
Grid2.TextMatrix(0, 7) = "Dept"
Else
Grid2.TextMatrix(0, 7) = "Emp Name"
End If

End Sub

Private Sub Option1_Click(Index As Integer)
cmdReset_Click
If Option1(1).Value = True Then
txtEmpid.Enabled = False
txtEmpid.Text = ""
cmbDept.SetFocus
Set rs = cn.Execute("select distinct(dept) from Employ order by dept")
cmbDept.Text = rs.Fields("dept")
While Not rs.EOF
cmbDept.AddItem rs.Fields("dept")
rs.MoveNext
Wend
End If
If Option1(0).Value = True Then
txtEmpid.Enabled = True
txtEmpid.SetFocus
End If
End Sub

Private Sub txtEmpid_LostFocus()
If Option1(0).Value = True Then
Set rs = cn.Execute("select empname,dept from employ where empcode = " & Val(txtEmpid) & " ")
If Not rs.EOF Then
cmbDept.Text = rs.Fields("dept")
Indname = rs.Fields("empname")
sql = "SELECT ItemRequest.itemname, ItemIssue.breif_i, Indent.empcode, Employ.empname"
sql = sql & " FROM ((ItemRequest INNER JOIN ItemIssue ON ItemRequest.rno = ItemIssue.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.issued = True and Indent.empcode = " & Val(txtEmpid) & " order by itemname;"
Set rs1 = cn.Execute(sql)
If Not rs1.EOF Then
cmbItem.Text = rs1.Fields("itemname")
End If
While Not rs1.EOF
cmbItem.AddItem rs1.Fields("itemname")
rs1.MoveNext
Wend
Else
If Me.ActiveControl <> cmdReset And ActiveControl <> cmdCancel Then
MsgBox "Please Check Employee Code", vbInformation, "Item Return"
txtEmpid.SetFocus
End If
End If
End If
End Sub

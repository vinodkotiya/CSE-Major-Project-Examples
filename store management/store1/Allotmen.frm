VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   25
      Top             =   2520
      Width           =   255
   End
   Begin VB.TextBox txtIsqty 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5640
      TabIndex        =   15
      Text            =   " "
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtItem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2160
      TabIndex        =   13
      Text            =   " "
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtRsqty 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Text            =   " "
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtEname 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      TabIndex        =   12
      Text            =   " "
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox txtDept 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5640
      TabIndex        =   11
      Text            =   " "
      Top             =   1680
      Width           =   1815
   End
   Begin VB.ComboBox cmbItemid 
      Height          =   315
      Left            =   2160
      TabIndex        =   1
      Text            =   " "
      Top             =   1920
      Width           =   1815
   End
   Begin VB.TextBox txtRegno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5640
      TabIndex        =   10
      Text            =   " "
      Top             =   720
      Width           =   1815
   End
   Begin VB.TextBox txtRemark 
      Appearance      =   0  'Flat
      DataField       =   " "
      Height          =   405
      Left            =   2160
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2400
      Width           =   3015
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   6
      Top             =   4920
      Width           =   2295
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   495
      Left            =   10440
      TabIndex        =   5
      Top             =   4920
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mhIssue 
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   3413
      _Version        =   393216
      BackColor       =   16777088
      Cols            =   8
      FixedCols       =   0
      BackColorFixed  =   16744703
      BackColorBkg    =   16777088
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mhBalance 
      Height          =   2055
      Left            =   7560
      TabIndex        =   3
      Top             =   720
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   3625
      _Version        =   393216
      BackColor       =   16777088
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   16744703
      BackColorSel    =   16744703
      BackColorBkg    =   16777088
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mhIndent 
      Height          =   2295
      Left            =   240
      TabIndex        =   2
      Top             =   5640
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   4048
      _Version        =   393216
      BackColor       =   16777088
      BackColorFixed  =   16744703
      BackColorBkg    =   16777088
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker dtpIdate 
      Height          =   375
      Left            =   5640
      TabIndex        =   14
      Top             =   1200
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   24772611
      CurrentDate     =   37364
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000040&
      BorderWidth     =   5
      X1              =   0
      X2              =   11880
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ALLOTMENT FORM"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   615
      Left            =   3480
      TabIndex        =   27
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Transfer"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   255
      Left            =   5760
      TabIndex        =   26
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00400040&
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   255
      Left            =   4080
      TabIndex        =   23
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Quantity"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   255
      Left            =   4080
      TabIndex        =   22
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Requested Quantity "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Employ Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H00400040&
      Height          =   255
      Left            =   4080
      TabIndex        =   19
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Sr. No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label Label10 
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
      ForeColor       =   &H00400040&
      Height          =   255
      Left            =   4080
      TabIndex        =   17
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Remark"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   1695
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
        
Option Explicit
Dim Myear As String
Dim i As Integer, j As Integer, Bal As Integer
Dim sans As VbMsgBoxResult
Dim flag As Boolean
Dim Nflag As Boolean

'Private Sub cmbItemid_GotFocus()
'cmbItemid = ""
'End Sub

Private Sub cmbItemid_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub
Private Sub cmbItemid_LostFocus()
If Nflag = True Then
MsgBox mhIndent.TextMatrix(1, 4) & " " & "Not Available", vbInformation, "Allotment"
Nflag = False
End If
'If cmbItemid = "" And Me.ActiveControl <> cmdNext Then
If Me.ActiveControl <> cmdNext Then
'MsgBox "Select Item -ID From The List", vbInformation, "Allotment"
'cmbItemid.SetFocus
'cmbItemid_GotFocus
'Else
Set rs = cn.Execute("select Item.MinOrdQty,ItemStock.balance From Item,ItemStock where Item.itemname = '" & txtItem & "' and ItemStock.itemid = " & Val(cmbItemid) & " ")
If Not rs.EOF Then
If rs.Fields("MinOrdQty") >= rs.Fields("balance") And Me.ActiveControl <> cmdNext And Me.ActiveControl <> cmdExit Then
MsgBox "Item Is Now On Minimum Order Level", vbExclamation, "Warning"
End If
End If
End If
txtIsqty.SetFocus
cmdNext.Enabled = True
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub


Private Sub cmdNext_Click()
If j = mhIndent.Rows Then
cmdOk_Click
cmdNext.Enabled = False
End If
If j < mhIndent.Rows Then
cmdOk_Click
'flag = False
Set rs = cn.Execute("select empname,dept from Employ where empcode = " & mhIndent.TextMatrix(j, 3) & " ")
txtEname = rs.Fields("empname")
txtDept = rs.Fields("dept")
txtItem = mhIndent.TextMatrix(j, 4)
txtRegno = mhIndent.TextMatrix(j, 1)
txtRsqty = mhIndent.TextMatrix(j, 6)
Set rs1 = cn.Execute("select stocksno,ItemName,Breif,Balance from ItemStock where ItemName = '" & mhIndent.TextMatrix(j, 4) & "' ")
If rs1.EOF Then
cmbItemid.Clear
mhBalance.Clear
MsgBox mhIndent.TextMatrix(j, 4) & " " & "Not Available", vbInformation, "Allotment"
'flag = True
Else
cmbItemid.Clear
cmbItemid.Text = rs1.Fields("stocksno")
While Not rs1.EOF
cmbItemid.AddItem rs1.Fields("stocksno")
rs1.MoveNext
Wend
Set mhBalance.DataSource = rs1
End If
j = j + 1
cmbItemid.SetFocus
End If
cmdNext.Enabled = False
End Sub

Private Sub cmdOk_Click()
If flag = False Then
mhIssue.TextMatrix(i, 0) = Val(cmbItemid)
mhIssue.TextMatrix(i, 1) = txtItem
Set rs = cn.Execute("select breif from ItemStock where stocksno = " & Val(cmbItemid) & " ")
If Not rs.EOF Then
mhIssue.TextMatrix(i, 2) = rs.Fields("breif")
End If
mhIssue.TextMatrix(i, 4) = Format(dtpIdate.Value, "dd-MMM-yy")
mhIssue.TextMatrix(i, 3) = Val(txtIsqty)
mhIssue.TextMatrix(i, 5) = txtRegno
mhIssue.TextMatrix(i, 6) = Me.Check1.Value
mhIssue.TextMatrix(i, 7) = txtRemark
mhIssue.Rows = mhIssue.Rows + 1
If Val(txtIsqty) <> 0 Then
Set rs = cn.Execute("select balance from ItemStock where stocksno = " & Val(cmbItemid.Text) & " ")
cn.Execute ("update ItemStock set balance = " & rs.Fields("balance") - txtIsqty & " where stocksno = " & Val(cmbItemid) & " ")
End If
i = i + 1
Call Clear
End If
cmdOk.Enabled = False
End Sub

Private Sub cmdSave_Click()

If mhIssue.Rows > 1 Then
sans = MsgBox("Are You Sure About Entered Information ", vbYesNo + vbQuestion, "Allotment")

If sans = vbNo Then
Me.Show
End If

If sans = vbYes Then
 i = 1
 While Len(mhIssue.TextMatrix(i, 0)) <> 0
 If mhIssue.TextMatrix(i, 3) = 0 Then
 sql = "insert into ItemIssue(rno,remark) values(" & mhIssue.TextMatrix(i, 5) & ",'" & mhIssue.TextMatrix(i, 7) & "')"
  Else
  If mhIssue.TextMatrix(i, 6) = 0 Then
  sql = "insert into ItemIssue(rno,stocksno,breif_i,issue_date,qty_issue,issued,remark ) values(" & mhIssue.TextMatrix(i, 5) & "," & mhIssue.TextMatrix(i, 0) & ",'" & mhIssue.TextMatrix(i, 2) & "',# " & Format(mhIssue.TextMatrix(i, 4), "dd-MMM-yy") & " # ," & mhIssue.TextMatrix(i, 3) & ",True,'" & mhIssue.TextMatrix(i, 7) & "') "
  Else
  sql = "Insert into ItemIssue(rno,stocksno,breif_i,issue_date,qty_issue,transfer,remark) values(" & mhIssue.TextMatrix(i, 5) & "," & mhIssue.TextMatrix(i, 0) & ",'" & mhIssue.TextMatrix(i, 2) & "',# " & Format(mhIssue.TextMatrix(i, 4), "dd-MMM-yy") & " # ," & mhIssue.TextMatrix(i, 3) & " ,True ,'" & mhIssue.TextMatrix(i, 7) & "')  "
  End If
  'Set rs = cn.Execute("select balance from ItemStock where stocksno = " & mhIssue.TextMatrix(i, 0) & " ")
  'cn.Execute ("update ItemStock set balance = " & rs.Fields("balance") - mhIssue.TextMatrix(i, 3) & " where stocksno =" & mhIssue.TextMatrix(i, 0) & "  ")
 End If
 cn.Execute sql
 cn.Execute ("update ItemRequest set consider = Yes  where rno = " & mhIssue.TextMatrix(i, 5) & "  ")
 i = i + 1
 Wend
 mhIssue.Clear
 Call GridSet
End If
End If
End Sub

'Private Sub cmdSok_Click()
'Select Case Strip1.SelectedItem.Index
'Case 2:
'sql = "select ItemIssue.regno as RegnNo,format(Indent.request_date,'dd-MMM-yy') as RequestDate,Indent.empcode,ItemIssue.ItemName,ItemIssue.breif as BreifDescription,ItemIssue.qty_request as Quantity   from Indent,ItemIssue where ItemIssue.itemid = 0 and  ItemIssue.qty_issue = -1  and  val(itemissue.regno) = indent.regno and Indent.request_date = #" & Format(dtpSdate.Value, "dd-MMM-yy") & "# "
'Set rs = cn.Execute(sql)
'Set mhIndent.DataSource = rs
'Case 3:
'sql = "select ItemIssue.regno as RegnNo,format(Indent.request_date,'dd-MMM-yy') as RequestDate,Indent.empcode,ItemIssue.ItemName,ItemIssue.breif as BreifDescription,ItemIssue.qty_request as Quantity   from Indent,ItemIssue where ItemIssue.itemid = 0 and  ItemIssue.qty_issue = -1  and  val(itemissue.regno) = indent.regno and  Month(Indent.request_date) = '" & Month(dtpSdate.Value) & " ' "
'Set rs = cn.Execute(sql)
'Set mhIndent.DataSource = rs
'Case 4:
'sql = "select ItemIssue.regno as RegnNo,format(Indent.request_date,'dd-MMM-yy') as RequestDate,Indent.empcode,ItemIssue.ItemName,ItemIssue.breif as BreifDescription,ItemIssue.qty_request as Quantity   from Indent,ItemIssue where ItemIssue.itemid = 0 and  ItemIssue.qty_issue = -1  and  val(itemissue.regno) = indent.regno and  Year(Indent.request_date) = '" & Year(dtpSdate.Value) & " ' "
'Set rs = cn.Execute(sql)
'Set mhIndent.DataSource = rs
'End Select
'End Sub


Private Sub Form_Load()
Call Connect
Call GridSet
cmdNext.Enabled = False
i = 1
j = 2
flag = False
dtpIdate.Value = Format(Now, "dd-MMM-yy")
sql = "select ItemRequest.rno ,format(Indent.request_date,'dd-MMM-yy') as RequestDate,Indent.empcode,ItemRequest.ItemName,ItemRequest.breif_r as BreifDescription,ItemRequest.qty_request as Quantity   from Indent,ItemRequest where ItemRequest.consider = false and  ItemRequest.indentno = indent.indentno "
Set rs = cn.Execute(sql)
Set mhIndent.DataSource = rs
If mhIndent.Rows > 1 Then
Set rs = cn.Execute("select empname,dept from Employ where empcode = " & mhIndent.TextMatrix(1, 3) & " ")
txtEname = rs.Fields("empname")
txtDept = rs.Fields("dept")
Set rs1 = cn.Execute("select stocksno,ItemName,Breif,Balance from ItemStock where itemname = '" & mhIndent.TextMatrix(1, 4) & "' and balance > 0  ")
If Not rs1.EOF Then
cmbItemid.Text = rs1.Fields("stocksno")
While Not rs1.EOF
cmbItemid.AddItem rs1.Fields("stocksno")
rs1.MoveNext
Wend
Set mhBalance.DataSource = rs1
Else
cmbItemid.Clear
mhBalance.Clear
Nflag = True
'MsgBox mhIndent.TextMatrix(1, 4) & " " & "Not Available", vbInformation, "Allotment"
End If
txtItem = mhIndent.TextMatrix(1, 4)
txtRegno = mhIndent.TextMatrix(1, 1)
txtRsqty = mhIndent.TextMatrix(1, 6)
End If
End Sub

'Private Sub Strip1_Click()
'Select Case Strip1.SelectedItem.Index
'Case 1:
'Label4.Visible = False
'dtpSdate.Visible = False
'cmdSok.Visible = False
'sql = "select ItemIssue.regno as RegnNo,format(Indent.request_date,'dd-MMM-yy') as RequestDate,Indent.empcode,ItemIssue.ItemName,ItemIssue.breif as BreifDescription,ItemIssue.qty_request as Quantity   from Indent,ItemIssue where ItemIssue.itemid = 0 and  ItemIssue.qty_issue = -1  and  val(itemissue.regno) = indent.regno"
'Set rs = cn.Execute(sql)
'Set mhIndent.DataSource = rs
'Case 2:
'Label4.Visible = True
'Label4.Caption = "Select Date From Date Picker"
'dtpSdate.Visible = True
'cmdSok.Visible = True
'Case 3:
'Label4.Visible = True
'Label4.Caption = "Select Month And Year From Date Picker"
'dtpSdate.Visible = True
'cmdSok.Visible = True
'Case 4:
'Label4.Visible = True
'Label4.Caption = "Select Year From Date Picker"
'dtpSdate.Visible = True
'cmdSok.Visible = True
'End Select
'End Sub
Public Sub GridSet()
mhIndent.ColWidth(0) = 700
mhIndent.ColWidth(1) = 1200
mhIndent.ColWidth(2) = 1500
mhIndent.ColWidth(3) = 1250
mhIndent.ColWidth(4) = 2400
mhIndent.ColWidth(5) = 3200
mhIndent.ColWidth(6) = 1000
mhBalance.ColWidth(0) = 500
mhBalance.ColWidth(1) = 900
mhBalance.ColWidth(2) = 2100
mhBalance.ColWidth(3) = 675
mhIssue.ColWidth(0) = 800
mhIssue.ColWidth(1) = 1400
mhIssue.ColWidth(2) = 2900
mhIssue.ColWidth(3) = 1500
mhIssue.ColWidth(4) = 1200
mhIssue.ColWidth(5) = 900
mhIssue.TextMatrix(0, 0) = "Item-Id"
mhIssue.TextMatrix(0, 1) = "Item Name"
mhIssue.TextMatrix(0, 2) = "Breif Description "
mhIssue.TextMatrix(0, 3) = "Issue Quantity"
mhIssue.TextMatrix(0, 4) = "Issue Date"
mhIssue.TextMatrix(0, 5) = "Regn. No."
mhIssue.TextMatrix(0, 6) = "Transfer"
mhIssue.TextMatrix(0, 7) = "Remark"
End Sub



Private Sub txtIsqty_KeyPress(KeyAscii As Integer)
If KeyAscii > 47 And KeyAscii < 58 Then
If cmbItemid <> "" Then
cmdOk.Enabled = True
End If
Else
KeyAscii = 0
End If
End Sub

Private Sub txtIsqty_LostFocus()
If mhBalance.TextMatrix(0, 1) <> "" Then
Set rs = cn.Execute("select balance from itemstock where stocksno = " & Val(cmbItemid) & "  ")
Bal = rs.Fields("balance")
If Val(txtIsqty) > Bal Then
MsgBox "Not Available", vbInformation, "Allotment"
txtIsqty = ""
txtIsqty.SetFocus
End If
End If
End Sub
Public Sub Clear()
txtItem = ""
txtIsqty = ""
txtRsqty = ""
txtDept = ""
txtEname = ""
cmbItemid.Clear
txtRemark = ""
End Sub
Private Sub txtRsqty_GotFocus()
cmbItemid.SetFocus
End Sub

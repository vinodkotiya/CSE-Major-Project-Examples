VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form11 
   BackColor       =   &H00404080&
   Caption         =   "Form11"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form11"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbItem 
      Height          =   315
      Left            =   3360
      TabIndex        =   9
      Text            =   " "
      Top             =   2160
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ALL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0FFFF&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0FFFF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtEmpcode 
      Height          =   285
      Left            =   3480
      TabIndex        =   3
      Text            =   " "
      Top             =   2160
      Width           =   2055
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Employee  Code"
            Key             =   "sbc"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Department"
            Key             =   "sdept"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Item Name"
            Key             =   "sbin"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Date Of Issue"
            Key             =   "sbd"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Month Of Issue"
            Key             =   "sbm"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Between  Dates Of Issue"
            Key             =   "sby"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mhIndent 
      Height          =   3495
      Left            =   120
      TabIndex        =   4
      Top             =   4680
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   6165
      _Version        =   393216
      BackColor       =   16777152
      BackColorFixed  =   8421631
      BackColorBkg    =   12648384
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker dtpIndent 
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   24444931
      CurrentDate     =   37379
   End
   Begin MSComCtl2.DTPicker dtpIndent1 
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   24444931
      CurrentDate     =   37379
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
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
      Index           =   0
      Left            =   5400
      TabIndex        =   11
      Top             =   2160
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   5
      X1              =   0
      X2              =   11880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Employee Code"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ALLOTMENT SEARCH FORM"
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
      Left            =   2280
      TabIndex        =   0
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAll_Click()
sql = "SELECT format(ItemIssue.issue_date,'dd-MMM-yy'),ItemIssue.rno, ItemRequest.itemname,ItemIssue.breif_i,ItemIssue.stocksno,ItemIssue.qty_issue,ItemIssue.issued,ItemIssue.transfer,ItemIssue.remark,Indent.empcode, Employ.empname,Employ.dept"
sql = sql & "  FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.qty_issue > 0  order by ItemIssue.issue_date;"
Set rs = cn.Execute(sql)
Set mhIndent.DataSource = rs
Call Setting
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Select Case Tab1.SelectedItem.Index
Case 1:
sql = "SELECT format(ItemIssue.issue_date,'dd-MMM-yy'),ItemIssue.rno, ItemRequest.itemname,ItemIssue.breif_i,ItemIssue.stocksno,ItemIssue.qty_issue,ItemIssue.issued,ItemIssue.transfer,ItemIssue.remark,Indent.empcode, Employ.empname,Employ.dept"
sql = sql & "  FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.qty_issue > 0  and ItemIssue.issued = True and Indent.empcode = " & Val(txtEmpcode) & " order by ItemIssue.issue_date;"
Case 2:
sql = "SELECT format(ItemIssue.issue_date,'dd-MMM-yy'),ItemIssue.rno, ItemRequest.itemname,ItemIssue.breif_i,ItemIssue.stocksno,ItemIssue.qty_issue,ItemIssue.issued,ItemIssue.transfer,ItemIssue.remark,Indent.empcode, Employ.empname,Employ.dept"
sql = sql & "  FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.qty_issue > 0  and ItemIssue.transfer = True and employ.dept = '" & Me.cmbItem & "' order by ItemIssue.issue_date;"
Case 4:
sql = "SELECT format(ItemIssue.issue_date,'dd-MMM-yy'),ItemIssue.rno, ItemRequest.itemname,ItemIssue.breif_i,ItemIssue.stocksno,ItemIssue.qty_issue,ItemIssue.issued,ItemIssue.transfer,ItemIssue.remark,Indent.empcode, Employ.empname,Employ.dept"
sql = sql & "  FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.qty_issue > 0  and ItemIssue.issue_date = #" & Format(dtpIndent.Value, "dd-MMM-yy") & "#   order by ItemIssue.issue_date;"
Case 3:
sql = "SELECT format(ItemIssue.issue_date,'dd-MMM-yy'),ItemIssue.rno, ItemRequest.itemname,ItemIssue.breif_i,ItemIssue.stocksno,ItemIssue.qty_issue,ItemIssue.issued,ItemIssue.transfer,ItemIssue.remark,Indent.empcode, Employ.empname,Employ.dept"
sql = sql & "  FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.qty_issue > 0  and ItemREquest.itemname = '" & Me.cmbItem & "' order by ItemIssue.issue_date;"
Case 5:
sql = "SELECT format(ItemIssue.issue_date,'dd-MMM-yy'),ItemIssue.rno, ItemRequest.itemname,ItemIssue.breif_i,ItemIssue.stocksno,ItemIssue.qty_issue,ItemIssue.issued,ItemIssue.transfer,ItemIssue.remark,Indent.empcode, Employ.empname,Employ.dept"
sql = sql & "  FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.qty_issue > 0  and Month(ItemIssue.issue_date) = '" & Month(dtpIndent) & "' and year(ItemIssue.issue_date) = '" & Year(dtpIndent) & "'   order by ItemIssue.issue_date;"
Case 6:
If dtpIndent.Value = dtpIndent1.Value Then
sql = "SELECT format(ItemIssue.issue_date,'dd-MMM-yy'),ItemIssue.rno, ItemRequest.itemname,ItemIssue.breif_i,ItemIssue.stocksno,ItemIssue.qty_issue,ItemIssue.issued,ItemIssue.transfer,ItemIssue.remark,Indent.empcode, Employ.empname,Employ.dept"
sql = sql & "  FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.qty_issue > 0  and ItemIssue.issue_date = #" & Format(dtpIndent.Value, "dd-MMM-yy") & "#   order by ItemIssue.issue_date;"
Else
sql = "SELECT format(ItemIssue.issue_date,'dd-MMM-yy'),ItemIssue.rno, ItemRequest.itemname,ItemIssue.breif_i,ItemIssue.stocksno,ItemIssue.qty_issue,ItemIssue.issued,ItemIssue.transfer,ItemIssue.remark,Indent.empcode, Employ.empname,Employ.dept"
sql = sql & "  FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.qty_issue > 0  and ItemIssue.issue_date >= #" & dtpIndent.Value & "# and ItemIssue.issue_date <= #" & dtpIndent1.Value & "#   order by ItemIssue.issue_date;"
End If
End Select
Set rs = cn.Execute(sql)
Set mhIndent.DataSource = rs
Call Setting
End Sub

Private Sub Form_Load()
Call Connect
sql = "SELECT format(ItemIssue.issue_date,'dd-MMM-yy'),ItemIssue.rno, ItemRequest.itemname,ItemIssue.breif_i,ItemIssue.stocksno,ItemIssue.qty_issue,ItemIssue.issued,ItemIssue.transfer,ItemIssue.remark,Indent.empcode, Employ.empname,Employ.dept"
sql = sql & "  FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.qty_issue > 0  order by ItemIssue.issue_date;"
Set rs = cn.Execute(sql)
Set mhIndent.DataSource = rs
Call Setting
cmbItem.Visible = False
dtpIndent.Visible = False
Label3(0).Visible = False
dtpIndent1.Visible = False
End Sub
Public Sub Setting()
mhIndent.ColWidth(0) = 300
mhIndent.ColWidth(1) = 1200
mhIndent.ColWidth(2) = 1200
mhIndent.ColWidth(3) = 1500
mhIndent.ColWidth(4) = 2000
mhIndent.ColWidth(5) = 700
mhIndent.ColWidth(6) = 700
mhIndent.ColWidth(7) = 700
mhIndent.ColWidth(8) = 700
mhIndent.ColWidth(9) = 2000
mhIndent.ColWidth(10) = 700
mhIndent.ColWidth(11) = 1600
mhIndent.ColWidth(12) = 1400
mhIndent.TextMatrix(0, 1) = "IssueDate"
mhIndent.TextMatrix(0, 2) = "ReqItmNo"
mhIndent.TextMatrix(0, 3) = "ItemName"
mhIndent.TextMatrix(0, 4) = "Breif"
mhIndent.TextMatrix(0, 5) = "StkSno"
mhIndent.TextMatrix(0, 6) = "IssueQty"
mhIndent.TextMatrix(0, 7) = "Issued"
mhIndent.TextMatrix(0, 8) = "Transfered"
mhIndent.TextMatrix(0, 9) = "Remark"
mhIndent.TextMatrix(0, 10) = "Ecode"
mhIndent.TextMatrix(0, 11) = "Ename"
mhIndent.TextMatrix(0, 12) = "Department"
End Sub

Private Sub Tab1_Click()
Select Case Tab1.SelectedItem.Index
Case 1:
Label2.Caption = "Enter Employee Code"
txtEmpcode.Visible = True
cmbItem.Visible = False
dtpIndent.Visible = False
Label3(0).Visible = False
dtpIndent1.Visible = False
Case 2:
Label2.Caption = "Select Item Name"
txtEmpcode.Visible = False
cmbItem.Visible = True
dtpIndent.Visible = False
Label3(0).Visible = False
dtpIndent1.Visible = False
Set rs = cn.Execute("select distinct(dept) from Employ")
cmbItem.Clear
While Not rs.EOF
cmbItem.AddItem rs.Fields("dept")
rs.MoveNext
Wend
Case 3:
Label2.Caption = "Select Item Name"
txtEmpcode.Visible = False
cmbItem.Visible = True
dtpIndent.Visible = False
Label3(0).Visible = False
dtpIndent1.Visible = False
Set rs = cn.Execute("select itemname from Item")
cmbItem.Clear
While Not rs.EOF
cmbItem.AddItem rs.Fields("itemname")
rs.MoveNext
Wend
Case 4:
Label2.Caption = "Select Date"
txtEmpcode.Visible = False
cmbItem.Visible = False
dtpIndent.Visible = True
Label3(0).Visible = False
dtpIndent1.Visible = False
Case 5:
Label2.Caption = " Select Month & Year"
txtEmpcode.Visible = False
cmbItem.Visible = False
dtpIndent.Visible = True
Label3(0).Visible = False
dtpIndent1.Visible = False
Case 6:
Label2.Caption = "Select Date From:"
txtEmpcode.Visible = False
cmbItem.Visible = False
dtpIndent.Visible = True
Label3(0).Visible = True
dtpIndent1.Visible = True
End Select
End Sub


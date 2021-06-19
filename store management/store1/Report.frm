VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form4"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4695
      Left            =   2280
      TabIndex        =   10
      Top             =   1920
      Width           =   6615
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00C0C0C0&
         Caption         =   "PRINT"
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
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CommandButton cmdView 
         BackColor       =   &H00C0C0C0&
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
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00C0C0C0&
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
         Height          =   495
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4080
         Width           =   1575
      End
      Begin VB.OptionButton optReport 
         BackColor       =   &H00C0C0C0&
         Caption         =   "LIST OF ITEMS ON MINIMUM ORDER LEVEL"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   465
         Index           =   0
         Left            =   390
         TabIndex        =   18
         Top             =   600
         Width           =   4215
      End
      Begin VB.OptionButton optReport 
         BackColor       =   &H00C0C0C0&
         Caption         =   "STOCK POSITION OF SELECTED ITEM"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   495
         Index           =   1
         Left            =   390
         TabIndex        =   17
         Top             =   915
         Width           =   4215
      End
      Begin VB.OptionButton optReport 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RECORD OF ISSUED ITEMS ISSUE BETWEEN DATES"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   375
         Index           =   2
         Left            =   390
         TabIndex        =   16
         Top             =   1335
         Width           =   4935
      End
      Begin VB.OptionButton optReport 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RECORD OF TRNSFRED ITEMS TRANSFER BETWEEN DATES"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   495
         Index           =   3
         Left            =   390
         TabIndex        =   15
         Top             =   1650
         Width           =   5535
      End
      Begin VB.OptionButton optReport 
         BackColor       =   &H00C0C0C0&
         Caption         =   "INDIVIDUAL EMPLOYEE RECORD OF ITEMS ISSUED"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   495
         Index           =   4
         Left            =   360
         TabIndex        =   14
         Top             =   2070
         Width           =   4935
      End
      Begin VB.OptionButton optReport 
         BackColor       =   &H00C0C0C0&
         Caption         =   "INDIVIDUAL RECORD OF  DEPARTMENT OF ITEMS TRANSFERED "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   375
         Index           =   5
         Left            =   390
         TabIndex        =   13
         Top             =   2520
         Width           =   6015
      End
      Begin VB.OptionButton optReport 
         BackColor       =   &H00C0C0C0&
         Caption         =   "LIST OF SUPPLIER FOR INDIVIDUAL ITEM"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   495
         Index           =   6
         Left            =   390
         TabIndex        =   12
         Top             =   2925
         Width           =   4335
      End
      Begin VB.OptionButton optReport 
         BackColor       =   &H00C0C0C0&
         Caption         =   "DETAIL OF ITEMS RETURN"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400040&
         Height          =   495
         Index           =   7
         Left            =   390
         TabIndex        =   11
         Top             =   3360
         Width           =   3255
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Parameters For Generating Report"
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
         Height          =   375
         Left            =   360
         TabIndex        =   29
         Top             =   0
         Width           =   5775
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   6600
         Y1              =   480
         Y2              =   480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   8040
      Width           =   1695
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "PRINT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3000
         Width           =   1695
      End
      Begin VB.CommandButton cmd6 
         BackColor       =   &H00C0C0C0&
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
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3000
         Width           =   1695
      End
      Begin VB.CommandButton cmd5 
         BackColor       =   &H00C0C0C0&
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
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3000
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtp2 
         Height          =   375
         Left            =   5160
         TabIndex        =   23
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yy"
         Format          =   24576003
         CurrentDate     =   37410
      End
      Begin MSComCtl2.DTPicker dtp1 
         Height          =   375
         Left            =   2880
         TabIndex        =   22
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MMM-yy"
         Format          =   24576003
         CurrentDate     =   37410
      End
      Begin VB.ComboBox cmbItem 
         Height          =   315
         Left            =   3360
         TabIndex        =   8
         Text            =   " "
         Top             =   720
         Width           =   3255
      End
      Begin VB.CommandButton cmd4 
         BackColor       =   &H00C0C0C0&
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
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   3360
         TabIndex        =   6
         Text            =   " "
         Top             =   720
         Width           =   2055
      End
      Begin VB.CommandButton cmd1 
         BackColor       =   &H00C0C0C0&
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
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3000
         Width           =   1695
      End
      Begin VB.CommandButton cmd2 
         BackColor       =   &H00C0C0C0&
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
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3000
         Width           =   1695
      End
      Begin VB.CommandButton cmd3 
         BackColor       =   &H00C0C0C0&
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
         Height          =   615
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3000
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0C0&
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
         Height          =   615
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   -120
         X2              =   6480
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label3 
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
         ForeColor       =   &H00400040&
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   120
         Width           =   5775
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H00400040&
         Height          =   255
         Left            =   4560
         TabIndex        =   27
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00400040&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   2655
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00400040&
      BorderWidth     =   5
      Height          =   5175
      Left            =   2160
      Top             =   1680
      Width           =   6975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "COMPUTER EQUIPMENT PROFILE SYSTEM"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbReport2_Click()
sql = "select stocksno,itemname,breif,balance from ItemStock where itemname = '" & cmbItem & "' "
Set rs = cn.Execute(sql)
Set Report1.DataSource = rs
Report1.Sections("section4").Controls("label2").Caption = "PROFILE OF SELECTED ITEM IN THE STOCK"
Report1.Show
End Sub

Private Sub cmd1_Click()
sql = "select stocksno,itemname,breif,balance from ItemStock where itemname = '" & cmbItem & "' "
Set rs = cn.Execute(sql)
Set Report1.DataSource = rs
Report1.Sections("section4").Controls("label2").Caption = "PROFILE OF SELECTED ITEM IN THE STOCK"
Report1.Show
End Sub

Private Sub cmd2_Click()
If dtp1.Value = dtp2.Value Then
sql = "SELECT ItemIssue.rno,ItemRequest.itemname,ItemIssue.breif_i,ItemIssue.qty_issue,format(ItemIssue.issue_date,'dd-mm-yy') as Idate,ItemIssue.remark,Indent.empcode,Employ.empname,Employ.dept"
sql = sql & "  FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.issued = true and ItemIssue.issue_date = #" & Format(dtp1.Value, "dd-MMM-yy") & "#; "
Else
sql = "SELECT ItemIssue.rno,ItemRequest.itemname,ItemIssue.breif_i,ItemIssue.qty_issue,format(ItemIssue.issue_date,'dd-mm-yy')as Idate,ItemIssue.remark,Indent.empcode,Employ.empname,Employ.dept"
sql = sql & "  FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.issued = true and ItemIssue.issue_date >= #" & Format(dtp1.Value, "dd-MMM-yy") & "# and ItemIssue.issue_date <= #" & Format(dtp2.Value, "dd-MMM-yy") & "#; "
End If
Set rs = cn.Execute(sql)
Set Report2.DataSource = rs
'Report2.Sections("section4").Controls("label2").Caption = "PROFILE OF SELECTED ITEM IN THE STOCK"
Report2.Show
End Sub

Private Sub cmd3_Click()
If dtp1.Value = dtp2.Value Then
sql = "SELECT ItemIssue.rno,ItemRequest.itemname,ItemIssue.breif_i,ItemIssue.qty_issue,format(ItemIssue.issue_date,'dd-MMM-yy')as Idate,ItemIssue.remark,Indent.empcode,Employ.dept"
sql = sql & "  FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.transfer = true and ItemIssue.issue_date = #" & Format(dtp1.Value, "dd-MMM-yy") & "# ; "
Else
sql = "SELECT ItemIssue.rno,ItemRequest.itemname,ItemIssue.breif_i,ItemIssue.qty_issue,format(ItemIssue.issue_date,'dd-MMM-yy')as Idate,ItemIssue.remark,Indent.empcode,Employ.dept"
sql = sql & "  FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.transfer = true and ItemIssue.issue_date >= #" & Format(dtp1.Value, "dd-MMM-yy") & "# and ItemIssue.issue_date <= #" & Format(dtp2.Value, "dd-MMM-yy") & "#; "
End If
Set rs = cn.Execute(sql)
Set Report3.DataSource = rs
'Report2.Sections("section4").Controls("label2").Caption = "PROFILE OF SELECTED ITEM IN THE STOCK"
Report3.Show
End Sub

Private Sub cmd4_Click()
Set rs1 = cn.Execute("select * from employ where empcode = " & Val(Text1) & " ")
Report4.Sections("section2").Controls("lblEcode").Caption = rs1.Fields("empcode")
Report4.Sections("section2").Controls("lblEname").Caption = rs1.Fields("empname")
Report4.Sections("section2").Controls("lblDsn").Caption = rs1.Fields("dsn")
Report4.Sections("section2").Controls("lblDept").Caption = rs1.Fields("dept")
sql = "SELECT ItemIssue.rno,ItemRequest.itemname,ItemIssue.breif_i,ItemIssue.qty_issue,format(ItemIssue.issue_date,'dd-MMM-yy')as Idate,ItemIssue.remark,Indent.empcode"
sql = sql & "  FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.issued = true and Indent.empcode = " & Val(Text1) & "; "
Set rs = cn.Execute(sql)
Set Report4.DataSource = rs
Report4.Show
End Sub

Private Sub cmd5_Click()
sql = "SELECT ItemIssue.rno,ItemRequest.itemname,ItemIssue.breif_i,ItemIssue.qty_issue,format(ItemIssue.issue_date,'dd-MMM-yy')as Idate,ItemIssue.remark,Indent.empcode,Employ.dept"
sql = sql & "  FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.transfer = true and Employ.dept = '" & Me.cmbItem.Text & "'; "
Set rs = cn.Execute(sql)
'Set Report9.DataSource = rs
'Report9.Sections("section2").Controls("Label9").Caption = cmbItem.Text
'Report9.Show
End Sub

Private Sub cmd6_Click()
Set rs1 = cn.Execute("select * from Item where itemname = '" & cmbItem & "' ")
Report6.Sections("section2").Controls("lblItemid").Caption = rs1.Fields("itemid")
Report6.Sections("section2").Controls("lblIname").Caption = rs1.Fields("itemname")
sql = "select * from Supplier where supplier_id in(select supplier_id from ItemStock where itemname = '" & cmbItem & "' )"
Set rs = cn.Execute(sql)
Set Report6.DataSource = rs
Report6.Show
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub



Private Sub cmdView_Click()
If optReport(0).Value = True Then
sql = "select Item.itemid,item.itemname,item.MinOrdQty,ItemStock.stocksno,ItemStock.balance,ItemStock.breif from Item,ItemStock where ItemStock.balance <= Item.MinOrdQty and ItemStock.itemid = Item.itemid "
Set rs = cn.Execute(sql)
Set Report1.DataSource = rs
Report1.Show
End If
If optReport(7).Value = True Then
'sql = "SELECT ItemReturn.rno,ItemRequest.itemname,ItemIssue.breif_i,ItemIssue.qty_issue,format(ItemIssue.issue_date ,'dd-MMM-yy'),ItemReturn.return_qty,format(ItemReturn.return_date,'dd-MMM-yy'),ItemReturn.empcode,ItemReturn.dept "
'sql = sql & " FROM (ItemReturn INNER JOIN ItemIssue ON ItemReturn.rno = ItemIssue.rno) INNER JOIN ItemRequest ON ItemReturn.rno = ItemRequest.rno; "
sql = "SELECT ItemReturn.*, ItemIssue.*, ItemRequest.itemname"
sql = sql & "  FROM (ItemReturn INNER JOIN ItemIssue ON ItemReturn.rno = ItemIssue.rno) INNER JOIN ItemRequest ON ItemReturn.rno = ItemRequest.rno;"
Set rs = cn.Execute(sql)
Set Return1.DataSource = rs
Return1.Show
End If
End Sub

Private Sub Command1_Click()
Set rs1 = cn.Execute("select * from employ where empcode = " & Val(Text1) & " ")
Report4.Sections("section2").Controls("lblEcode").Caption = rs1.Fields("empcode")
Report4.Sections("section2").Controls("lblEname").Caption = rs1.Fields("empname")
Report4.Sections("section2").Controls("lblDsn").Caption = rs1.Fields("dsn")
Report4.Sections("section2").Controls("lblDept").Caption = rs1.Fields("dept")
sql = "SELECT ItemIssue.stocksno,ItemIssue.itemname,format(ItemIssue.issue_date,'dd-MMM-yy') as Idate ,ItemIssue.breif_i,ItemIssue.qty_issue,Indent.empcode from ItemIssue,Indent where val(ItemIssue.indentno_sno) = Indent.indentno and Indent.empcode = " & Val(Text1) & " and ItemIssue.qty_issue > 0 and ItemIssue.transfer = false"
Set rs = cn.Execute(sql)
Set Report4.DataSource = rs
Report4.Show
End Sub
Private Sub Command2_Click()
Set rs1 = cn.Execute("select * from Item where itemname = '" & cmbItem & "' ")
Report5.Sections("section2").Controls("lblItemid").Caption = rs1.Fields("itemid")
Report5.Sections("section2").Controls("lblIname").Caption = rs1.Fields("itemname")
sql = "SELECT Indent.empcode,Employ.*,ItemIssue.stocksno,ItemIssue.breif_i from Indent,Employ,ItemIssue where Indent.empcode = Employ.empcode and Indent.indentno = val(ItemIssue.indentno_sno) and ItemIssue.qty_issue >0 and ItemIssue.transfer = false and ItemIssue.itemname = '" & cmbItem & "' "
Set rs = cn.Execute(sql)
Set Report5.DataSource = rs
Report5.Show
End Sub

Private Sub Command3_Click()
Set rs1 = cn.Execute("select * from Item where itemname = '" & cmbItem & "' ")
Report6.Sections("section2").Controls("lblItemid").Caption = rs1.Fields("itemid")
Report6.Sections("section2").Controls("lblIname").Caption = rs1.Fields("itemname")
sql = "select * from Supplier where supplier_id in(select supplier_id from ItemStock where itemname = '" & cmbItem & "' )"
Set rs = cn.Execute(sql)
Set Report6.DataSource = rs
Report6.Show
End Sub

Private Sub Command4_Click()
Frame1.Visible = False
Frame2.Visible = True
Frame2.Height = 4695
Frame2.Width = 6615
Frame2.Top = 1920
Frame2.Left = 2280
End Sub

Private Sub Form_Load()
Call Connect
Set rs = cn.Execute("select itemname from Item")
While Not rs.EOF
cmbItem.AddItem rs.Fields("itemname")
rs.MoveNext
Wend
End Sub

Private Sub optReport_Click(Index As Integer)
Select Case optReport.Item(Index).Index
Case 0:
Frame2.Visible = True
Frame2.Height = 4695
Frame2.Width = 6615
Frame2.Top = 1920
Frame2.Left = 2280
Frame1.Visible = False
cmdPrint.Visible = True
cmdView.Visible = True
Case 1:
Frame1.Visible = True
Frame1.Height = 4695
Frame1.Width = 6615
Frame1.Top = 1920
Frame1.Left = 2280
Me.Label3.Caption = "Stock Report of Selected Item"
Frame2.Visible = False
Label1.Caption = "Item List"
Set rs = cn.Execute("select itemname from Item")
cmbItem.Clear
While Not rs.EOF
cmbItem.AddItem rs.Fields("itemname")
rs.MoveNext
Wend
Label4.Visible = False
cmbItem.Visible = True
Text1.Visible = False
dtp1.Visible = False
dtp2.Visible = False
cmdPrint.Visible = False
cmdView.Visible = False
Call DisButton
cmd1.Visible = True
Case 2:
Frame1.Visible = True
Frame1.Height = 4695
Frame1.Width = 6615
Frame1.Top = 1920
Frame1.Left = 2280
Me.Label3.Caption = "Report on Issued Items Between Dates "
Frame2.Visible = False
Label1.Caption = "Select Date From"
Label4.Visible = True
cmbItem.Visible = False
Text1.Visible = False
dtp1.Visible = True
dtp2.Visible = True
cmdPrint.Visible = False
cmdView.Visible = False
Call DisButton
cmd2.Visible = True
Case 3:
Frame1.Visible = True
Frame1.Height = 4695
Frame1.Width = 6615
Frame1.Top = 1920
Frame1.Left = 2280
Me.Label3.Caption = "Report on Transfered Items Between Dates "
Frame2.Visible = False
Label1.Caption = "Select Date From"
Label4.Visible = True
cmbItem.Visible = False
Text1.Visible = False
dtp1.Visible = True
dtp2.Visible = True
cmdPrint.Visible = False
cmdView.Visible = False
Call DisButton
cmd3.Visible = True
Case 4:
Frame1.Visible = True
Frame1.Height = 4695
Frame1.Width = 6615
Frame1.Top = 1920
Frame1.Left = 2280
Frame2.Visible = False
Me.Label3.Caption = "Report on Issued Items to an Individual Employee "
Label1.Caption = "Enter Employee Code"
Label4.Visible = False
cmbItem.Visible = False
Text1.Visible = True
dtp1.Visible = False
dtp2.Visible = False
cmdPrint.Visible = False
cmdView.Visible = False
Call DisButton
cmd4.Visible = True
Case 5:
Frame1.Visible = True
Frame1.Height = 4695
Frame1.Width = 6615
Frame1.Top = 1920
Frame1.Left = 2280
Me.Label3.Caption = "Report on Transfered Items to an Individual Dept. "
Frame2.Visible = False
Label1.Caption = "Select Department"
Label4.Visible = False
cmbItem.Visible = True
Text1.Visible = False
dtp1.Visible = False
dtp2.Visible = False
Set rs = cn.Execute("select distinct(dept) from Employ")
cmbItem.Clear
While Not rs.EOF
cmbItem.AddItem rs.Fields("dept")
rs.MoveNext
Wend
cmdPrint.Visible = False
cmdView.Visible = False
Call DisButton
cmd5.Visible = True
Case 6:
Frame1.Visible = True
Frame1.Height = 4695
Frame1.Width = 6615
Frame1.Top = 1920
Frame1.Left = 2280
Frame2.Visible = False
Me.Label3.Caption = "List Of Supplier for Supply of Specific Item  "
Label1.Caption = "Item List"
Label4.Visible = False
cmbItem.Visible = True
Text1.Visible = False
dtp1.Visible = False
dtp2.Visible = False
Set rs = cn.Execute("select itemname from Item")
cmbItem.Clear
While Not rs.EOF
cmbItem.AddItem rs.Fields("itemname")
rs.MoveNext
Wend
cmdPrint.Visible = False
cmdView.Visible = False
Call DisButton
cmd6.Visible = True
Case 7:
Frame2.Visible = True
Frame2.Height = 4695
Frame2.Width = 6615
Frame2.Top = 1920
Frame2.Left = 2280
Frame1.Visible = False
cmdPrint.Visible = True
cmdView.Visible = True
End Select
End Sub


Private Sub DisButton()
cmd1.Visible = False
cmd2.Visible = False
cmd3.Visible = False
cmd4.Visible = False
cmd5.Visible = False
cmd6.Visible = False
End Sub


'   Report on issued item
'SELECT ItemIssue.*, ItemRequest.itemname, Indent.empcode, Employ.empname
'FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode;
' Report on supplier list
'SELECT Supplier.*, ItemStock.supplier_id
'FROM Supplier INNER JOIN ItemStock ON Supplier.supplier_id = ItemStock.supplier_id;


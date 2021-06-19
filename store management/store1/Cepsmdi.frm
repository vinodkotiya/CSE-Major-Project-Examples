VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00000040&
   Caption         =   "MAIN FORM"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mnuind 
         Caption         =   "&Indent  "
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnustk 
         Caption         =   "&Stock "
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuallot 
         Caption         =   "&Allotment "
      End
      Begin VB.Menu RET 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRetitm 
         Caption         =   "Return Item"
      End
      Begin VB.Menu mnubar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuemp 
         Caption         =   "&Employ"
      End
      Begin VB.Menu mnub1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSupl 
         Caption         =   "&Supplier"
      End
      Begin VB.Menu mnub2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit "
      End
   End
   Begin VB.Menu mnuEdt 
      Caption         =   "&Edit"
      Begin VB.Menu mnuAdd 
         Caption         =   "&Addition"
         Begin VB.Menu mnuEmp1 
            Caption         =   "Employ"
         End
         Begin VB.Menu S1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSupl1 
            Caption         =   "Supplier"
         End
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpd 
         Caption         =   "Updation"
         Begin VB.Menu mnuInd1 
            Caption         =   "Indent"
         End
         Begin VB.Menu s11 
            Caption         =   "-"
         End
         Begin VB.Menu mnustk1 
            Caption         =   "Stock"
         End
         Begin VB.Menu s3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuallot1 
            Caption         =   "Allotment"
         End
         Begin VB.Menu s 
            Caption         =   "-"
         End
         Begin VB.Menu mnuemp2 
            Caption         =   "Employ"
         End
         Begin VB.Menu s5 
            Caption         =   "-"
         End
         Begin VB.Menu mnuspl2 
            Caption         =   "Supplier"
         End
      End
      Begin VB.Menu s6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDel 
         Caption         =   "Deletion"
         Begin VB.Menu mnuInd3 
            Caption         =   "Indent"
         End
         Begin VB.Menu s7 
            Caption         =   "-"
         End
         Begin VB.Menu mnustk3 
            Caption         =   "Stock"
         End
         Begin VB.Menu S8 
            Caption         =   "-"
         End
         Begin VB.Menu mnuallot3 
            Caption         =   "Allotment"
         End
         Begin VB.Menu S9 
            Caption         =   "-"
         End
         Begin VB.Menu mnuemp3 
            Caption         =   "Employ"
         End
         Begin VB.Menu s10 
            Caption         =   "-"
         End
         Begin VB.Menu mnuspl3 
            Caption         =   "Supplier"
         End
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuind4 
         Caption         =   "Request To Be Consider "
      End
      Begin VB.Menu s12 
         Caption         =   "-"
      End
      Begin VB.Menu mnustk4 
         Caption         =   "Stock"
      End
      Begin VB.Menu s13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuallot4 
         Caption         =   "Allotment"
      End
      Begin VB.Menu S14 
         Caption         =   "-"
      End
      Begin VB.Menu new1 
         Caption         =   "Complete Request"
      End
      Begin VB.Menu nw1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuemp4 
         Caption         =   "Employ"
      End
      Begin VB.Menu S15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuspl4 
         Caption         =   "Supplier"
      End
   End
   Begin VB.Menu mnuAny 
      Caption         =   "&Analysis"
      Begin VB.Menu mnuind5 
         Caption         =   "Indent"
      End
      Begin VB.Menu S16 
         Caption         =   "-"
      End
      Begin VB.Menu mnustk5 
         Caption         =   "Stock"
      End
      Begin VB.Menu S17 
         Caption         =   "-"
      End
      Begin VB.Menu mnuallot5 
         Caption         =   "Allotment"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "&Report"
   End
   Begin VB.Menu mnupwd 
      Caption         =   "&Password"
      Begin VB.Menu mnuCnu 
         Caption         =   "Create New User"
      End
      Begin VB.Menu pwd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChnp 
         Caption         =   "Change Password"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Eflag As Boolean
Public Flag1 As Boolean
Public Flag2 As Boolean

Private Sub MDIForm_Load()
Call Connect
Iflag = False
End Sub

Private Sub mnuallot_Click()
Form3.Show
End Sub

Private Sub mnuallot1_Click()
Form18.Show
Form18.cmdDelete.Visible = False
Form18.Label3(1).Caption = "SELECT RECORD FOR UPDATION"
End Sub

Private Sub mnuallot3_Click()
Form18.Show
Form18.cmdDelete.Visible = True
Form18.Label3(1).Caption = "SELECT RECORD FOR UPDATION"
End Sub

Private Sub mnuallot4_Click()
Form11.Show
End Sub

Private Sub mnuallot5_Click()
Form12.Show
Form12.Label1.Caption = "ALLOTMENT ANALYSIS"
sql = "SELECT format(ItemIssue.issue_date,'dd-MMM-yy'),ItemIssue.rno, ItemRequest.itemname,ItemIssue.breif_i,ItemIssue.stocksno,ItemIssue.qty_issue,ItemIssue.issued,ItemIssue.transfer,ItemIssue.remark,Indent.empcode, Employ.empname,Employ.dept"
sql = sql & "  FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.qty_issue > 0  order by ItemIssue.issue_date;"
Set rs = cn.Execute(sql)
Set Form12.mhAny.DataSource = rs
Call Form12.GridSet2
End Sub

Private Sub mnuChnp_Click()
Form20.Show
'Form20.Left = 2800
'Form20.Top = 2000
'Form20.Height = 4000
'Form20.Width = 6000
End Sub

Private Sub mnuCnu_Click()
Form21.Show
'Form21.Left = 2800
'Form21.Top = 2000
'Form21.Height = 4000
'Form21.Width = 6000
End Sub

Private Sub mnuemp_Click()
Eflag = True
Form5.Show
Form5.cmdDelete.Visible = False
Form5.cmdUpdate.Visible = False
Form5.cmdAdd.Visible = True
Flag1 = True
Form5.Label8.Visible = False
End Sub

Private Sub mnuEmp1_Click()
Eflag = True
Form5.Show
Form5.cmdDelete.Visible = False
Form5.cmdUpdate.Visible = False
Form5.cmdAdd.Visible = True
Flag1 = True
Form5.Label8.Visible = False
End Sub

Private Sub mnuemp2_Click()
Eflag = False
Form5.Show
Form5.cmdDelete.Visible = False
Form5.cmdUpdate.Visible = True
Form5.cmdAdd.Visible = False
Flag1 = False
End Sub

Private Sub mnuemp3_Click()
Eflag = False
Form5.Show
Form5.cmdDelete.Visible = True
Form5.cmdUpdate.Visible = False
Form5.cmdAdd.Visible = True
Flag1 = False
End Sub

Private Sub mnuemp4_Click()
Form7.Show
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuI_Click()
sql = " SELECT ItemIssue.stocksno,ItemIssue.ItemName,ItemIssue.Breif_i as full ,ItemIssue.qty_issue,Indent.EmpCode, Employ.EmpName,Employ.Dept FROM (ItemIssue INNER JOIN Indent ON val(ItemIssue.indentno_sno) = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.issued = True and ItemIssue.qty_issue>0 and ItemIssue.transfer = False ; "
Set rs = cn.Execute(sql)
Set Report2.DataSource = rs
Report2.Show
End Sub

Private Sub mnuIER_Click()
Form4.Show
Form4.Label1.Caption = "Enter Employee Code"
Form4.cmbItem.Visible = False
Form4.Text1.Visible = True
Form4.Command1.Visible = True
'Form4.Command2.Visible = False
'Form4.Command3.Visible = False
'Form4.cmbReport2.Visible = False
End Sub

Private Sub mnuIIIR_Click()
Form4.Show
Form4.Label1.Caption = "Select Item"
Form4.cmbItem.Visible = True
Form4.Text1.Visible = False
Form4.Command1.Visible = False
'Form4.Command2.Visible = True
'Form4.Command3.Visible = False
'Form4.cmbReport2.Visible = False
End Sub

Private Sub mnuind_Click()
Form1.Show
End Sub

Private Sub mnuInd1_Click()
Form17.Show
Form17.cmdDelete.Visible = False
Form17.Label3.Caption = "SELECT RECORD FOR UPDATION"
End Sub

Private Sub mnuInd3_Click()
Form17.Show
Form17.cmdUpdate.Visible = False
Form17.Label3.Caption = "SELECT RECORD FOR DELETION"
End Sub

Private Sub mnuind4_Click()
Form10.Show
End Sub

Private Sub mnuind5_Click()
Form12.Show
Form12.Label1.Caption = "INDENT ANALYSIS"
sql = "SELECT format(Indent.request_date,'dd-MMM-yy'),ItemRequest.rno,ItemRequest.itemname,ItemRequest.breif_r,ItemRequest.qty_request, Employ.empcode,Employ.empname,Employ.dept,ItemRequest.consider"
sql = sql & "  FROM (Indent INNER JOIN ItemRequest ON Indent.indentno = ItemRequest.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode order by Indent.request_date;"
Set rs = cn.Execute(sql)
Set Form12.mhAny.DataSource = rs
Call Form12.GridSet1
End Sub

Private Sub mnuISIR_Click()
Form4.Show
Form4.Label1.Caption = "Select Item"
Form4.cmbItem.Visible = True
Form4.Text1.Visible = False
Form4.Command1.Visible = False
'Form4.Command2.Visible = False
'Form4.Command3.Visible = True
'Form4.cmbReport2.Visible = False
End Sub

Private Sub mnuIT_Click()
sql = " SELECT ItemIssue.stocksno,ItemIssue.ItemName,ItemIssue.Breif_i ,ItemIssue.qty_issue,Indent.empcode,Employ.Dept FROM (ItemIssue INNER JOIN Indent ON val(ItemIssue.indentno_sno) = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.issued = False and ItemIssue.qty_issue>0 and ItemIssue.transfer = True ; "
Set rs = cn.Execute(sql)
Set Report3.DataSource = rs
Report3.Show
End Sub

Private Sub mnuRD_Click()
sql = "SELECT format(Indent.request_date,'dd-MMM-yy') as Rdate,ItemIssue.*,Indent.EmpCode, Employ.EmpName,Employ.Dept FROM (ItemIssue INNER JOIN Indent ON val(ItemIssue.indentno_sno) = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode order by Indent.request_date; "
Set rs = cn.Execute(sql)
Set Report7.DataSource = rs
Report7.Show
End Sub

Private Sub mnuReport_Click()
Form4.Show
End Sub

Private Sub mnuRetitm_Click()
Form16.Show
End Sub

Private Sub mnuspl2_Click()
Form6.Show
Form6.cmdDelete.Visible = False
Form6.cmdUpdate.Visible = True
Form6.cmdAdd.Visible = False
Flag2 = False
End Sub

Private Sub mnuspl3_Click()
Form5.Show
Form6.cmdDelete.Visible = True
Form6.cmdUpdate.Visible = False
Form6.cmdAdd.Visible = False
Flag2 = False
End Sub

Private Sub mnuspl4_Click()
Form8.Show
End Sub

Private Sub mnuSROSI_Click()
Form4.Show
Form4.Label1.Caption = "Select Item"
Form4.cmbItem.Visible = True
Form4.Text1.Visible = False
Form4.Command1.Visible = False
'Form4.Command2.Visible = False
'Form4.Command3.Visible = False
'Form4.cmbReport2.Visible = True
End Sub

Private Sub mnustk_Click()
Form2.Show
End Sub

Private Sub mnustk1_Click()
Form19.Show
Form19.cmdDelete.Visible = False
Form19.Label18.Caption = "SELECT RECORD FOR UPDATION"
End Sub

Private Sub mnustk3_Click()
Form19.Show
Form19.cmdUpdate.Visible = False
Form19.Label18.Caption = "SELECT RECORD FOR DELETION"
End Sub

Private Sub mnustk4_Click()
Form9.Show
End Sub

Private Sub mnustk5_Click()
Form12.Show
Form12.Label1.Caption = "STOCK ANALYSIS"
Set rs = cn.Execute("select StockSno,ItemId,ItemName,Breif,Detail, format(ItemStock.receipt_date,'dd-MMM-yy') as ReceiptDate,Balance,Receipt_Qty,Unit_Price,Discount,Amount,format(wstart_date,'dd-MMM-yy') as WarrentyStartDate,format(wend_date,'dd-MMM-yy') as WarrentyEndDate,Supplier_ID from ItemStock order by receipt_date")
Set Form12.mhAny.DataSource = rs
Call Form12.GridSet3
End Sub

Private Sub mnuSupl_Click()
Form6.Show
Form6.cmdDelete.Visible = False
Form6.cmdUpdate.Visible = False
Form6.cmdAdd.Visible = True
Flag2 = True
Form6.Label11.Visible = False
End Sub

Private Sub mnuSupl1_Click()
Form6.Show
Form6.cmdDelete.Visible = False
Form6.cmdUpdate.Visible = False
Form6.cmdAdd.Visible = True
Flag2 = True
Form6.Label11.Visible = False
End Sub

Private Sub new1_Click()
Form13.Show
End Sub

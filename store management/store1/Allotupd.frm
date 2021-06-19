VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form18 
   Caption         =   "Form18"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form18"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbBreif 
      Height          =   315
      Left            =   2280
      TabIndex        =   1
      Text            =   " "
      Top             =   2160
      Width           =   2655
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
      Height          =   495
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox cmbStkno 
      Height          =   315
      Left            =   7680
      TabIndex        =   2
      Text            =   " "
      Top             =   2160
      Width           =   1575
   End
   Begin VB.ComboBox cmbIname 
      Height          =   315
      Left            =   2280
      TabIndex        =   0
      Text            =   " "
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Transfer"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox txtRemark 
      Height          =   405
      Left            =   2280
      TabIndex        =   3
      Text            =   " "
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox txtIndentno 
      Enabled         =   0   'False
      Height          =   405
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   " "
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox txtIdate 
      Height          =   405
      Left            =   7680
      TabIndex        =   6
      Text            =   " "
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox txtIqty 
      Height          =   405
      Left            =   7680
      TabIndex        =   4
      Text            =   " "
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "UPDATE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "DELETE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mhIndent 
      Height          =   2415
      Left            =   360
      TabIndex        =   9
      Top             =   4440
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   4260
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   11
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ALLOTMENT UPDATE DELETE FORM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   615
      Left            =   1560
      TabIndex        =   20
      Top             =   0
      Width           =   9375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00400040&
      BorderWidth     =   5
      X1              =   0
      X2              =   12000
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Remark"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   18
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Sr. No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6000
      TabIndex        =   17
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Breif Description"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6000
      TabIndex        =   15
      Top             =   3240
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Req.ItemNo"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   6000
      TabIndex        =   14
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Quantity"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT RECORD FOR UPDATION "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   10
      Top             =   4080
      Width           =   5295
   End
End
Attribute VB_Name = "Form18"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tempsno  As Integer, Bal As Integer, Iqty As Integer
Dim Iflag As Boolean
Dim Bflag As Boolean
Dim Sflag As Boolean


Private Sub cmbBreif_Click()
If Me.ActiveControl <> mhIndent And Me.ActiveControl <> Me.cmdCancel Then
sql = "select stocksno from ItemStock where itemname = '" & cmbIname.Text & "' and  "
sql = sql & " breif = '" & cmbBreif.Text & "' and balance > 0  order by balance"
Set rs = cn.Execute(sql)
If Not rs.EOF Then
cmbStkno.Clear
cmbStkno.Text = rs.Fields("stocksno")
While Not rs.EOF
cmbStkno.AddItem rs.Fields("stocksno")
rs.MoveNext
Wend
Bflag = True
Else
MsgBox "Required Item Not Available", vbInformation, "Update/Delete"
End If
End If
End Sub

Private Sub cmbBreif_LostFocus()
If Me.ActiveControl <> mhIndent And Me.ActiveControl <> Me.cmdCancel Then
If Bflag = False Then
sql = "select stocksno from ItemStock where itemname = '" & cmbIname.Text & "' and  "
sql = sql & " breif = '" & cmbBreif.Text & "' and balance > 0  order by balance"
Set rs = cn.Execute(sql)
If Not rs.EOF Then
cmbStkno.Clear
cmbStkno.Text = rs.Fields("stocksno")
While Not rs.EOF
cmbStkno.AddItem rs.Fields("stocksno")
rs.MoveNext
Wend
Else
MsgBox "Required Item Not Available", vbInformation, "Update/Delete"
End If
End If
End If
End Sub

Private Sub cmbIname_Click()
If Me.ActiveControl <> mhIndent And Me.ActiveControl <> Me.cmdCancel Then
Set rs = cn.Execute("select distinct(breif) from ItemStock  where itemname = '" & cmbIname & "' ")
If Not rs.EOF Then
cmbBreif.Clear
cmbBreif.Text = rs.Fields("breif")
While Not rs.EOF
cmbBreif.AddItem rs.Fields("breif")
rs.MoveNext
Wend
Iflag = True
Else
MsgBox cmbIname.Text + " :" + " Not Available", vbInformation, "Update/Delete"
End If
End If
End Sub

Private Sub cmbIname_LostFocus()
If Me.ActiveControl <> mhIndent And Me.ActiveControl <> Me.cmdCancel Then
If Iflag = False Then
Set rs = cn.Execute("select distinct(breif) from ItemStock  where itemname = '" & cmbIname & "' ")
If Not rs.EOF Then
cmbBreif.Clear
cmbBreif.Text = rs.Fields("breif")
While Not rs.EOF
cmbBreif.AddItem rs.Fields("breif")
rs.MoveNext
Wend
Else
MsgBox cmbIname.Text + " :" + " Not Available", vbInformation, "Update/Delete"
End If
End If
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
cn.Execute ("delete from ItemIssue where indentno_sno = " & txtIndentno & " ")
mhIndent.RemoveItem (mhIndent.RowSel)
cmdDelete.Enabled = False
MsgBox "Record Deleted", vbInformation, "Update/Delete"
End Sub
Private Sub cmdUpdate_Click()
sql = "update ItemIssue set stocksno = " & Val(cmbStkno) & ", "
sql = sql & "  breif_i = '" & cmbBreif.Text & "',issue_date = #" & Format(txtIdate, "dd-MMM-yy") & "#,  "
If Check1.Value = 1 Then
sql = sql & " qty_issue = " & Val(txtIqty) & " ,remark = '" & txtRemark & "',transfer = True where rno = " & Val(txtIndentno) & "  "
Else
sql = sql & " qty_issue = " & Val(txtIqty) & " ,remark = '" & txtRemark & "',transfer = False where rno = " & Val(txtIndentno) & "  "
End If
cn.Execute (sql)

Set rs1 = cn.Execute("select balance from ItemStock where stocksno = " & Val(cmbStkno) & " ")
Bal = rs1.Fields("balance")
cn.Execute ("Update ItemStock set balance = " & Bal - txtIqty & " where stocksno =" & Val(cmbStkno) & " ")

Set rs = cn.Execute("select balance from ItemStock where stocksno = " & Tempsno & " ")
Bal = rs.Fields("balance")
cn.Execute ("Update ItemStock set balance = " & Bal + Iqty & " where stocksno =" & Tempsno & " ")

mhIndent.TextMatrix(mhIndent.RowSel, 6) = txtIqty
mhIndent.TextMatrix(mhIndent.RowSel, 5) = cmbStkno
mhIndent.TextMatrix(mhIndent.RowSel, 3) = cmbIname
mhIndent.TextMatrix(mhIndent.RowSel, 4) = cmbBreif
mhIndent.TextMatrix(mhIndent.RowSel, 1) = Format(txtIdate, "dd-MMM-yy")
txtIqty = mhIndent.TextMatrix(mhIndent.RowSel, 7)
If Check1.Value = 1 Then
Check1.Value = 1
mhIndent.TextMatrix(mhIndent.RowSel, 8) = "True"
Else
mhIndent.TextMatrix(mhIndent.RowSel, 8) = "False"
End If
mhIndent.TextMatrix(mhIndent.RowSel, 9) = txtRemark
txtIndentno = ""
cmbStkno.Clear
cmbIname.Clear
cmbBreif.Clear
txtIdate = ""
txtIqty = ""
txtRemark = ""
Check1.Value = 0
cmdUpdate.Enabled = False
MsgBox "Record Updated", vbInformation, "Update/Delete"
End Sub

Private Sub Form_Load()
Call Connect
sql = "SELECT format(ItemIssue.issue_date,'dd-MMM-yy'),ItemIssue.rno, ItemRequest.itemname,ItemIssue.breif_i,ItemIssue.stocksno,ItemIssue.qty_issue,ItemIssue.issued,ItemIssue.transfer,ItemIssue.remark,Indent.empcode, Employ.empname,Employ.dept"
sql = sql & "  FROM ((ItemIssue INNER JOIN ItemRequest ON ItemIssue.rno = ItemRequest.rno) INNER JOIN Indent ON ItemRequest.indentno = Indent.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemIssue.qty_issue > 0  order by ItemIssue.issue_date;"
Set rs = cn.Execute(sql)
Set mhIndent.DataSource = rs
Call Setting
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

Private Sub mhIndent_Click()
If mhIndent.Row <> 0 And Val(mhIndent.TextMatrix(mhIndent.RowSel, 2)) <> 0 Then
Tempsno = mhIndent.TextMatrix(mhIndent.RowSel, 5)
Iqty = mhIndent.TextMatrix(mhIndent.RowSel, 6)
Set rs = cn.Execute("select itemname from Item")
cmbIname.Clear
cmbIname.Text = rs.Fields("itemname")
While Not rs.EOF
cmbIname.AddItem rs.Fields("itemname")
rs.MoveNext
Wend
txtIndentno = mhIndent.TextMatrix(mhIndent.RowSel, 2)
cmbStkno.Text = mhIndent.TextMatrix(mhIndent.RowSel, 5)
cmbIname = mhIndent.TextMatrix(mhIndent.RowSel, 3)
cmbBreif.Text = mhIndent.TextMatrix(mhIndent.RowSel, 4)
txtIdate = Format(mhIndent.TextMatrix(mhIndent.RowSel, 1), "dd-MMM-yy")
txtIqty = mhIndent.TextMatrix(mhIndent.RowSel, 6)
If mhIndent.TextMatrix(mhIndent.RowSel, 8) = True Then
Check1.Value = 1
Else
Check1.Value = 0
End If
txtRemark = mhIndent.TextMatrix(mhIndent.RowSel, 9)
cmdDelete.Enabled = True
cmdUpdate.Enabled = True
End If
End Sub

Private Sub txtIqty_LostFocus()
If Me.ActiveControl <> cmdCancel Then
Set rs = cn.Execute("select balance from ItemStock where stocksno = " & Val(cmbStkno.Text) & " ")

If Val(txtIqty) = 0 Then
MsgBox "Please Enter Quantity", vbInformation, "Update/Delete"
txtIqty.SetFocus
End If

If Val(txtIqty) > rs.Fields("balance") Then
MsgBox "This Much Of Quantity Not Available" + vbCrLf + "Only" + Str(rs.Fields("balance")) + " " + cmbIname.Text + "  Available", vbInformation, "Update/Delete"
txtIqty.SetFocus
End If

End If
End Sub

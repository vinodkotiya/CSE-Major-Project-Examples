VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form19 
   Caption         =   "Form19"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtSno 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   18
      Text            =   " "
      Top             =   1680
      Width           =   1215
   End
   Begin VB.ComboBox cmbSupplier 
      Height          =   315
      Left            =   4560
      TabIndex        =   17
      Text            =   "cmbSupplier"
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtCdate 
      Height          =   285
      Left            =   1680
      TabIndex        =   16
      Text            =   " "
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtDiscount 
      Height          =   285
      Left            =   4560
      TabIndex        =   15
      Text            =   " "
      Top             =   2160
      Width           =   2415
   End
   Begin VB.ComboBox cmbTime 
      Height          =   315
      Left            =   6000
      Sorted          =   -1  'True
      TabIndex        =   14
      Text            =   "Year"
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtWperiod 
      Height          =   285
      Left            =   4560
      TabIndex        =   13
      Text            =   " "
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox txtAmount 
      Height          =   285
      Left            =   8520
      TabIndex        =   12
      Text            =   " "
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox txtUprice 
      Height          =   285
      Left            =   8520
      TabIndex        =   11
      Text            =   " "
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox txtRqty 
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Text            =   " "
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtDetail 
      Height          =   285
      Left            =   2280
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   3720
      Width           =   7575
   End
   Begin VB.TextBox txtBreif 
      Height          =   285
      Left            =   5280
      TabIndex        =   8
      Top             =   3240
      Width           =   4575
   End
   Begin VB.ComboBox cmbItemlist 
      Height          =   315
      Left            =   4560
      Sorted          =   -1  'True
      TabIndex        =   7
      Text            =   "cmbItemlist"
      Top             =   1680
      Width           =   2415
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "DELETE"
      Enabled         =   0   'False
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "UPDATE"
      Enabled         =   0   'False
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtSrno 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Text            =   " "
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtSid 
      Height          =   285
      Left            =   8520
      TabIndex        =   3
      Text            =   " "
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mhStock 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   4800
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4683
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   14
   End
   Begin MSComCtl2.DTPicker dtpWend 
      Height          =   375
      Left            =   8520
      TabIndex        =   19
      Top             =   2640
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   24707075
      CurrentDate     =   37360
   End
   Begin MSComCtl2.DTPicker dtpWstart 
      Height          =   375
      Left            =   1680
      TabIndex        =   20
      Top             =   2640
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   24707075
      CurrentDate     =   37360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK UPDATE DELETE FORM"
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
      Height          =   495
      Left            =   2040
      TabIndex        =   36
      Top             =   0
      Width           =   8055
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
      Caption         =   "Item- ID"
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
      Left            =   240
      TabIndex        =   35
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier"
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
      Height          =   255
      Left            =   3240
      TabIndex        =   34
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Warrenty Ending Date"
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
      Left            =   7200
      TabIndex        =   33
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Warrenty Starting Date"
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
      Height          =   615
      Left            =   240
      TabIndex        =   32
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   " Date"
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
      Height          =   255
      Left            =   240
      TabIndex        =   31
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount In  (% )"
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
      Height          =   255
      Left            =   3240
      TabIndex        =   30
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Warrenty Period"
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
      Left            =   3240
      TabIndex        =   29
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Height          =   255
      Left            =   7200
      TabIndex        =   28
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price"
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
      Height          =   255
      Left            =   7200
      TabIndex        =   27
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
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
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Detail Description"
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
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   3240
      TabIndex        =   24
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Item List"
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
      Height          =   255
      Left            =   3240
      TabIndex        =   23
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label14 
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
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier ID"
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
      Height          =   255
      Left            =   7200
      TabIndex        =   21
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label18 
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   4440
      Width           =   5655
   End
End
Attribute VB_Name = "Form19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim active As Integer

Private Sub cmbItemlist_Click()
Set rs1 = cn.Execute("select itemid from Item where itemname = '" & cmbItemlist & "' ")
If Not rs.EOF Then
txtSno = rs1.Fields("itemid")
Else
Set rs1 = cn.Execute("select max(itemid) as nxtid from Item ")

If IsNull(rs1.Fields("nxtid")) = True Then
txtSno = 1
Else
txtSno = rs1.Fields("nxtid")
End If

End If
End Sub

Private Sub cmbSupplier_Click()
Set rs = cn.Execute("select supplier_id from Supplier where supplier_name = '" & cmbSupplier & "' ")
txtSid = rs.Fields("supplier_id")
End Sub
Private Sub cmdDelete_Click()
cn.Execute ("delete from ItemStock where stocksno = " & Val(txtSrno) & " ")
mhStock.RemoveItem (mhStock.RowSel)
MsgBox "Record Deleted", vbInformation, "Update/Delete"
cmdDelete.Enabled = False
End Sub

Private Sub cmdUpdate_Click()
cn.Execute ("delete from ItemStock where stocksno = " & Val(txtSrno) & " ")
txtAmount = (Val(txtRqty) * Val(txtUprice)) - (Val(txtRqty) * Val(txtUprice) * (Val(txtDiscount) / 100))
If Trim(txtWperiod) <> "" Then
Select Case cmbTime
Case "Year":
dtpWend.Value = DateAdd("d", Val(txtWperiod) * 365, dtpWstart.Value)
Case "Month":
dtpWend.Value = DateAdd("d", Val(txtWperiod) * 30, dtpWstart.Value)
Case "Day":
dtpWend.Value = DateAdd("d", Val(txtWperiod), dtpWstart.Value)
Case "Week":
dtpWend.Value = DateAdd("d", Val(txtWperiod) * 7, dtpWstart.Value)
End Select
End If

 mhStock.TextMatrix(mhStock.Row, 2) = txtSrno
 mhStock.TextMatrix(mhStock.Row, 5) = txtCdate
 mhStock.TextMatrix(mhStock.Row, 8) = txtUprice
 mhStock.TextMatrix(mhStock.Row, 10) = txtAmount
 mhStock.TextMatrix(mhStock.Row, 7) = txtRqty
 mhStock.TextMatrix(mhStock.Row, 11) = Format(dtpWstart.Value, "dd-MMM-yy")
 mhStock.TextMatrix(mhStock.Row, 12) = Format(dtpWend.Value, "dd-MMM-yy")
 mhStock.TextMatrix(mhStock.Row, 0) = txtSno
 mhStock.TextMatrix(mhStock.Row, 9) = txtDiscount
 mhStock.TextMatrix(mhStock.Row, 3) = txtBreif
 mhStock.TextMatrix(mhStock.Row, 4) = txtDetail
 mhStock.TextMatrix(mhStock.Row, 1) = cmbItemlist
 mhStock.TextMatrix(mhStock.Row, 13) = txtSid

sql = " insert into ItemStock values(" & Val(txtSrno) & "," & txtSno & ",'" & cmbItemlist & "','" & txtBreif & "','" & txtDetail & "',#" & Format(txtCdate, "dd-MMM-yy") & "#," & Val(txtRqty) & "," & Val(txtRqty) & "," & Val(txtUprice) & "," & Val(txtDiscount) & "," & Val(txtAmount) & ",#" & Format(dtpWstart.Value, "dd-MMM-yy") & "#,#" & Format(dtpWend.Value, "dd-MMM-yy") & "#," & Val(txtSid) & ")"
cn.Execute sql
MsgBox "Record Updated", vbInformation, "Update/Delete"
cmdUpdate.Enabled = False
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call Connect
Set rs = cn.Execute("select ItemId,ItemName,StockSno as StkSno,Breif,Detail, format(ItemStock.receipt_date,'dd-MMM-yy') as ReceiptDate,Balance,Receipt_Qty as ReceiptQty,Unit_Price as UnitPrice,Discount,Amount,format(wstart_date,'dd-MMM-yy') as WarrentyStartDate,format(wend_date,'dd-MMM-yy') as WarrentyEndDate,Supplier_ID from ItemStock order by receipt_date")
Set mhStock.DataSource = rs
Call Setting
End Sub
Public Sub Setting()
mhStock.ColWidth(0) = 700
mhStock.ColWidth(1) = 1100
mhStock.ColWidth(2) = 700
mhStock.ColWidth(3) = 1400
mhStock.ColWidth(4) = 1700
mhStock.ColWidth(5) = 1200
mhStock.ColWidth(6) = 900
mhStock.ColWidth(7) = 900
mhStock.ColWidth(8) = 850
mhStock.ColWidth(9) = 800
mhStock.ColWidth(10) = 900
mhStock.ColWidth(11) = 1450
mhStock.ColWidth(12) = 1450
mhStock.ColWidth(13) = 900
End Sub

Private Sub mhStock_Click()
If Val(mhStock.TextMatrix(mhStock.Row, 2)) <> 0 And mhStock.Row <> 0 Then
Set rs = cn.Execute("select itemname from Item")
Set rs1 = cn.Execute("select supplier_name from Supplier")
While Not rs.EOF
cmbItemlist.AddItem rs.Fields("itemname")
rs.MoveNext
Wend
While Not rs1.EOF
cmbSupplier.AddItem rs1.Fields("supplier_name")
rs1.MoveNext
Wend
txtSrno = mhStock.TextMatrix(mhStock.Row, 2)
txtCdate = mhStock.TextMatrix(mhStock.Row, 5)
txtUprice = mhStock.TextMatrix(mhStock.Row, 8)
txtAmount = mhStock.TextMatrix(mhStock.Row, 10)
txtRqty = mhStock.TextMatrix(mhStock.Row, 7)
If mhStock.TextMatrix(mhStock.Row, 11) <> "" Then
dtpWstart.Value = mhStock.TextMatrix(mhStock.Row, 11)
End If
If mhStock.TextMatrix(mhStock.Row, 12) <> "" Then
dtpWend.Value = mhStock.TextMatrix(mhStock.Row, 12)
End If
txtSno = mhStock.TextMatrix(mhStock.Row, 0)
txtDiscount = mhStock.TextMatrix(mhStock.Row, 9)
txtBreif = mhStock.TextMatrix(mhStock.Row, 3)
txtDetail = mhStock.TextMatrix(mhStock.Row, 4)
cmbItemlist.Text = mhStock.TextMatrix(mhStock.Row, 1)
txtSid = mhStock.TextMatrix(mhStock.Row, 13)
Set rs1 = cn.Execute("select supplier_name from Supplier where supplier_id =  " & mhStock.TextMatrix(mhStock.Row, 13) & " ")
cmbSupplier.Text = rs1.Fields("supplier_name")
txtWperiod = DateDiff("d", dtpWstart.Value, dtpWend.Value) / 365

cmdDelete.Enabled = True
cmdUpdate.Enabled = True
End If
End Sub

VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   Caption         =   "Stock Entry Form"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbItem 
      Height          =   315
      Left            =   7560
      Sorted          =   -1  'True
      TabIndex        =   1
      Text            =   " "
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txtItemid 
      Height          =   285
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   " "
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox txtBreif 
      Height          =   285
      Left            =   7560
      TabIndex        =   3
      Top             =   1680
      Width           =   2535
   End
   Begin VB.TextBox txtDetail 
      Height          =   525
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   3960
      Width           =   7335
   End
   Begin VB.TextBox txtRqty 
      Height          =   285
      Left            =   2760
      TabIndex        =   4
      Text            =   " "
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox txtUprice 
      Height          =   285
      Left            =   7560
      TabIndex        =   5
      Text            =   " "
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox txtAmount 
      Height          =   285
      Left            =   7560
      TabIndex        =   11
      Text            =   " "
      Top             =   3480
      Width           =   2535
   End
   Begin VB.TextBox txtWperiod 
      Height          =   285
      Left            =   7560
      TabIndex        =   8
      Text            =   " "
      Top             =   3120
      Width           =   1695
   End
   Begin VB.ComboBox cmbTime 
      Height          =   315
      Left            =   9240
      Sorted          =   -1  'True
      TabIndex        =   9
      Text            =   "Year"
      Top             =   3120
      Width           =   855
   End
   Begin VB.TextBox txtDiscount 
      Height          =   285
      Left            =   7560
      TabIndex        =   6
      Text            =   " "
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox txtCdate 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2760
      TabIndex        =   15
      Text            =   " "
      Top             =   960
      Width           =   2295
   End
   Begin VB.ComboBox cmbSupplier 
      Height          =   315
      Left            =   7560
      TabIndex        =   0
      Text            =   " "
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtUstock 
      Height          =   285
      Left            =   2760
      TabIndex        =   27
      Text            =   " "
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox txtMorder 
      Height          =   285
      Left            =   7560
      TabIndex        =   26
      Text            =   " "
      Top             =   2760
      Width           =   2535
   End
   Begin VB.TextBox txtSno 
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Text            =   " "
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtTamount 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9480
      TabIndex        =   24
      Text            =   " "
      Top             =   7560
      Width           =   2175
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H00FFFF00&
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdTok 
      BackColor       =   &H00FFFF00&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFF00&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   495
      Left            =   10320
      TabIndex        =   19
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFF00&
      Caption         =   "Delete"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H00FFFF00&
      Caption         =   "Update"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFF00&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   960
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mhStock 
      Height          =   2535
      Left            =   120
      TabIndex        =   14
      Top             =   4920
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4471
      _Version        =   393216
      BackColor       =   8421631
      Rows            =   20
      Cols            =   14
      FixedCols       =   0
      BackColorFixed  =   8421631
      BackColorBkg    =   8421631
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   14
   End
   Begin MSComCtl2.DTPicker dtpWend 
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   3360
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   24969219
      CurrentDate     =   37360
   End
   Begin MSComCtl2.DTPicker dtpWstart 
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   2880
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarForeColor=   0
      CalendarTitleBackColor=   12632064
      CalendarTitleForeColor=   12632064
      CalendarTrailingForeColor=   12632064
      CustomFormat    =   "dd-MMM-yy"
      Format          =   24969219
      CurrentDate     =   37360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   5
      X1              =   0
      X2              =   11880
      Y1              =   720
      Y2              =   720
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5400
      TabIndex        =   44
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Sr. No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   43
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5400
      TabIndex        =   42
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Detail Description"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   41
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Quantity"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Price Rs /-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5400
      TabIndex        =   39
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount Rs /-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5400
      TabIndex        =   38
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Warrenty Period"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   5400
      TabIndex        =   37
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Discount In  (% )"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5400
      TabIndex        =   36
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label11 
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
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   35
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Warrenty Starting Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   34
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Warrenty Ending Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5400
      TabIndex        =   32
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit In Stock(Nos.)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   120
      TabIndex        =   31
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Min. Order Level"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   5400
      TabIndex        =   30
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Item- ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK ENTRY FORM"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   735
      Left            =   3960
      TabIndex        =   25
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount Rs /-"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   6960
      TabIndex        =   23
      Top             =   7560
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Record From List For Updation Or Deletion (In Case Of Mistake During Entry)"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   22
      Top             =   4560
      Width           =   10935
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Nextitm As Integer, nextid As Integer, rowno As Integer
Dim nextid1 As Integer, nextid2 As Integer
Dim i As Integer, j As Integer
Dim Flag1 As Boolean
Dim Flag2  As String
Dim Flag3 As Boolean
Dim delitem As Integer
Dim sid As Integer

Private Sub cmbItem_GotFocus()
cmbItem = " "
End Sub

Private Sub cmbItem_LostFocus()
If Me.ActiveControl <> cmdExit And Me.ActiveControl <> mhStock Then
If Trim(cmbItem.Text) = "" Then
MsgBox "You Must Select/Enter An Item Name", vbExclamation, "Warning"
cmbItem.SetFocus
cmbItem_GotFocus
Else
 Set rs = cn.Execute("select itemid from Item where itemname = '" & Initcap(cmbItem) & "'  ")
 If rs.EOF Then
 Set rs = cn.Execute("select max(itemid) as maxcode from Item ")
  If IsNull(rs.Fields("maxcode")) Then
  nextid1 = 1
  Else
  nextid1 = rs.Fields("maxcode") + 1
  End If
  If Trim(cmbItem.Text) <> "" Then
 cn.Execute ("insert into Item(itemid,itemname)  values(" & nextid1 & ",'" & cmbItem & "')")
 End If
 txtSno = nextid1
 Else
 txtSno = rs.Fields("itemid")
 End If
 txtBreif.SetFocus
End If
End If
Set rs = cn.Execute("select sum(balance) as total  from ItemStock where Itemstock.itemname = '" & cmbItem & "'")
If IsNull(rs.Fields("total")) = False Then
txtUstock = rs.Fields("total")
End If
If cmbItem <> " " Then
Set rs = cn.Execute("select MinOrdQty from Item where itemname = '" & cmbItem & "'")
txtMorder = rs.Fields("MinOrdQty")
End If
End Sub

Private Sub cmbSupplier_LostFocus()
Set rs = cn.Execute("select supplier_id from Supplier where supplier_name = '" & cmbSupplier & "' ")
If rs.EOF Then
Set rs1 = cn.Execute("select max(supplier_id) as id from Supplier")
If IsNull(rs1.Fields("id")) = True Then
nextid2 = 1
Else
nextid2 = rs1.Fields("id") + 1
End If
If Trim(cmbSupplier.Text) <> "" Then
cn.Execute ("Insert into Supplier(supplier_id,supplier_name) values(" & nextid2 & ",'" & cmbSupplier & "') ")
End If
End If
End Sub

Private Sub cmbTime_LostFocus()
If txtWperiod <> " " Then
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
End Sub

Private Sub cmdDelete_Click()
Flag2 = MsgBox("Are You Sure You Want To Remove Record ", vbYesNo + vbInformation, "Stoc Entry")
If Flag2 = vbYes Then
i = mhStock.TextMatrix(mhStock.RowSel, 3)
j = mhStock.RowSel
mhStock.RemoveItem (mhStock.RowSel)
While Val(mhStock.TextMatrix(j, 3)) <> 0
mhStock.TextMatrix(j, 3) = i
i = i + 1
j = j + 1
Wend
Flag3 = True
rowno = j
nextid = nextid - 1
Call Clear
txtItemid = nextid
cmdDelete.Enabled = False
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdReset_Click()
Unload Me
Me.Show
End Sub

Private Sub cmdSave_Click()
Flag2 = MsgBox("Are You Sure About Entered Data ", vbYesNo + vbInformation, "Stoc Entry")
If Flag2 = vbNo Then
Me.Show
End If
If Flag2 = vbYes Then
'cn.Execute ("insert into LotDetail values(" & Val(mhStock.TextMatrix(1, 1)) & ", #" & txtCdate & "#)")
j = 1
While Val(mhStock.TextMatrix(j, 3)) <> 0
Set rs = cn.Execute("select supplier_id from Supplier where supplier_name = '" & mhStock.TextMatrix(j, 2) & "'  ")
sid = rs.Fields("supplier_id")
sql = "insert into ItemStock values(" & mhStock.TextMatrix(j, 0) & "," & mhStock.TextMatrix(j, 3) & ","
sql = sql & " '" & mhStock.TextMatrix(j, 4) & "','" & mhStock.TextMatrix(j, 5) & "','" & Val(mhStock.TextMatrix(j, 13)) & "',#" & mhStock.TextMatrix(j, 1) & "#,"
sql = sql & " " & Val(mhStock.TextMatrix(j, 6)) & "," & Val(mhStock.TextMatrix(j, 6)) & "," & Val(mhStock.TextMatrix(j, 7)) & ","
sql = sql & " " & Val(mhStock.TextMatrix(j, 8)) & "," & Val(mhStock.TextMatrix(j, 9)) & ","
sql = sql & " #" & Format(mhStock.TextMatrix(j, 10), "dd-MMM-yy") & "#  , #" & Format(mhStock.TextMatrix(j, 12), "dd-MMM-yy") & "# , " & sid & ") "
On Error GoTo Ehand
cn.Execute sql
j = j + 1
Wend
End If
Ehand:
If Err.Number = -2147217900 Then
Resume Next
End If
cmdReset_Click
End Sub

Private Sub cmdTok_Click()
If rowno > mhStock.Rows - 1 Then
mhStock.Rows = mhStock.Rows + 10
End If
mhStock.TextMatrix(rowno, 0) = Me.txtItemid
mhStock.TextMatrix(rowno, 1) = Me.txtCdate
mhStock.TextMatrix(rowno, 2) = Me.cmbSupplier
mhStock.TextMatrix(rowno, 3) = txtSno
mhStock.TextMatrix(rowno, 4) = Me.cmbItem
mhStock.TextMatrix(rowno, 5) = txtBreif
mhStock.TextMatrix(rowno, 6) = txtRqty
mhStock.TextMatrix(rowno, 7) = txtUprice
mhStock.TextMatrix(rowno, 8) = txtDiscount
mhStock.TextMatrix(rowno, 9) = txtAmount
mhStock.TextMatrix(rowno, 10) = Format(dtpWstart, "dd-MMM-yy")
mhStock.TextMatrix(rowno, 11) = Val(txtWperiod) & "-" & cmbTime
mhStock.TextMatrix(rowno, 12) = Format(dtpWend, "dd-MMM-yy")
mhStock.TextMatrix(rowno, 13) = txtDetail
rowno = rowno + 1
cmdTok.Visible = False
txtTamount = Val(txtTamount) + txtAmount
nextid = nextid + 1
Call Clear
End Sub

Private Sub cmdUpdate_Click()
mhStock.TextMatrix(mhStock.RowSel, 0) = txtItemid
mhStock.TextMatrix(mhStock.RowSel, 1) = txtCdate
mhStock.TextMatrix(mhStock.RowSel, 2) = Me.cmbSupplier
mhStock.TextMatrix(mhStock.RowSel, 3) = txtSno
mhStock.TextMatrix(mhStock.RowSel, 4) = Me.cmbItem
mhStock.TextMatrix(mhStock.RowSel, 5) = txtBreif
mhStock.TextMatrix(mhStock.RowSel, 6) = txtRqty
mhStock.TextMatrix(mhStock.RowSel, 7) = txtUprice
mhStock.TextMatrix(mhStock.RowSel, 8) = txtDiscount
mhStock.TextMatrix(mhStock.RowSel, 9) = txtAmount
mhStock.TextMatrix(mhStock.RowSel, 10) = Format(dtpWstart, "dd-MMM-yy")
mhStock.TextMatrix(mhStock.RowSel, 11) = Val(txtWperiod) & "-" & cmbTime
mhStock.TextMatrix(mhStock.RowSel, 12) = Format(dtpWend, "dd-MMM-yy")
mhStock.TextMatrix(mhStock.RowSel, 13) = txtDetail
cmdTok.Visible = False
cmdUpdate.Enabled = True
cmdDelete.Enabled = True
txtTamount = Val(txtTamount) + txtAmount
Call Clear
cmbItem.SetFocus
End Sub

Private Sub Form_Load()
Call Connect
rowno = 1
i = 1
Flag1 = False
Flag3 = False
txtCdate = Format(Now, "dd-MMM-yy")
Set rs = cn.Execute("select itemname from Item ")
While Not rs.EOF
cmbItem.AddItem rs.Fields("itemname")
rs.MoveNext
Wend

Set rs = cn.Execute("select supplier_name from supplier ")
While Not rs.EOF
cmbSupplier.AddItem rs.Fields("supplier_name")
rs.MoveNext
Wend

Set rs1 = cn.Execute("select max(stocksno) as maxcode from ItemStock ")
If IsNull(rs1.Fields("maxcode")) Then
nextid = 1
Else
nextid = rs1.Fields("maxcode") + 1
End If
txtItemid = nextid
Call Heading
End Sub

Private Sub mhStock_Click()
If mhStock.Rows > 2 Then
txtItemid = mhStock.TextMatrix(mhStock.Row, 0)
txtCdate = mhStock.TextMatrix(mhStock.Row, 1)
cmbSupplier = mhStock.TextMatrix(mhStock.Row, 2)
txtSno = mhStock.TextMatrix(mhStock.Row, 3)
cmbItem = mhStock.TextMatrix(mhStock.Row, 4)
txtBreif = mhStock.TextMatrix(mhStock.Row, 5)
txtRqty = mhStock.TextMatrix(mhStock.Row, 6)
txtUprice = mhStock.TextMatrix(mhStock.Row, 7)
txtDiscount = mhStock.TextMatrix(mhStock.Row, 8)
txtAmount = mhStock.TextMatrix(mhStock.Row, 9)
dtpWstart = mhStock.TextMatrix(mhStock.Row, 10)
txtWperiod = Val(mhStock.TextMatrix(mhStock.Row, 11))
cmbTime = Mid(mhStock.TextMatrix(mhStock.Row, 11), InStr(mhStock.TextMatrix(mhStock.Row, 11), "-") + 1)
dtpWend = mhStock.TextMatrix(mhStock.Row, 12)
txtDetail = mhStock.TextMatrix(mhStock.Row, 13)
Flag1 = True
txtTamount = Val(txtTamount) - Val(mhStock.TextMatrix(mhStock.Row, 8))
cmdUpdate.Enabled = True
cmdDelete.Enabled = True

Set rs = cn.Execute(" select sum(balance) as total  from ItemStock where Itemstock.itemname = '" & mhStock.TextMatrix(mhStock.Row, 4) & "'")
If IsNull(rs.Fields("total")) = False Then
txtUstock = rs.Fields("total")
End If
If cmbItem <> " " Then
Set rs = cn.Execute(" select MinOrdQty from Item where itemname = '" & mhStock.TextMatrix(mhStock.Row, 4) & "'")
txtMorder = rs.Fields("MinOrdQty")
End If
End If
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
Call NumberOnly(KeyAscii)
End Sub

Private Sub txtDetail_LostFocus()
txtDetail = Initcap(txtDetail)
End Sub

Private Sub txtDiscount_KeyPress(KeyAscii As Integer)
Call NumberOnly(KeyAscii)
End Sub

Private Sub txtDiscount_LostFocus()
If Val(txtDiscount) = 0 Then
txtAmount = Val(txtRqty) * Val(txtUprice)
Else
txtAmount = (Val(txtRqty) * Val(txtUprice)) - ((Val(txtRqty) * Val(txtUprice)) * Val(txtDiscount) / 100)
End If
End Sub
Public Sub Heading()
mhStock.ColWidth(0) = 1200
mhStock.ColWidth(1) = 1100
mhStock.ColWidth(2) = 1700
mhStock.ColWidth(3) = 700
mhStock.ColWidth(4) = 2000
mhStock.ColWidth(5) = 2800
mhStock.ColWidth(6) = 800
mhStock.ColWidth(7) = 800
mhStock.ColWidth(8) = 800
mhStock.ColWidth(9) = 800
mhStock.ColWidth(10) = 1100
mhStock.ColWidth(11) = 800
mhStock.ColWidth(12) = 1100
mhStock.ColWidth(13) = 3400
mhStock.TextMatrix(0, 0) = "Stock Sr.No."
mhStock.TextMatrix(0, 1) = "Date"
mhStock.TextMatrix(0, 2) = "Supplier Name"
mhStock.TextMatrix(0, 3) = "Item-ID"
mhStock.TextMatrix(0, 4) = "Item Name"
mhStock.TextMatrix(0, 5) = "Brief Description"
mhStock.TextMatrix(0, 6) = "Quantity"
mhStock.TextMatrix(0, 7) = "Unit Rate"
mhStock.TextMatrix(0, 8) = "Discount"
mhStock.TextMatrix(0, 9) = "Amount"
mhStock.TextMatrix(0, 10) = "WarrentyStart"
mhStock.TextMatrix(0, 11) = "Warrenty"
mhStock.TextMatrix(0, 12) = "WarrentyEnd"
mhStock.TextMatrix(0, 13) = "Detail Description"
End Sub

Private Sub txtLotno_KeyPress(KeyAscii As Integer)
Call NumberOnly(KeyAscii)
End Sub

Private Sub txtRqty_KeyPress(KeyAscii As Integer)
Call NumberOnly(KeyAscii)
End Sub

Private Sub txtUprice_KeyPress(KeyAscii As Integer)
Call NumberOnly(KeyAscii)
End Sub

Private Sub txtUprice_LostFocus()
If Flag1 = False Then
cmdTok.Visible = True
txtAmount = Val(txtRqty) * Val(txtUprice)
End If
End Sub
Public Sub Clear()
'cmbSupplier = ""
cmbItem = ""
txtUstock = ""
txtMorder = ""
txtItemid = nextid
txtBreif = ""
txtRqty = ""
txtUprice = ""
txtDiscount = ""
txtWperiod = ""
txtDetail = ""
txtAmount = ""
Flag1 = False
cmdUpdate.Enabled = False
cmdDelete.Enabled = False
cmbItem.SetFocus
End Sub

Private Sub txtWperiod_KeyPress(KeyAscii As Integer)
Call NumberOnly(KeyAscii)
End Sub

Private Sub txtWperiod_LostFocus()
dtpWend.Value = DateAdd("d", Val(txtWperiod) * 365, dtpWstart.Value)
End Sub

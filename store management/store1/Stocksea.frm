VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form9 
   BackColor       =   &H00400040&
   Caption         =   "Form9"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form9"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ALL"
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
      Left            =   10680
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
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9120
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
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mhStock 
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5318
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   14
   End
   Begin MSComCtl2.DTPicker dtpStk 
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   24510467
      CurrentDate     =   37379
   End
   Begin VB.ComboBox cmbItem 
      Height          =   315
      Left            =   3240
      TabIndex        =   3
      Text            =   " "
      Top             =   2160
      Width           =   3255
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1080
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search By Item Name"
            Key             =   "sbn"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search By Receipt Date"
            Key             =   "sbd"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search By Receipt Month "
            Key             =   "sbm"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search Between  Receipt Dates"
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
   Begin MSComCtl2.DTPicker dtpStk1 
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   24510467
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
      Left            =   5160
      TabIndex        =   10
      Top             =   2280
      Width           =   375
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
      Caption         =   "Select Item Name"
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
      Left            =   360
      TabIndex        =   2
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK PROFILE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim active As Integer

Private Sub cmdAll_Click()
Call Setting
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdOk_Click()
Select Case Tab1.SelectedItem.Index
Case 1:
Set rs = cn.Execute("select ItemId,ItemName,Stocksno as StkSno,Breif,Detail, format(ItemStock.receipt_date,'dd-MMM-yy') as ReceiptDate,Balance,Receipt_Qty as ReceiptQty,Unit_Price as UnitPrice,Amount,Discount,format(wstart_date,'dd-MMM-yy') as WarrentyStartDate,format(wend_date,'dd-MMM-yy') as WarrentyEndDate,Supplier_ID from ItemStock where itemname = '" & cmbItem.Text & "' ")
Set mhStock.DataSource = rs

Case 2:
Set rs = cn.Execute("select ItemId,ItemName,stocksno as StkSno,Breif,Detail, format(ItemStock.receipt_date,'dd-MMM-yy') as ReceiptDate,Balance,Receipt_Qty as ReceiptQty,Unit_Price as UnitPrice,Amount,Discount,format(wstart_date,'dd-MMM-yy') as WarrentyStartDate,format(wend_date,'dd-MMM-yy') as WarrentyEndDate,Supplier_ID from ItemStock  where receipt_date = #" & Format(dtpStk.Value, "dd-MMM-yy") & "# ")
Set mhStock.DataSource = rs


Case 3:
Set rs = cn.Execute("select ItemId,ItemName,stocksno as StkSno,Breif,Detail, format(ItemStock.receipt_date,'dd-MMM-yy') as ReceiptDate,Balance,Receipt_Qty as ReceiptQty,Unit_Price as UnitPrice,Amount,Discount,format(wstart_date,'dd-MMM-yy') as WarrentyStartDate,format(wend_date,'dd-MMM-yy') as WarrentyEndDate,Supplier_ID from ItemStock where Month(receipt_date) = '" & Month(dtpStk.Value) & "' and Year(receipt_date) = '" & Year(dtpStk.Value) & "'  ")
Set mhStock.DataSource = rs


Case 4:
Set rs = cn.Execute("select ItemId,ItemName,stocksno as StkSno,Breif,Detail, format(ItemStock.receipt_date,'dd-MMM-yy') as ReceiptDate,Balance,Receipt_Qty as ReceiptQty,Unit_Price as UnitPrice,Amount,Discount,format(wstart_date,'dd-MMM-yy') as WarrentyStartDate,format(wend_date,'dd-MMM-yy') as WarrentyEndDate,Supplier_ID from ItemStock where receipt_date >= #" & Format(dtpStk.Value, "dd-MMM-yy") & "#  and receipt_date <= #" & Format(dtpStk1.Value, "dd-MMM-yy") & "# ")
Set mhStock.DataSource = rs
End Select
Call Setting
End Sub
Private Sub Form_Load()
Call Connect
Set rs = cn.Execute("select StockSno,ItemId,ItemName,Breif,Detail, format(ItemStock.receipt_date,'dd-MMM-yy') as ReceiptDate,Balance,Receipt_Qty,Unit_Price,Discount,Amount,format(wstart_date,'dd-MMM-yy') as WarrentyStartDate,format(wend_date,'dd-MMM-yy') as WarrentyEndDate,Supplier_ID from ItemStock order by receipt_date")
Set mhStock.DataSource = rs
Call Setting
dtpStk.Visible = False
cmbItem.Visible = True
dtpStk1.Visible = False
Label3.Visible = False
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

Private Sub Tab1_Click()
Select Case Tab1.SelectedItem.Index
Case 1:
Label2.Caption = "Select Item Name "
dtpStk.Visible = False
cmbItem.Visible = True
dtpStk1.Visible = False
Label3.Visible = False
Set rs = cn.Execute("select itemname from Item")
While Not rs.EOF
cmbItem.AddItem rs.Fields("itemname")
rs.MoveNext
Wend
Case 2:
Label2.Caption = "Select Receipt Date "
dtpStk.Visible = True
cmbItem.Visible = False
dtpStk1.Visible = False
Label3.Visible = False
Case 3:
Label2.Caption = "Select  Receipt Month & Year "
dtpStk.Visible = True
cmbItem.Visible = False
dtpStk1.Visible = False
Label3.Visible = False

Case 4:
Label2.Caption = "Select Receipt Date From: "
dtpStk.Visible = True
cmbItem.Visible = False
dtpStk1.Visible = True
Label3.Visible = True
End Select
End Sub

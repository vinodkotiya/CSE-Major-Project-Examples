VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form10 
   BackColor       =   &H00404040&
   Caption         =   "Form10"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form10"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbItem 
      Height          =   315
      Left            =   3480
      TabIndex        =   8
      Text            =   " "
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00C0C0C0&
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
      Left            =   10320
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0C0&
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
      TabIndex        =   6
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0C0C0&
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
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mhIndent 
      Height          =   3375
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   5953
      _Version        =   393216
      BackColor       =   14737632
      BackColorFixed  =   8421504
      BackColorBkg    =   12632256
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker dtpIndent 
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yy"
      Format          =   24576003
      CurrentDate     =   37379
   End
   Begin VB.TextBox txtEmpcode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Text            =   " "
      Top             =   2160
      Width           =   2055
   End
   Begin MSComctlLib.TabStrip Tab1 
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search By Employee Code"
            Key             =   "sbc"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search By Item Name"
            Key             =   "sbin"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search By Date"
            Key             =   "sbd"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search By Month"
            Key             =   "sbm"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search Between Dates"
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
      Format          =   24576003
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
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "INDENT SEARCH"
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
      Left            =   3480
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAll_Click()
sql = "SELECT format(Indent.request_date,'dd-MMM-yy'),ItemRequest.rno,ItemRequest.itemname,ItemRequest.breif_r,ItemRequest.qty_request, Employ.empcode,Employ.empname,Employ.dept"
sql = sql & "  FROM (Indent INNER JOIN ItemRequest ON Indent.indentno = ItemRequest.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemRequest.consider = false order by Indent.request_date;"
Set rs = cn.Execute(sql)
Set mhIndent.DataSource = rs
Call GridSet
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Select Case Tab1.SelectedItem.Index
Case 1:
sql = "SELECT format(Indent.request_date,'dd-MMM-yy'),ItemRequest.rno,ItemRequest.itemname,ItemRequest.breif_r,ItemRequest.qty_request, Employ.empcode,Employ.empname,Employ.dept"
sql = sql & "  FROM (Indent INNER JOIN ItemRequest ON Indent.indentno = ItemRequest.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemRequest.consider = false and Indent.empcode = " & Val(txtEmpcode) & "  order by Indent.request_date  ;"
Case 2:
sql = "SELECT format(Indent.request_date,'dd-MMM-yy'),ItemRequest.rno,ItemRequest.itemname,ItemRequest.breif_r,ItemRequest.qty_request, Employ.empcode,Employ.empname,Employ.dept"
sql = sql & "  FROM (Indent INNER JOIN ItemRequest ON Indent.indentno = ItemRequest.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemRequest.consider = false and ItemRequest.itemname = '" & cmbItem & "' order by Indent.request_date  ;"
Case 3:
sql = "SELECT format(Indent.request_date,'dd-MMM-yy'),ItemRequest.rno,ItemRequest.itemname,ItemRequest.breif_r,ItemRequest.qty_request, Employ.empcode,Employ.empname,Employ.dept"
sql = sql & "  FROM (Indent INNER JOIN ItemRequest ON Indent.indentno = ItemRequest.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemRequest.consider = false and Indent.request_date = #" & Format(dtpIndent.Value, "dd-MMM-yy") & "#  order by Indent.request_date ;"
Case 4:
sql = "SELECT format(Indent.request_date,'dd-MMM-yy'),ItemRequest.rno,ItemRequest.itemname,ItemRequest.breif_r,ItemRequest.qty_request, Employ.empcode,Employ.empname,Employ.dept"
sql = sql & "  FROM (Indent INNER JOIN ItemRequest ON Indent.indentno = ItemRequest.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode  where ItemRequest.consider = false and Month(indent.request_date) = '" & Month(dtpIndent.Value) & "' and year(indent.request_date) = '" & Year(dtpIndent.Value) & "'  order by Indent.request_date ;"
Case 5:
If dtpIndent.Value = dtpIndent1.Value Then
sql = "SELECT format(Indent.request_date,'dd-MMM-yy'),ItemRequest.rno,ItemRequest.itemname,ItemRequest.breif_r,ItemRequest.qty_request, Employ.empcode,Employ.empname,Employ.dept"
sql = sql & "  FROM (Indent INNER JOIN ItemRequest ON Indent.indentno = ItemRequest.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemRequest.consider = false and Indent.request_date = #" & Format(dtpIndent.Value, "dd-MMM-yy") & "#  order by Indent.request_date ;"
Else
sql = "SELECT format(Indent.request_date,'dd-MMM-yy'),ItemRequest.rno,ItemRequest.itemname,ItemRequest.breif_r,ItemRequest.qty_request, Employ.empcode,Employ.empname,Employ.dept"
sql = sql & "  FROM (Indent INNER JOIN ItemRequest ON Indent.indentno = ItemRequest.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode  where ItemRequest.consider = false and indent.request_date >= #" & Format(dtpIndent.Value, "dd-MMM-yy") & "# and indent.request_date <= #" & Format(dtpIndent1.Value, "dd-MMM-yy") & "#  order by Indent.request_date ;"
End If
End Select
Set rs = cn.Execute(sql)
Set mhIndent.DataSource = rs
Call GridSet
End Sub
Private Sub Form_Load()
Call Connect
sql = "SELECT format(Indent.request_date,'dd-MMM-yy'),ItemRequest.rno,ItemRequest.itemname,ItemRequest.breif_r,ItemRequest.qty_request, Employ.empcode,Employ.empname,Employ.dept"
sql = sql & "  FROM (Indent INNER JOIN ItemRequest ON Indent.indentno = ItemRequest.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemRequest.consider = false order by Indent.request_date;"
Set rs = cn.Execute(sql)
Set mhIndent.DataSource = rs
Call GridSet
cmbItem.Visible = False
dtpIndent.Visible = False
Label3.Visible = False
dtpIndent1.Visible = False
End Sub
Private Sub Tab1_Click()
Select Case Tab1.SelectedItem.Index
Case 1:
Label2(0).Caption = "Enter Employee Code"
txtEmpcode.Visible = True
cmbItem.Visible = False
dtpIndent.Visible = False
Label3.Visible = False
dtpIndent1.Visible = False
Case 2:
Label2(0).Caption = "Select Item Name"
txtEmpcode.Visible = False
cmbItem.Visible = True
dtpIndent.Visible = False
Label3.Visible = False
dtpIndent1.Visible = False
Set rs = cn.Execute("select itemname from Item")
While Not rs.EOF
cmbItem.AddItem rs.Fields("itemname")
rs.MoveNext
Wend
Case 3:
Label2(0).Caption = "Select Date"
txtEmpcode.Visible = False
cmbItem.Visible = False
dtpIndent.Visible = True
Label3.Visible = False
dtpIndent1.Visible = False
Case 4:
Label2(0).Caption = " Select Month & Year"
txtEmpcode.Visible = False
cmbItem.Visible = False
dtpIndent.Visible = True
Label3.Visible = False
dtpIndent1.Visible = False
Case 5:
Label2(0).Caption = "Select Date From:"
txtEmpcode.Visible = False
cmbItem.Visible = False
dtpIndent.Visible = True
Label3.Visible = True
dtpIndent1.Visible = True
End Select
End Sub
Public Sub GridSet()
mhIndent.ColWidth(0) = 200
mhIndent.ColWidth(1) = 900
mhIndent.ColWidth(2) = 900
mhIndent.ColWidth(3) = 1500
mhIndent.ColWidth(4) = 2400
mhIndent.ColWidth(5) = 800
mhIndent.ColWidth(6) = 800
mhIndent.ColWidth(7) = 1700
mhIndent.ColWidth(8) = 1600
mhIndent.TextMatrix(0, 1) = "Date"
mhIndent.TextMatrix(0, 2) = "IndentNo"
mhIndent.TextMatrix(0, 3) = "ItemName"
mhIndent.TextMatrix(0, 4) = "Description"
mhIndent.TextMatrix(0, 5) = "Quantity"
mhIndent.TextMatrix(0, 6) = "EmpCode"
mhIndent.TextMatrix(0, 7) = "Employee Name"
mhIndent.TextMatrix(0, 8) = "Department"
End Sub

Private Sub txtEmpcode_KeyPress(KeyAscii As Integer)
Call NumberOnly(KeyAscii)
End Sub

VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form17 
   Caption         =   "Form17"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form17"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
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
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
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
      TabIndex        =   9
      Top             =   1440
      Width           =   1335
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
      TabIndex        =   8
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtRqty 
      Height          =   285
      Left            =   7440
      TabIndex        =   7
      Text            =   " "
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox txtIname 
      Height          =   285
      Left            =   7440
      TabIndex        =   6
      Text            =   " "
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtRdate 
      Height          =   285
      Left            =   7440
      TabIndex        =   5
      Text            =   " "
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtIbreif 
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Text            =   " "
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox txtEmpid 
      Height          =   285
      Left            =   2400
      TabIndex        =   3
      Text            =   " "
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox txtIndentno 
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   " "
      Top             =   1440
      Width           =   2655
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mhIndent 
      Height          =   2415
      Left            =   480
      TabIndex        =   0
      Top             =   3600
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   4260
      _Version        =   393216
      BackColor       =   16777215
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00400040&
      BorderWidth     =   5
      X1              =   -120
      X2              =   11880
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "INDENT UPDATE DELETE FORM"
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
      Left            =   1920
      TabIndex        =   17
      Top             =   0
      Width           =   8055
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Requested Quantity"
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
      Left            =   5280
      TabIndex        =   16
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Indent No."
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
      Left            =   480
      TabIndex        =   15
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Code"
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
      Left            =   480
      TabIndex        =   14
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label11 
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
      Left            =   480
      TabIndex        =   13
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Request Date"
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
      Left            =   5280
      TabIndex        =   12
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label9 
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
      Left            =   5280
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
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
      Left            =   3480
      TabIndex        =   1
      Top             =   3240
      Width           =   5295
   End
End
Attribute VB_Name = "Form17"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag As Boolean

Private Sub cmdDelete_Click()
If mhIndent.Rows = 2 Then
mhIndent.Rows = 3
End If
cn.Execute ("delete from ItemRequest where rno = " & Val(txtIndentno) & "  ")
mhIndent.RemoveItem (mhIndent.RowSel)
cmdDelete.Enabled = False
MsgBox "Record Deleted", vbInformation, "Update/Delete"
End Sub

Private Sub cmdUpdate_Click()
mhIndent.TextMatrix(mhIndent.RowSel, 2) = txtIndentno
mhIndent.TextMatrix(mhIndent.RowSel, 1) = Format(txtRdate, "dd-MMM-yy")
mhIndent.TextMatrix(mhIndent.RowSel, 6) = txtEmpid
mhIndent.TextMatrix(mhIndent.RowSel, 3) = txtIname
mhIndent.TextMatrix(mhIndent.RowSel, 4) = txtIbreif
mhIndent.TextMatrix(mhIndent.RowSel, 5) = txtRqty
Set rs1 = cn.Execute("select empname,dept from Employ where empcode = " & Val(txtEmpid) & " ")
mhIndent.TextMatrix(mhIndent.RowSel, 7) = rs1.Fields("empname")
mhIndent.TextMatrix(mhIndent.RowSel, 8) = rs1.Fields("dept")
Set rs = cn.Execute("select indentno from ItemRequest where rno = " & Val(txtIndentno) & " ")
cn.Execute ("update Indent set request_date = #" & Format(txtRdate, "dd-MMM-yy") & "#,empcode = " & Val(txtEmpid) & " where indentno = " & rs.Fields("indentno") & " ")
sql = "update ItemRequest set itemname = '" & txtIname & "',breif_r = '" & txtIbreif & "',qty_request = " & Val(txtRqty) & " where rno = " & Val(txtIndentno) & " "
cn.Execute sql
cmdUpdate.Enabled = False
MsgBox "Record Updated", vbInformation, "Update/Delete"
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call Connect
sql = "SELECT format(Indent.request_date,'dd-MMM-yy'),ItemRequest.rno,ItemRequest.itemname,ItemRequest.breif_r,ItemRequest.qty_request, Employ.empcode,Employ.empname,Employ.dept"
sql = sql & "  FROM (Indent INNER JOIN ItemRequest ON Indent.indentno = ItemRequest.indentno) INNER JOIN Employ ON Indent.empcode = Employ.empcode where ItemRequest.consider = false order by Indent.request_date;"
Set rs = cn.Execute(sql)
Set mhIndent.DataSource = rs
Call GridSet
End Sub
Private Sub mhIndent_Click()
If mhIndent.Row <> 0 And Val(mhIndent.TextMatrix(mhIndent.Row, 2)) <> 0 Then
 txtIndentno = mhIndent.TextMatrix(mhIndent.Row, 2)
 txtRdate = Format(mhIndent.TextMatrix(mhIndent.Row, 1), "dd-MMM-yy")
 txtEmpid = mhIndent.TextMatrix(mhIndent.Row, 6)
 txtIname = mhIndent.TextMatrix(mhIndent.Row, 3)
 txtIbreif = mhIndent.TextMatrix(mhIndent.Row, 4)
 txtRqty = mhIndent.TextMatrix(mhIndent.Row, 5)
 cmdDelete.Enabled = True
 cmdUpdate.Enabled = True
End If
End Sub

Private Sub txtEmpcode_KeyPress(KeyAscii As Integer)
Call NumberOnly(KeyAscii)
End Sub

Private Sub txtEmpid_KeyPress(KeyAscii As Integer)
Call NumberOnly(KeyAscii)
flag = True
End Sub

Private Sub txtIndentno_KeyPress(KeyAscii As Integer)
flag = True
End Sub

Private Sub txtRdate_KeyPress(KeyAscii As Integer)
flag = True
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
Private Sub txtRqty_KeyPress(KeyAscii As Integer)
Call NumberOnly(KeyAscii)
End Sub


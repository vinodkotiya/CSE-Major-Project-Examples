VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   MDIChild        =   -1  'True
   Picture         =   "Supplier.frx":0000
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0080FF80&
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
      Left            =   10200
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdReset 
      BackColor       =   &H0080FF80&
      Caption         =   "RESET"
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
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4320
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mhSupl 
      Height          =   2055
      Left            =   360
      TabIndex        =   23
      Top             =   5640
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   10
      FixedCols       =   0
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
      _Band(0).Cols   =   10
   End
   Begin VB.CommandButton cmdUpdate 
      BackColor       =   &H0080FF80&
      Caption         =   "UPDATE"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H0080FF80&
      Caption         =   "DELETE"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H0080FF80&
      Caption         =   "ADD"
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
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4320
      Width           =   1455
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   8640
      TabIndex        =   9
      Text            =   " "
      Top             =   3720
      Width           =   2895
   End
   Begin VB.TextBox txtFax 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2880
      TabIndex        =   8
      Text            =   " "
      Top             =   3720
      Width           =   3135
   End
   Begin VB.TextBox txtPh2 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   8640
      TabIndex        =   7
      Text            =   " "
      Top             =   3240
      Width           =   2895
   End
   Begin VB.TextBox txtPh1 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2880
      TabIndex        =   6
      Text            =   " "
      Top             =   3240
      Width           =   3135
   End
   Begin VB.TextBox txtCity 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   8640
      TabIndex        =   5
      Text            =   " "
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox txtAddress 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2880
      TabIndex        =   4
      Text            =   " "
      Top             =   2760
      Width           =   3135
   End
   Begin VB.TextBox txtCtitle 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   8640
      TabIndex        =   3
      Text            =   " "
      Top             =   2280
      Width           =   2895
   End
   Begin VB.TextBox txtCname 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Text            =   " "
      Top             =   2280
      Width           =   3135
   End
   Begin VB.TextBox txtSname 
      BackColor       =   &H00C0FFC0&
      Height          =   285
      Left            =   8640
      TabIndex        =   1
      Text            =   " "
      Top             =   1800
      Width           =   2895
   End
   Begin VB.TextBox txtSid 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   " "
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   5
      X1              =   0
      X2              =   11880
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER REGISTRATION FORM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   1680
      TabIndex        =   27
      Top             =   120
      Width           =   8655
   End
   Begin VB.Label Label11 
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
      ForeColor       =   &H00FFFF00&
      Height          =   495
      Left            =   3240
      TabIndex        =   24
      Top             =   5160
      Width           =   5895
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "E Mail"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   6240
      TabIndex        =   19
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No-2"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   6240
      TabIndex        =   17
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No-1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   360
      TabIndex        =   16
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "City Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   6240
      TabIndex        =   15
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person Title"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   6240
      TabIndex        =   13
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Person Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   6240
      TabIndex        =   11
      Top             =   1800
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1800
      Width           =   1695
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nextid As Integer

Private Sub cmdAdd_Click()
sql = " insert into supplier values(" & Val(txtSid) & ",'" & txtSname & "', "
sql = sql & " '" & txtCname & "','" & txtCtitle & "','" & txtAddress & "','" & txtCity & "',  "
sql = sql & " " & Val(txtPh1) & "," & Val(txtPh2) & "," & Val(txtFax) & ",'" & txtEmail & "') "
cn.Execute sql
cmdReset_Click
Call Setting
Form_Load
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
If mhSupl.Rows = 2 Then
mhSupl.Rows = 3
End If
If mhSupl.Rows > 2 Then
mhSupl.RemoveItem (mhSupl.RowSel)
cn.Execute ("delete from supplier where supplier_id = " & Val(txtSid) & " ")
cmdReset_Click
Call Setting
Form_Load
End If
End Sub

Private Sub cmdReset_Click()
Me.txtAddress = ""
Me.txtCity = ""
Me.txtCname = ""
Me.txtCtitle = ""
Me.txtEmail = ""
Me.txtFax = ""
Me.txtPh1 = ""
Me.txtPh2 = ""
Me.txtSid = ""
Me.txtSname = ""
End Sub

Private Sub cmdUpdate_Click()
mhSupl.TextMatrix(mhSupl.Row, 0) = txtSid
mhSupl.TextMatrix(mhSupl.Row, 1) = txtSname
mhSupl.TextMatrix(mhSupl.Row, 2) = txtCname
mhSupl.TextMatrix(mhSupl.Row, 3) = txtCtitle
mhSupl.TextMatrix(mhSupl.Row, 4) = txtAddress
mhSupl.TextMatrix(mhSupl.Row, 5) = txtCity
mhSupl.TextMatrix(mhSupl.Row, 6) = txtPh1
mhSupl.TextMatrix(mhSupl.Row, 7) = txtPh2
mhSupl.TextMatrix(mhSupl.Row, 8) = txtFax
mhSupl.TextMatrix(mhSupl.Row, 9) = txtEmail
cn.Execute ("delete from supplier where supplier_id = " & Val(txtSid) & " ")
cmdAdd_Click
End Sub

Private Sub Form_Load()
Call Connect
Set rs = cn.Execute("select max(supplier_id) as sid from supplier")
If IsNull(rs.Fields("sid")) Then
nextid = 1
Else
nextid = rs.Fields("sid") + 1
End If
txtSid = nextid
cmdAdd.Enabled = False
cmdDelete.Enabled = False
cmdUpdate.Enabled = False
Call Setting
End Sub

Private Sub mhSupl_Click()
If MDIForm1.Flag2 = False Then
txtSid = mhSupl.TextMatrix(mhSupl.Row, 0)
txtSname = mhSupl.TextMatrix(mhSupl.Row, 1)
txtCname = mhSupl.TextMatrix(mhSupl.Row, 2)
txtCtitle = mhSupl.TextMatrix(mhSupl.Row, 3)
txtAddress = mhSupl.TextMatrix(mhSupl.Row, 4)
txtCity = mhSupl.TextMatrix(mhSupl.Row, 5)
txtPh1 = mhSupl.TextMatrix(mhSupl.Row, 6)
txtPh2 = mhSupl.TextMatrix(mhSupl.Row, 7)
txtFax = mhSupl.TextMatrix(mhSupl.Row, 8)
txtEmail = mhSupl.TextMatrix(mhSupl.Row, 9)
cmdDelete.Enabled = True
cmdUpdate.Enabled = True
End If
End Sub

Private Sub txtSid_GotFocus()
txtSname.SetFocus
End Sub
Private Sub txtSname_LostFocus()
cmdAdd.Enabled = True
End Sub
Public Sub Setting()
mhSupl.ColWidth(0) = 700
mhSupl.ColWidth(1) = 1700
mhSupl.ColWidth(2) = 1700
mhSupl.ColWidth(3) = 1700
mhSupl.ColWidth(4) = 1700
mhSupl.ColWidth(5) = 1500
mhSupl.ColWidth(6) = 900
mhSupl.ColWidth(7) = 900
mhSupl.ColWidth(8) = 1200
mhSupl.ColWidth(9) = 1500
Set rs = cn.Execute("select * from supplier order by supplier_name")
Set mhSupl.DataSource = rs
mhSupl.TextMatrix(0, 0) = "ID"
mhSupl.TextMatrix(0, 1) = "Supplier Name"
mhSupl.TextMatrix(0, 2) = "Contact Person"
mhSupl.TextMatrix(0, 3) = "Contact Title"
mhSupl.TextMatrix(0, 4) = "Address"
mhSupl.TextMatrix(0, 5) = "City"
mhSupl.TextMatrix(0, 6) = "Phone"
mhSupl.TextMatrix(0, 7) = "Phone"
mhSupl.TextMatrix(0, 8) = "Fax"
mhSupl.TextMatrix(0, 9) = "E-Mail"
End Sub

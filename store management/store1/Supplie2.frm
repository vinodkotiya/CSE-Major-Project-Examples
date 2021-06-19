VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   Picture         =   "SupplierSearch.frx":0000
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00FFC0FF&
      Caption         =   "ALL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC0FF&
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0FF&
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mhSupl 
      Height          =   2895
      Left            =   480
      TabIndex        =   5
      Top             =   3240
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   5106
      _Version        =   393216
      BackColor       =   8454016
      Cols            =   10
      BackColorFixed  =   8454143
      BackColorBkg    =   12648384
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   10
   End
   Begin VB.TextBox txtSupplyid 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   " "
      Top             =   2400
      Width           =   1815
   End
   Begin VB.ComboBox cmbSupply 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFC0C0&
      BorderWidth     =   5
      X1              =   0
      X2              =   11880
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
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
      ForeColor       =   &H0080FF80&
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER PROFILE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag As Boolean

Private Sub cmbSupply_Click()
Set rs1 = cn.Execute("select supplier_id from Supplier where supplier_name = '" & cmbSupply & "' ")
txtSupplyid = rs1.Fields("supplier_id")
End Sub

'Private Sub cmbSupply_LostFocus()
'Set rs1 = cn.Execute("select supplier_id from Supplier where supplier_name = '" & cmbSupply & "' ")
'txtSupplyid = rs1.Fields("supplier_id")
'End Sub

Private Sub cmdAll_Click()
cmbSupply.Clear
txtSupplyid = ""
flag = False
Call Setting
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub Command1_Click()
flag = True
Call Setting
End Sub

Private Sub Form_Load()
Call Connect
Call Setting
flag = False
Set rs = cn.Execute("select supplier_name from Supplier")
While Not rs.EOF
cmbSupply.AddItem rs.Fields("supplier_name")
rs.MoveNext
Wend
End Sub

Public Sub Setting()
mhSupl.ColWidth(0) = 300
mhSupl.ColWidth(1) = 700
mhSupl.ColWidth(2) = 1600
mhSupl.ColWidth(3) = 1600
mhSupl.ColWidth(4) = 1600
mhSupl.ColWidth(5) = 1800
mhSupl.ColWidth(6) = 900
mhSupl.ColWidth(7) = 900
mhSupl.ColWidth(8) = 1200
mhSupl.ColWidth(9) = 1500
If flag = True Then
Set rs = cn.Execute("select * from supplier where supplier_id = " & Val(txtSupplyid) & " ")
Set mhSupl.DataSource = rs
Else
Set rs = cn.Execute("select * from supplier order by supplier_name")
Set mhSupl.DataSource = rs
End If
mhSupl.TextMatrix(0, 1) = "Supp-ID"
mhSupl.TextMatrix(0, 2) = "Supplier Name"
mhSupl.TextMatrix(0, 3) = "Contact Person"
mhSupl.TextMatrix(0, 4) = "Contact Title"
mhSupl.TextMatrix(0, 5) = "Address"
mhSupl.TextMatrix(0, 6) = "City"
mhSupl.TextMatrix(0, 7) = "Phone"
mhSupl.TextMatrix(0, 8) = "Phone"
mhSupl.TextMatrix(0, 9) = "Fax"
mhSupl.TextMatrix(0, 10) = "E-Mail"
End Sub


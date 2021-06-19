VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
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
      Height          =   375
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton cmdAll 
      BackColor       =   &H00C0C0C0&
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
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00C0C0C0&
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
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mhEmploy 
      Height          =   3495
      Left            =   840
      TabIndex        =   2
      Top             =   2640
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   6165
      _Version        =   393216
      BackColor       =   16777152
      Cols            =   7
      FixedCols       =   0
      BackColorFixed  =   16761024
      BackColorBkg    =   16777152
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin VB.TextBox txtEmpcode 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3360
      TabIndex        =   0
      Text            =   " "
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00400000&
      BorderWidth     =   5
      X1              =   0
      X2              =   11880
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE PROFILE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   615
      Left            =   3600
      TabIndex        =   4
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE CODE"
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
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   2295
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag As Boolean

Private Sub cmdAll_Click()
txtEmpcode = ""
flag = False
Call Setting
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
flag = True
Call Setting
End Sub

Private Sub Form_Load()
Call Connect
Call Setting
flag = False
End Sub
Public Sub Setting()
mhEmploy.ColWidth(0) = 1000
mhEmploy.ColWidth(1) = 2500
mhEmploy.ColWidth(2) = 1800
mhEmploy.ColWidth(3) = 1600
mhEmploy.ColWidth(4) = 1300
mhEmploy.ColWidth(5) = 2500
If flag = True Then
Set rs = cn.Execute("select * from Employ where empcode = " & Val(txtEmpcode) & " ")
Set mhEmploy.DataSource = rs
Else
Set rs = cn.Execute("select * from Employ order by empname")
Set mhEmploy.DataSource = rs
End If
mhEmploy.TextMatrix(0, 0) = "EmpCode"
mhEmploy.TextMatrix(0, 1) = "Employee Name"
mhEmploy.TextMatrix(0, 2) = "Designation"
mhEmploy.TextMatrix(0, 3) = "Department"
mhEmploy.TextMatrix(0, 4) = "Intercom No."
mhEmploy.TextMatrix(0, 5) = "E-Mail"
End Sub


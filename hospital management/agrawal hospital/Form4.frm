VERSION 5.00
Begin VB.Form addservice 
   BackColor       =   &H00FFC0C0&
   Caption         =   "ADDING NEW SERVICE"
   ClientHeight    =   5790
   ClientLeft      =   3150
   ClientTop       =   2145
   ClientWidth     =   5895
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   LockControls    =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   480
      TabIndex        =   12
      Top             =   4320
      Width           =   5175
      Begin VB.CommandButton cmdcat 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtcat 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Category Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.ComboBox cmbcat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form4.frx":0000
      Left            =   3000
      List            =   "Form4.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
   Begin VB.TextBox txtetem 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3000
      TabIndex        =   1
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtrate 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3000
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4380
      TabIndex        =   5
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   300
      TabIndex        =   3
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdcancle 
      Caption         =   "Cancle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2340
      TabIndex        =   4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add Service form"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   615
      Left            =   600
      TabIndex        =   11
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Select Catagory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Item Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Rate/Per"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   8
      Top             =   2280
      Width           =   1575
   End
End
Attribute VB_Name = "addservice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim d, m As Integer

Private Sub cmdadd_Click()
If cmbcat.Text = "" Or txtrate = "" Or txtetem = "" Then
MsgBox "PLEASE ENTER THE VALUES", vbOKOnly, "USER INFORMATION"
cmbcat.SetFocus
Exit Sub
End If
rs.AddNew
d = rs.RecordCount
rs!itemno = (d + 1)
rs!itemname = "" & txtetem
rs!Rate = Val(txtrate)
rs!category = "" & cmbcat.Text
rs.Update
txtetem.Text = ""
txtrate.Text = ""
End Sub
Private Sub Cmdcancle_Click()
txtetem.Text = ""
txtrate.Text = ""
End Sub
Private Sub CMDEXIT_Click()

main.Enabled = True

If rs.State = 1 Then rs.Close
If rs1.State = 1 Then rs1.Close
If cn.State = 1 Then cn.Close
Unload Me
Load main
main.Show
End Sub

Private Sub cmdcat_Click()
If txtcat = "" Then
MsgBox "enter category name", vbOKOnly
txtcat.SetFocus

Exit Sub
End If
rs1.AddNew
rs1!cat = "" & txtcat
rs1.Update
txtcat = ""
rs2cmb
End Sub

Private Sub Form_Load()
main.Enabled = False

cn.CursorLocation = adUseClient
cn.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\agrawal.mdb;Persist Security Info=False"
If rs.State = 1 Then rs.Close
rs.Open "select * from service_master", cn, adOpenDynamic, adLockOptimistic
If rs1.State = 1 Then rs1.Close
rs1.Open "select cat from catgri", cn, adOpenDynamic, adLockOptimistic
rs2cmb
End Sub

Private Sub Form_Unload(Cancel As Integer)
If rs.State = 1 Then rs.Close
If rs1.State = 1 Then rs1.Close
If cn.State = 1 Then cn.Close
End Sub

Private Sub txtrate_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
MsgBox "PLEASE ENTER CORRECT VALUE IN RATE/PER", vbOKOnly, "USER INFORMATION"
KeyAscii = 0
txtrate = ""
Exit Sub
End If
End Sub


Private Sub rs2cmb()
If rs1.RecordCount > 0 And Not (rs1.BOF Or rs1.EOF) Then
rs1.MoveFirst
While Not (rs1.EOF)
cmbcat.AddItem rs1!cat
rs1.MoveNext
Wend
End If
Exit Sub

End Sub

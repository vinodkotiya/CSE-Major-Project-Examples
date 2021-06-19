VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form payment 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Recieved  payment form"
   ClientHeight    =   6375
   ClientLeft      =   660
   ClientTop       =   1425
   ClientWidth     =   10320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   ScaleHeight     =   6375
   ScaleWidth      =   10320
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtpay 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6173
      TabIndex        =   4
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
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
      Left            =   8273
      TabIndex        =   7
      Top             =   5460
      Width           =   1335
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "CANCLE"
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
      Left            =   4493
      TabIndex        =   6
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "SAVE"
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
      Left            =   713
      TabIndex        =   5
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   3773
      TabIndex        =   17
      Top             =   1080
      Width           =   6255
      Begin VB.TextBox txtward 
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
         Left            =   1080
         TabIndex        =   12
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtage 
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
         Left            =   1080
         TabIndex        =   11
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtname 
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
         Left            =   1080
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtpre 
         Height          =   375
         Left            =   4680
         TabIndex        =   8
         Top             =   2280
         Width           =   1095
      End
      Begin VB.TextBox txtpcode 
         Height          =   375
         Left            =   1800
         TabIndex        =   21
         Top             =   600
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.TextBox txtsex 
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
         Left            =   4680
         TabIndex        =   9
         Top             =   1455
         Width           =   1095
      End
      Begin VB.TextBox txtdoa 
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
         Left            =   4680
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ward"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Name"
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
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Age"
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
         Left            =   120
         TabIndex        =   25
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Previous Balance"
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
         Left            =   2760
         TabIndex        =   23
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sex"
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
         Left            =   2760
         TabIndex        =   19
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Date of Admission"
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
         Left            =   2760
         TabIndex        =   18
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Patient selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   293
      TabIndex        =   14
      Top             =   1080
      Width           =   3375
      Begin VB.CommandButton cmdclear 
         Caption         =   "CLEAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdsearch 
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtsearch 
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
         Left            =   1800
         TabIndex        =   1
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cmbsearch 
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
         ItemData        =   "PAYREC.frx":0000
         Left            =   1800
         List            =   "PAYREC.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin MSDataGridLib.DataGrid dgsearch 
         Height          =   1215
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Search "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Select Serch Option"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Recieved   payment   Rs."
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
      Left            =   2573
      TabIndex        =   24
      Top             =   4680
      Width           =   2655
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Recieved Payment form"
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
      Height          =   735
      Left            =   2393
      TabIndex        =   22
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public z As Integer

Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim st As String
Dim i As Long


Private Sub Cmdsave_Click()
If txtpay = "" Then
    MsgBox "ENTER AMOUNT ", vbOKOnly
    Exit Sub
End If
txt2rs1
rs12txt
End Sub
Private Sub Cmdcancel_Click()
txtempty
End Sub
Private Sub CMDEXIT_Click()
main.Enabled = True

If z = 1 Then
If rs.State = 1 Then rs.Close
If rs1.State = 1 Then rs1.Close
If rs2.State = 1 Then rs2.Close
If cn.State = 1 Then cn.Close
Unload Me
Load billing
billing.Show
billing.Cmdok.SetFocus

Else
rs.Close
rs1.Close
rs2.Close
cn.Close
Unload Me
Load main
main.Show
End If
End Sub
Private Sub cmdsearch_Click()
If Cmbsearch.Text = "" Then
MsgBox "PLEASE SELECT  FIELD TO SEARCH", vbInformation, "USER INFORMATION"
Exit Sub
End If
If txtsearch = "" Then
MsgBox "PLEASE ENTER TEXT TO SEARCH", vbInformation, "USER INFORMATION"
Exit Sub
End If
If txtsearch = "" And Cmbsearch.Text = "pcode" Then
        MsgBox "Please Enter code to Search", vbInformation, "User Information"
        txtsearch.SetFocus
  ElseIf txtsearch = "" And Cmbsearch.ListIndex = 1 Then
  MsgBox "Please Enter name to Search", vbInformation, "User Information"
        txtsearch.SetFocus
Else
        Set dgsearch.DataSource = Nothing
        txtempty
        rs.Filter = st & " like '" & txtsearch & "*'"
        Set dgsearch.DataSource = rs
        If rs.RecordCount = 0 Then
            MsgBox "No Records Found", vbInformation, "User Information"
            cmdclear_Click
        Else
              rs2text
        End If
    End If
End Sub

Private Sub cmdclear_Click()
txtsearch.Text = ""

st = ""
    rs.Filter = 0
    rs.Requery
    Set dgsearch.DataSource = rs
  
End Sub

Private Sub dgsearch_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
rs2text
End Sub

Private Sub Form_Load()
main.Enabled = False

cn.CursorLocation = adUseClient
If cn.State = 1 Then cn.Close
cn.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\agrawal.mdb;Persist Security Info=False"
If rs.State = 1 Then rs.Close
rs.Open "select * from patient", cn, adOpenDynamic, adLockOptimistic
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from amt_paid", cn, adOpenDynamic, adLockOptimistic
If rs2.State = 1 Then rs2.Close
rs2.Open "select pcode,category from bed_tra", cn, adOpenDynamic, adLockOptimistic
Set dgsearch.DataSource = rs
rs2text
End Sub
Private Sub cmbsearch_Click()
Select Case Cmbsearch.ListIndex
 Case 0
  st = "pcode"
 Case 1
  st = "pname"
 Case Else
  st = ""
 End Select
End Sub
Private Sub txtempty()
Txtname = ""
txtage = ""
txtsex = ""
Txtdoa = ""
txtpcode = ""
txtward = ""
txtpay = ""
End Sub
Private Sub rs2text()
If rs.RecordCount > 0 And rs2.RecordCount > 0 Then
rs.MoveFirst
Txtname = rs!pname
Txtdoa = rs!doa
txtage = rs!age
txtsex = rs!sex
txtpcode = rs!pcode
rs2.Filter = "pcode" & " like '" & txtpcode & "*'"
If rs2.RecordCount > 0 Then
rs2.MoveLast
txtward = rs2!category
Else
txtempty
End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If rs.State = 1 Then rs.Close
If rs1.State = 1 Then rs1.Close
If rs2.State = 1 Then rs2.Close
If cn.State = 1 Then cn.Close
End Sub

Private Sub txtpcode_Change()
rs12txt
End Sub

Private Sub txtsearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdsearch_Click
End If
End Sub
Private Sub txt2rs1()
If txtpcode = "" Then
Exit Sub
End If
rs1.AddNew
rs1!pcode = txtpcode & ""
rs1!amtpaid = Val(txtpay)
rs1!Date = Date
rs1.Update
txtpay = ""
End Sub
Private Sub rs12txt()
If txtpcode = "" Then
txtpre = 0
Else
rs1.Filter = "pcode" & " like '" & txtpcode & "*'"
i = 0
If rs1.RecordCount = 0 Then
txtpre = 0

Exit Sub
End If
rs1.MoveFirst
While Not (rs1.EOF)
i = (i + rs1!amtpaid)
rs1.MoveNext
Wend
txtpre = i
End If
End Sub

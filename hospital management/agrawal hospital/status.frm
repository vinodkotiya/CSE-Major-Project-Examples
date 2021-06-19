VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form status 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Status form"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10050
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   Picture         =   "status.frx":0000
   ScaleHeight     =   6795
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport Crpt 
      Left            =   840
      Top             =   5280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton cmddetail 
      Caption         =   "DETAIL "
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
      Left            =   4418
      TabIndex        =   4
      Top             =   6240
      Width           =   1215
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
      Height          =   375
      Left            =   7178
      TabIndex        =   6
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
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
      Left            =   1658
      TabIndex        =   5
      Top             =   6240
      Width           =   1215
   End
   Begin VB.TextBox txtamt 
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
      Left            =   6030
      TabIndex        =   25
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox txtexp 
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
      Left            =   6030
      TabIndex        =   24
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Height          =   3135
      Left            =   3758
      TabIndex        =   12
      Top             =   1080
      Width           =   6015
      Begin VB.TextBox txtpcode 
         Height          =   375
         Left            =   2400
         TabIndex        =   18
         Text            =   "Text11"
         Top             =   600
         Visible         =   0   'False
         Width           =   150
      End
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
         Left            =   2040
         TabIndex        =   17
         Top             =   2160
         Width           =   1335
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
         Left            =   2040
         TabIndex        =   16
         Top             =   1440
         Width           =   1335
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
         Left            =   4920
         TabIndex        =   15
         Top             =   1320
         Width           =   855
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
         Left            =   4920
         TabIndex        =   14
         Top             =   600
         Width           =   855
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
         Height          =   405
         Left            =   2040
         TabIndex        =   13
         Top             =   600
         Width           =   1335
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
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   2160
         Width           =   1695
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
         Left            =   240
         TabIndex        =   22
         Top             =   1440
         Width           =   1695
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
         Left            =   3960
         TabIndex        =   21
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
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
         Left            =   3960
         TabIndex        =   20
         Top             =   600
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
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Patient Selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   278
      TabIndex        =   8
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
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Top             =   1200
         Width           =   1095
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
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
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
         ItemData        =   "status.frx":0342
         Left            =   1920
         List            =   "status.frx":034C
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
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
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid dgsearch 
         Height          =   1215
         Left            =   0
         TabIndex        =   9
         Top             =   1800
         Width           =   3375
         _ExtentX        =   5953
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
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Search "
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
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Select Serch Option"
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
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Total Amount Given by Patient"
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
      Left            =   2685
      TabIndex        =   27
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Total expence of Patient"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2685
      TabIndex        =   26
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Patient Status"
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
      Left            =   3278
      TabIndex        =   7
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim rs5 As New ADODB.Recordset
Dim st As String
Dim a As Integer
Dim i As Integer
Dim z As Integer
Dim c As Integer
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
Private Sub CMDEXIT_Click()
main.Enabled = True

If rs.State = 1 Then rs.Close
If rs1.State = 1 Then rs1.Close
If rs2.State = 1 Then rs2.Close
If rs3.State = 1 Then rs3.Close
If rs4.State = 1 Then rs4.Close
If rs5.State = 1 Then rs5.Close
If cn.State = 1 Then cn.Close
Unload Me
Load main
main.Show
End Sub
Private Sub Cmddetail_Click()
Dim str As String
Crpt.Reset
Crpt.DiscardSavedData = True
Crpt.ReportFileName = App.Path & "\rep4.rpt"
str = " {patient.pcode} = '" & txtpcode & "'"
Crpt.ReplaceSelectionFormula str
Crpt.WindowState = crptMaximized
If MsgBox(" CLICK YES FOR PRINT OR NO FOR PREVIEW", vbYesNo, "USER INFORMATION") = vbYes Then

Crpt.Destination = crptToPrinter
On Error GoTo e1
Else
Crpt.Destination = crptToWindow
End If
Crpt.Action = 1


Exit Sub
e1: MsgBox "THERE IS NO DEFAULT PRINTER AVAILABLE", vbCritical, "PRINTER ERROR"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If rs.State = 1 Then rs.Close
If rs1.State = 1 Then rs1.Close
If rs2.State = 1 Then rs2.Close
If rs3.State = 1 Then rs3.Close
If rs4.State = 1 Then rs4.Close
If rs5.State = 1 Then rs5.Close
If cn.State = 1 Then cn.Close
End Sub

Private Sub txtpcode_Change()
If txtpcode = "" Then
txtamt = 0
txtexp = 0
Exit Sub
Else
rs1.Filter = "pcode" & " like '" & txtpcode & "*'"
If rs1.RecordCount = 0 Then
txtamt.Text = 0
rs4.Filter = "pcode" & " like '" & txtpcode & "*'"
If rs4.RecordCount = 0 Then
txtexp.Text = 0

End If
End If
On Error Resume Next
rs1.MoveFirst
a = 0
While Not (rs1.EOF)
a = a + rs1!amtpaid
rs1.MoveNext
Wend
txtamt.Text = a
End If

i = 0
On Error Resume Next
rs4.MoveFirst
While Not (rs4.EOF)
i = i + rs4!qr
rs4.MoveNext
Wend
txtexp.Text = i


z = 0
rs5.Filter = "pcode" & " like '" & txtpcode & " '"
rs5.MoveFirst
While Not (rs5.EOF)
If rs5!ndays >= 0 Then
If rs5!ndays = 0 Then
c = 1
Else
c = rs5!ndays
End If
z = (z + (c) * (rs5!Rate))
rs5.MoveNext
End If
Wend
txtexp.Text = Val(txtexp.Text) + z
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
rs2.Open "select rate from service_master where itemno in (select itemno from treatment where pcode= '" & txtpcode.Text & "')", cn, adOpenDynamic, adLockOptimistic
If rs3.State = 1 Then rs3.Close
rs3.Open "select * from bed_tra", cn, adOpenDynamic, adLockOptimistic
If rs4.State = 1 Then rs4.Close
rs4.Open "select * from treatment", cn, adOpenDynamic, adLockOptimistic
If rs5.State = 1 Then rs5.Close
rs5.Open "select * from bed_tra", cn, adOpenStatic, adLockOptimistic
Set dgsearch.DataSource = rs
rs2text

End Sub
Private Sub rs2text()
If rs.EOF Then
txtempty
Exit Sub
End If
Txtname = rs!pname
txtage = rs!age
Txtdoa = rs!doa
txtsex = rs!sex
txtpcode = rs!pcode
rs3.Filter = "pcode" & " like '" & txtpcode & "*'"
rs3.MoveLast
txtward = rs3!category
End Sub
Private Sub txtempty()
Txtname = ""
Txtdoa = ""
txtsex = ""
txtage = ""
txtpcode = ""
txtward = ""
End Sub
Private Sub dgsearch_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If rs.RecordCount > 0 And Not (rs.BOF Or rs.EOF) Then
rs2text
End If
End Sub
Private Sub cmdsearch_Click()
If Cmbsearch.Text = "" Then
MsgBox "PLEASE SELECT  FIELD TO SEARCH", vbInformation, "USER INFORMATION"
Exit Sub
End If
If txtsearch.Text = "" Then
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
dgsearch.Enabled = True
End Sub
Private Sub txtsearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 cmdsearch_Click
 End If
 End Sub

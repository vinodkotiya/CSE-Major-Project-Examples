VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form billing 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Dischrage form"
   ClientHeight    =   6795
   ClientLeft      =   420
   ClientTop       =   1425
   ClientWidth     =   10335
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport Crpt 
      Left            =   1080
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Height          =   3135
      Left            =   3671
      TabIndex        =   19
      Top             =   960
      Width           =   6375
      Begin VB.TextBox txtdis 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   4920
         TabIndex        =   26
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox Txtname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   1920
         TabIndex        =   25
         Top             =   600
         Width           =   1215
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
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   4920
         TabIndex        =   24
         Top             =   600
         Width           =   1215
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
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   4920
         TabIndex        =   23
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Txtdoa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   1920
         TabIndex        =   22
         Top             =   1320
         Width           =   1215
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
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   1920
         TabIndex        =   21
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtpcode 
         Height          =   375
         Left            =   2040
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Bed / ward"
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
         TabIndex        =   32
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Date of admission"
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
         TabIndex        =   31
         Top             =   1320
         Width           =   1815
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
         Left            =   3360
         TabIndex        =   30
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
         Left            =   3360
         TabIndex        =   29
         Top             =   600
         Width           =   975
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
         Left            =   240
         TabIndex        =   28
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Discharge date"
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
         Left            =   3360
         TabIndex        =   27
         Top             =   2040
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Patient selection"
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
      Left            =   318
      TabIndex        =   15
      Top             =   960
      Width           =   3255
      Begin VB.CommandButton Cmdclear 
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
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Cmdsearch 
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
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox Cmbsearch 
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
         ItemData        =   "discharge.frx":0000
         Left            =   1920
         List            =   "discharge.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1215
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
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid dgsearch 
         Height          =   1215
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2143
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
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
         TabIndex        =   18
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Select search option"
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
         TabIndex        =   17
         Top             =   240
         Width           =   1815
      End
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
      Left            =   6131
      TabIndex        =   11
      Top             =   4200
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
      Left            =   6131
      TabIndex        =   10
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox txtbal 
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
      Left            =   6131
      TabIndex        =   9
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton Cmdok 
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
      Height          =   495
      Left            =   1815
      TabIndex        =   5
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton Cmdcancel 
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
      Left            =   3615
      TabIndex        =   6
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton Cmddetail 
      Caption         =   "DETAILS "
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
      Left            =   5535
      TabIndex        =   3
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton Cmdexit 
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
      Left            =   7455
      TabIndex        =   7
      Top             =   6120
      Width           =   1080
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
      Height          =   255
      Left            =   3018
      TabIndex        =   14
      Top             =   4320
      Width           =   2895
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
      Height          =   255
      Left            =   3018
      TabIndex        =   13
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Balance"
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
      Left            =   3018
      TabIndex        =   12
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Discharge form"
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
      Height          =   855
      Left            =   2595
      TabIndex        =   8
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "billing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public z As Boolean
Dim q As Boolean
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim rs5 As New ADODB.Recordset
Dim rs6 As New ADODB.Recordset
Dim rs7 As New ADODB.Recordset
Dim st As String
Dim a As Long
Dim i As Long
Dim c As Long
Dim bal As Long

Private Sub cmbsearch_Click()
Select Case cmbsearch.ListIndex
 Case 0
   st = "pcode"
 Case 1
   st = "pname"
 Case Else
   st = ""
End Select
End Sub

Private Sub cmdok_Click()
If txtpcode = "" Then
MsgBox "NO PATIENT TO DISCHARGE", vbOKOnly
Exit Sub
End If

If txtbal = "" Then
Exit Sub
End If

If txtbal.Text < 0 Then
MsgBox "THE  PATIENT  HAVE  TO  PAID  Rs.  " & -Val(txtbal.Text), vbCritical, "PAYMENT WARNING"
z = True
If rs.State = 1 Then rs.Close
If rs1.State = 1 Then rs1.Close
If rs2.State = 1 Then rs2.Close
If rs3.State = 1 Then rs3.Close
If rs4.State = 1 Then rs4.Close
If rs5.State = 1 Then rs5.Close
If rs6.State = 1 Then rs6.Close
If rs7.State = 1 Then rs7.Close
If cn.State = 1 Then cn.Close


Unload Me
Load payment
payment.Show
payment.txtpay.SetFocus
Exit Sub
End If

If MsgBox("CONFIRM PATIENT DISCHARGED", vbYesNo) = 6 Then
rs.Filter = "pcode" & " like '" & txtpcode & "*'"
rs6.AddNew
rs6!pcode = rs!pcode
rs6!pname = rs!pname
rs6!age = rs!age
rs6!sex = rs!sex
rs6!education = rs!education
rs6!address = rs!address
rs6!ph = rs!ph
rs6!toa = rs!toa
rs6!doa = rs!doa
rs6!pccomplaint = rs!pccomplaint
rs6!paname = rs!paname
rs6!ddate = Date
rs6.Update

rs1.Filter = "pcode" & " like '" & txtpcode & "*'"

If rs1.RecordCount > 0 Then
rs1.MoveFirst
While Not (rs1.EOF)
rs1.Delete
rs1.Update
rs1.MoveNext

Wend

End If

rs5.Filter = "pcode" & " like '" & txtpcode & "*'"
If rs5.RecordCount > 0 Then
rs5.MoveFirst
While Not (rs5.EOF)
rs5.Delete
rs5.Update
rs5.MoveNext
Wend

End If

rs7.Filter = "category" & " like '" & txtward & "*'"
If rs7.RecordCount > 0 Then

rs7!empty = (rs7!empty + 1)
rs7!booked = (rs7!booked - 1)
rs7.Update
End If
rs4.Filter = "pcode" & " like '" & txtpcode & "*'"
If rs4.RecordCount > 0 Then
rs4.MoveFirst
While Not (rs4.EOF)
rs4.Delete
rs4.Update
rs4.MoveNext
Wend


End If
rs.Filter = "pcode" & " like '" & txtpcode & "*'"
rs.Delete
rs.Update
rs.Requery
rs.Filter = 0
txtempty
txtamt = 0
txtexp = 0
txtbal = 0

Set dgsearch.DataSource = rs

End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
If rs.State = 1 Then rs.Close
If rs1.State = 1 Then rs1.Close
If rs2.State = 1 Then rs2.Close
If rs3.State = 1 Then rs3.Close
If rs4.State = 1 Then rs4.Close
If rs5.State = 1 Then rs5.Close
If rs6.State = 1 Then rs6.Close
If rs7.State = 1 Then rs7.Close
If cn.State = 1 Then cn.Close
End Sub

Private Sub txtsearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 cmdsearch_Click
End If
End Sub
Private Sub CMDEXIT_Click()
If rs.State = 1 Then rs.Close
If rs1.State = 1 Then rs1.Close
If rs2.State = 1 Then rs2.Close
If rs3.State = 1 Then rs3.Close
If rs4.State = 1 Then rs4.Close
If rs5.State = 1 Then rs5.Close
If rs6.State = 1 Then rs6.Close
If rs7.State = 1 Then rs7.Close
If cn.State = 1 Then cn.Close
main.Enabled = True

Unload Me
Load main
main.Show

End Sub
Private Sub Cmddetail_Click()
Dim str As String
Crpt.Reset
Crpt.DiscardSavedData = True
Crpt.ReportFileName = App.Path & "\rep2.rpt"
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
q = True
'If rs.State = 1 Then rs.Close
'If rs1.State = 1 Then rs1.Close
'If rs2.State = 1 Then rs2.Close
'If rs3.State = 1 Then rs3.Close
'If rs4.State = 1 Then rs4.Close
'If rs5.State = 1 Then rs5.Close
''If rs6.State = 1 Then rs6.Close
'If rs7.State = 1 Then rs7.Close
'If cn.State = 1 Then cn.Close

Exit Sub
e1: MsgBox "THERE IS NO DEFAULT PRINTER AVAILABLE", vbCritical, "PRINTER ERROR"

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
End If

rs4.Filter = "pcode" & " like '" & txtpcode & "*'"
If rs4.RecordCount = 0 Then
txtexp.Text = 0
txtbal = 0

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
 
 rs4.MoveFirst
i = 0
While Not (rs4.EOF)
 i = (i + rs4!qr)
 rs4.MoveNext
Wend
txtexp.Text = i

a = 0

rs5.Filter = "pcode" & " like '" & txtpcode & " '"
rs5.MoveFirst
While Not (rs5.EOF)
If rs5!ndays >= 0 Then
If rs5!ndays = 0 Then
c = 1
Else
c = rs5!ndays
End If
a = (a + (c) * (rs5!Rate))
rs5.MoveNext
End If
Wend
txtexp.Text = Val(txtexp.Text) + a
txtbal.Text = (Val(txtamt.Text) - Val(txtexp.Text))

End Sub
Private Sub Form_Load()
z = False
cn.CursorLocation = adUseClient
If cn.State = 1 Then cn.Close
cn.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\agrawal.mdb;Persist Security Info=False"
If rs.State = 1 Then rs.Close
rs.Open "select * from patient", cn, adOpenDynamic, adLockOptimistic
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from amt_paid", cn, adOpenDynamic, adLockOptimistic
If rs2.State = 1 Then rs2.Close
rs2.Open "select * from service_master", cn, adOpenDynamic, adLockOptimistic
If rs3.State = 1 Then rs3.Close
rs3.Open "select pcode,category from bed_tra", cn, adOpenDynamic, adLockOptimistic
If rs4.State = 1 Then rs4.Close
rs4.Open "select * from treatment", cn, adOpenDynamic, adLockOptimistic
If rs5.State = 1 Then rs5.Close
rs5.Open "select * from bed_tra", cn, adOpenDynamic, adLockOptimistic
If rs6.State = 1 Then rs6.Close
rs6.Open "select * from oldpatient", cn, adOpenDynamic, adLockOptimistic
If rs7.State = 1 Then rs7.Close
rs7.Open "select * from bed", cn, adOpenDynamic, adLockOptimistic
txtdis = Date
txtdis.Enabled = False
main.Enabled = False

Set dgsearch.DataSource = rs
rs2text
End Sub
Private Sub rs2text()
If rs.EOF Then
txtempty
Exit Sub
End If
txtname = rs!pname
txtage = rs!age
txtdoa = rs!doa
txtsex = rs!sex
txtpcode = rs!pcode

rs3.Filter = "pcode" & " like '" & txtpcode & "*'"
If rs3.RecordCount > 0 Then
rs3.MoveLast
txtward = rs3!category
End If
End Sub
Private Sub txtempty()
txtname = ""
txtdoa = ""
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
If cmbsearch.Text = "" Then
MsgBox "PLEASE SELECT  FIELD TO SEARCH", vbInformation, "USER INFORMATION"
Exit Sub
End If
If txtsearch = "" Then
MsgBox "PLEASE ENTER TEXT TO SEARCH", vbInformation, "USER INFORMATION"
Exit Sub
End If
If txtsearch = "" And cmbsearch.Text = "pcode" Then
MsgBox "Please Enter code to Search", vbInformation, "User Information"
        txtsearch.SetFocus
  ElseIf txtsearch = "" And cmbsearch.ListIndex = 1 Then
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



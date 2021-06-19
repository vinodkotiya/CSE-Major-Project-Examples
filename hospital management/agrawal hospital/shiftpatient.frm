VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form shiftpatient 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Shift Patient"
   ClientHeight    =   6120
   ClientLeft      =   885
   ClientTop       =   1305
   ClientWidth     =   10560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   10560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCANCLE 
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
      Left            =   4560
      TabIndex        =   6
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton Cmdupdate 
      Caption         =   "Update"
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
      Left            =   1433
      TabIndex        =   5
      Top             =   4920
      Width           =   1455
   End
   Begin VB.ComboBox Cmbcate 
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
      Left            =   8640
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      Height          =   3135
      Left            =   3593
      TabIndex        =   17
      Top             =   1080
      Width           =   6855
      Begin VB.TextBox txtrate 
         Height          =   375
         Left            =   1800
         TabIndex        =   26
         Top             =   600
         Visible         =   0   'False
         Width           =   150
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
         Left            =   1320
         TabIndex        =   12
         Top             =   600
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
         Left            =   5040
         TabIndex        =   8
         Top             =   1320
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
         Height          =   405
         Left            =   1320
         TabIndex        =   11
         Top             =   1380
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
         Left            =   5040
         TabIndex        =   9
         Top             =   600
         Width           =   1335
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
         Height          =   405
         Left            =   1320
         TabIndex        =   10
         Top             =   2160
         Width           =   1335
      End
      Begin VB.TextBox txtpcode 
         Height          =   375
         Left            =   2040
         TabIndex        =   18
         Text            =   "Text11"
         Top             =   600
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Select New Ward"
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
         Left            =   2880
         TabIndex        =   25
         Top             =   2160
         Width           =   2055
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
         TabIndex        =   23
         Top             =   720
         Width           =   735
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
         Left            =   2880
         TabIndex        =   22
         Top             =   1440
         Width           =   855
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
         Left            =   240
         TabIndex        =   21
         Top             =   1380
         Width           =   975
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
         Left            =   2880
         TabIndex        =   20
         Top             =   720
         Width           =   1935
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
         TabIndex        =   19
         Top             =   2160
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   3135
      Left            =   113
      TabIndex        =   13
      Top             =   1080
      Width           =   3375
      Begin VB.TextBox txtsearch 
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cmbsearch 
         Height          =   315
         ItemData        =   "shiftpatient.frx":0000
         Left            =   1920
         List            =   "shiftpatient.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1335
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
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
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
      Begin MSDataGridLib.DataGrid dgsearch 
         Height          =   1215
         Left            =   120
         TabIndex        =   14
         Top             =   1800
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
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Select Search option"
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
         TabIndex        =   16
         Top             =   240
         Width           =   1455
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
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CommandButton Cmdexit 
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
      Left            =   7673
      TabIndex        =   7
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Patient Shift form"
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
      Left            =   2453
      TabIndex        =   24
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "shiftpatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
'Dim rs4 As New ADODB.Recordset
Dim st As String
Dim DT As Integer
Dim dte As String
Dim d As Date

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
Private Sub Cmdcancle_Click()
empty1
End Sub
Private Sub CMDEXIT_Click()
main.Enabled = True

If rs.State = 1 Then rs.Close
If rs1.State = 1 Then rs1.Close
If rs2.State = 1 Then rs2.Close
If rs3.State = 1 Then rs3.Close
If cn.State = 1 Then cn.Close
Unload Me
Load main
main.Show
End Sub
Private Sub cmdsearch_Click()
If cmbsearch.Text = "" Then
   MsgBox "PLEASE SELECT  FIELD TO SEARCH", vbInformation, "USER INFORMATION"
   Exit Sub
End If
If txtsearch.Text = "" Then
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
Private Sub Cmdupdate_Click()

If txtward = "" Or Cmbcate.Text = "" Then
 MsgBox "INVALID SHIFTING", vbOKOnly
 Exit Sub
 End If

If txtward.Text = Cmbcate.Text Then
 MsgBox "Can not transfer in same ward", vbOKOnly
 Exit Sub
 End If
If MsgBox("CONFIRM RECORD UPDATED", vbYesNo, "USER QUESTION") = 6 Then
    rs2.Filter = "category" & " like '" & Cmbcate.Text & "*'"
  If rs2!empty = 0 Then
    MsgBox " THIS WARD IS FULL", vbOKOnly
    Cmbcate.SetFocus
    Exit Sub
    End If
    
   rs2.Filter = "category" & " like '" & txtward.Text & "'"
    rs2!booked = (rs2!booked - 1)
    rs2!empty = (rs2!empty + 1)
     
     rs2.Filter = "category" & " like '" & Cmbcate.Text & "*'"
    rs2!booked = rs2!booked + 1
    rs2!empty = rs2!empty - 1
    
    rs2.UpdateBatch
    txt2rs
    empty1
Else
 Exit Sub
End If
End Sub
Private Sub Dgsearch_Click()
If rs.RecordCount > 0 And Not (rs.BOF Or rs.EOF) Then
rs2text
End If
End Sub
Private Sub Form_Load()
main.Enabled = False

cn.CursorLocation = adUseClient
If cn.State = 1 Then cn.Close
cn.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\agrawal.mdb;Persist Security Info=False"
If rs.State = 1 Then rs.Close
rs.Open "select * from patient", cn, adOpenDynamic, adLockOptimistic
If rs1.State = 1 Then rs1.Close
rs1.Open "select * from bed_tra", cn, adOpenDynamic, adLockOptimistic
If rs2.State = 1 Then rs2.Close
rs2.Open "select * from bed", cn, adOpenDynamic, adLockOptimistic
If rs3.State = 1 Then rs3.Close
rs3.Open "select * from bed_tra ", cn, adOpenDynamic, adLockBatchOptimistic
rs2cmb
Set dgsearch.DataSource = rs
rs2text
End Sub
Private Sub empty1()
txtname.Text = ""
txtage.Text = ""
txtsex.Text = ""
txtward.Text = ""
txtdoa.Text = ""
End Sub
Private Sub rs2cmb()
If rs2.RecordCount > 0 And Not (rs2.BOF Or rs2.EOF) Then
rs2.MoveFirst
While Not (rs2.EOF)
Cmbcate.AddItem rs2!category
rs2.MoveNext
Wend
End If
Exit Sub
End Sub
Private Sub txtempty()
txtname = ""
txtdoa = ""
txtsex = ""
txtage = ""
txtpcode = ""
txtward = ""
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
If rs3.State = 1 Then rs3.Close
rs3.Open "select category from bed_tra where pcode ='" & txtpcode.Text & "' ", cn, adOpenDynamic, adLockBatchOptimistic
rs3.MoveLast
txtward = rs3!category

End Sub

Private Sub Form_Unload(Cancel As Integer)
If rs.State = 1 Then rs.Close
If rs1.State = 1 Then rs1.Close
If rs2.State = 1 Then rs2.Close
If rs3.State = 1 Then rs3.Close
If cn.State = 1 Then cn.Close
End Sub

Private Sub txtsearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdsearch_Click
End If
End Sub
Private Sub txt2rs()
rs1.AddNew
rs1!pcode = txtpcode & ""
rs1!tos = Time
rs1!Date = Date
 rs2.Filter = "category" & " like '" & Cmbcate.Text & "*'"
rs1!Rate = rs2!Rate
rs1!category = Cmbcate.Text & ""
If rs3.State = 1 Then rs3.Close
rs3.Open "select * from bed_tra where pcode ='" & txtpcode.Text & "' ", cn, adOpenDynamic, adLockBatchOptimistic
If rs3.RecordCount > 0 Then
rs3.MoveLast
d = rs3!Date
rs1!ndays = DateDiff("d", d, Date)
rs1.Update
Else
rs1!ndays = -1
rs1.Update
End If
End Sub

VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form updpatient 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Udate Paient Form"
   ClientHeight    =   7500
   ClientLeft      =   1125
   ClientTop       =   825
   ClientWidth     =   9645
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   9645
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtdoa 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6157
      TabIndex        =   11
      Top             =   4395
      Width           =   1455
   End
   Begin VB.TextBox txtedu 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6157
      TabIndex        =   10
      Top             =   3870
      Width           =   1455
   End
   Begin VB.TextBox txtrfa 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6157
      TabIndex        =   8
      Top             =   2805
      Width           =   3015
   End
   Begin VB.TextBox txtadd 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6157
      TabIndex        =   7
      Top             =   2280
      Width           =   3015
   End
   Begin VB.TextBox txtatten 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6157
      TabIndex        =   6
      Top             =   1800
      Width           =   3015
   End
   Begin VB.TextBox txtname 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6120
      TabIndex        =   5
      Top             =   1080
      Width           =   3015
   End
   Begin VB.TextBox txtpcode 
      Height          =   375
      Left            =   6232
      TabIndex        =   34
      Top             =   1080
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
      Height          =   360
      Left            =   6157
      TabIndex        =   9
      Top             =   3345
      Width           =   1455
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
      Height          =   360
      Left            =   6157
      TabIndex        =   13
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox txag 
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
      Left            =   6157
      TabIndex        =   12
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox txp 
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
      Left            =   6157
      TabIndex        =   14
      Top             =   6120
      Width           =   1455
   End
   Begin VB.CommandButton CMDEXIT 
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
      Left            =   6742
      TabIndex        =   16
      Top             =   6720
      Width           =   1440
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "UPDATE"
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
      Left            =   1462
      TabIndex        =   15
      Top             =   6720
      Width           =   1440
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
      Height          =   5535
      Left            =   472
      TabIndex        =   29
      Top             =   960
      Width           =   3255
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
         ItemData        =   "updpatient.frx":0000
         Left            =   2040
         List            =   "updpatient.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   1095
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
         Left            =   2040
         TabIndex        =   2
         Top             =   840
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
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1440
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
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid dgsearch 
         Height          =   3255
         Left            =   120
         TabIndex        =   30
         Top             =   2160
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   5741
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
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
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "select serch option"
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
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "search "
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
         TabIndex        =   31
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.TextBox txtage 
      BackColor       =   &H00FFFFFF&
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
      Index           =   0
      Left            =   10800
      TabIndex        =   18
      Top             =   1080
      Width           =   975
   End
   Begin VB.ComboBox cmbsex 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      ItemData        =   "updpatient.frx":001B
      Left            =   10800
      List            =   "updpatient.frx":0028
      TabIndex        =   17
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtph 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   10800
      TabIndex        =   1
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Update Patient Record"
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
      Left            =   1320
      TabIndex        =   33
      Top             =   0
      Width           =   7095
   End
   Begin VB.Label Label8 
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
      Left            =   4312
      TabIndex        =   28
      Top             =   3390
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Address"
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
      Left            =   4312
      TabIndex        =   27
      Top             =   2295
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Patient Name"
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
      Index           =   0
      Left            =   4312
      TabIndex        =   26
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "So/Do/Wo"
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
      Left            =   4312
      TabIndex        =   25
      Top             =   1740
      Width           =   1695
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   4312
      TabIndex        =   24
      Top             =   5025
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
      Height          =   255
      Left            =   4312
      TabIndex        =   23
      Top             =   5565
      Width           =   1695
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Diagnosis"
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
      Left            =   4312
      TabIndex        =   22
      Top             =   2835
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Education"
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
      Left            =   4312
      TabIndex        =   21
      Top             =   3930
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Phone no."
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
      Left            =   4312
      TabIndex        =   20
      Top             =   6120
      Width           =   1695
   End
   Begin VB.Label Label11 
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
      Height          =   255
      Left            =   4312
      TabIndex        =   19
      Top             =   4470
      Width           =   1650
   End
End
Attribute VB_Name = "updpatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim st As String
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
Private Sub cmdclear_Click()
 txtsearch.Text = ""
 
 st = ""
 rs.Filter = 0
 rs.Requery
 Set dgsearch.DataSource = rs
 dgsearch.Enabled = True
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

Private Sub Cmdupdate_Click()
 If MsgBox("CONFIRM RECORD UPDATED", vbYesNo, "USER QUESTION") = 6 Then
        If txtname = "" Or txag = "" Or txtdoa = "" Or txtadd = "" Or txtsex.Text = "" Then
            MsgBox "PLEASE ENTER PATIENT NAME ,AGE,SEX,ADDRESS,DATE OF ADMISSION", vbOKOnly, "USER INFORMATION"
        Exit Sub
        End If
    txt2rs
 Else
   Exit Sub
 End If
End Sub

Private Sub dgsearch_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If rs1.RecordCount > 0 And Not (rs1.BOF Or rs1.EOF) Then
   rs2txt
  End If
 Exit Sub
End Sub
Private Sub Form_Load()
main.Enabled = False

cn.CursorLocation = adUseClient
If cn.State = 1 Then cn.Close
cn.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\agrawal.mdb;Persist Security Info=False"
If rs.State = 1 Then rs.Close
rs.Open "select * from patient", cn, adOpenDynamic, adLockOptimistic
If rs1.State = 1 Then rs1.Close
rs1.Open "select pcode,category from bed_tra", cn, adOpenDynamic, adLockOptimistic
Set dgsearch.DataSource = rs
txtpcode.Visible = False
rs2txt
End Sub
Private Sub rs2txt()
If rs.EOF Then
txtempty
Exit Sub
End If
txtrfa = rs!pccomplaint
txtedu = rs!education
txtdoa = rs!doa
txtatten = rs!paname
txtname = rs!pname
txtadd = rs!address
txag = rs!age
txtdoa = rs!doa
txp = rs!ph
txtsex = rs!sex
txtpcode = rs!pcode
rs1.Filter = "pcode" & " like '" & txtpcode & "*'"
txtward = rs1!category
End Sub
Private Sub txtempty()
txtrfa = ""
txtedu = ""
txtdoa = ""
txtatten = ""
txtname = ""
txag = ""
txtdoa = ""
txp = ""
txtsex = ""
txtward = ""
txtadd = ""
txtpcode = ""
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
              rs2txt
        End If
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If rs.State = 1 Then rs.Close
If rs1.State = 1 Then rs1.Close
If cn.State = 1 Then cn.Close
End Sub

Private Sub txtsex_GotFocus()
  If txag = "" Then
    MsgBox "PLEASE ENTER PATIENT AGE", vbOKOnly, "USER INFORMATION"
    txag.SetFocus
   End If
End Sub
Private Sub txtatten_GotFocus()
 If txtname.Text = "" Then
     MsgBox "PLEASE ENTER PATIENT NAME", vbOKOnly, "USER INFORMATION"
     txtname.SetFocus
  End If
End Sub
Private Sub txag_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
     MsgBox "PLEASE ENTER CORRECT PATIENT AGE", vbOKOnly, "USER INFORMATION"
     KeyAscii = 0
     txag = ""
  End If
End Sub
Private Sub txtname_KeyPress(KeyAscii As Integer)
   If (KeyAscii >= 48 And KeyAscii <= 57) Then
     MsgBox "PLEASE ENTER CORRECT PATIENT'S NAME", vbOKOnly
     KeyAscii = 0
     txtname = ""
     Exit Sub
   End If
End Sub
Private Sub txp_KeyPress(KeyAscii As Integer)
  If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
    MsgBox "PLEASE ENTER CORRECT PATIENT'S PHONE NO.", vbOKOnly
    KeyAscii = 0
    txp = ""
  Exit Sub
 End If
End Sub
Private Sub txtrfa_GotFocus()
   If txtadd = "" Then
     MsgBox "PLEASE ENTER PATIENT ADDRESS", vbOKOnly, "USER INFORMATION"
       txtadd.SetFocus
   End If
End Sub
Private Sub txtatten_KeyPress(KeyAscii As Integer)
  If KeyAscii >= 48 And KeyAscii <= 57 Then
    MsgBox "PLEASE ENTER CORRECT ATTENDER'S NAME", vbOKOnly, "USER INFORMATION"
    KeyAscii = 0
    txtatten = ""
  End If
End Sub
Private Sub txtrfa_KeyPress(KeyAscii As Integer)
   If KeyAscii >= 48 And KeyAscii <= 57 Then
      MsgBox "PLEASE ENTER CORRECT PATIENT'S DISEASE", vbOKOnly, "USER INFORMATION"
      KeyAscii = 0
      txtrfa = ""
    End If
End Sub
Private Sub txtedu_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 48 And KeyAscii <= 57 Then
       MsgBox "PLEASE ENTER CORRECT PATIENT'S EDUCATION", vbOKOnly, "USER INFORMATION"
       KeyAscii = 0
     End If
End Sub
Private Sub txt2rs()
rs!pname = txtname & ""
rs!age = Val(txag)
rs!address = txtadd & ""
rs!ph = txp & ""
rs!education = txtedu & ""
rs!doa = Date
rs!toa = Time
rs!pccomplaint = txtrfa & ""
rs!paname = txtatten & ""
rs.Update
End Sub

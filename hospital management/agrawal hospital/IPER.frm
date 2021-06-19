VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form entry 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Entry form"
   ClientHeight    =   7020
   ClientLeft      =   765
   ClientTop       =   1170
   ClientWidth     =   10425
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00000040&
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form5"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   10425
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport Crpt 
      Left            =   3360
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.TextBox txtrate 
      Height          =   375
      Left            =   3000
      TabIndex        =   24
      Top             =   960
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdsave 
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
      Height          =   375
      Left            =   4544
      TabIndex        =   11
      Top             =   6360
      Width           =   1335
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
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
      Left            =   562
      TabIndex        =   12
      Top             =   6360
      Width           =   1335
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
      Height          =   375
      Left            =   8527
      TabIndex        =   13
      Top             =   6360
      Width           =   1335
   End
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
      Height          =   375
      Left            =   3045
      TabIndex        =   6
      Top             =   5520
      Width           =   1575
   End
   Begin VB.TextBox txtph 
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
      Left            =   8085
      TabIndex        =   10
      Top             =   2640
      Width           =   1935
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
      Height          =   375
      Left            =   3045
      TabIndex        =   5
      Top             =   4830
      Width           =   1575
   End
   Begin VB.ComboBox cmbsex 
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
      ItemData        =   "IPER.frx":0000
      Left            =   8085
      List            =   "IPER.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1800
      Width           =   975
   End
   Begin VB.ComboBox cmbward 
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
      ItemData        =   "IPER.frx":0027
      Left            =   3045
      List            =   "IPER.frx":003A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4065
      Width           =   1575
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
      Left            =   3045
      TabIndex        =   3
      Top             =   3285
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
      Left            =   3045
      TabIndex        =   2
      Top             =   2505
      Width           =   3015
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
      Left            =   8085
      TabIndex        =   8
      Top             =   960
      Width           =   975
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
      Left            =   3045
      TabIndex        =   1
      Top             =   1740
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
      Left            =   3045
      TabIndex        =   0
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Patient Admit form"
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
      Left            =   2745
      TabIndex        =   23
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label11 
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
      Height          =   255
      Left            =   1005
      TabIndex        =   22
      Top             =   5520
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
      Height          =   375
      Left            =   6885
      TabIndex        =   21
      Top             =   2640
      Width           =   975
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
      Left            =   1005
      TabIndex        =   20
      Top             =   4785
      Width           =   1575
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
      Left            =   1005
      TabIndex        =   19
      Top             =   4035
      Width           =   1575
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
      Left            =   1005
      TabIndex        =   18
      Top             =   3300
      Width           =   1575
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
      Left            =   1005
      TabIndex        =   17
      Top             =   2565
      Width           =   1575
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
      Left            =   6885
      TabIndex        =   16
      Top             =   1800
      Width           =   975
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
      Height          =   375
      Left            =   6885
      TabIndex        =   15
      Top             =   960
      Width           =   975
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
      Left            =   1005
      TabIndex        =   14
      Top             =   1815
      Width           =   1575
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
      Left            =   1005
      TabIndex        =   7
      Top             =   1080
      Width           =   1575
   End
End
Attribute VB_Name = "entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim temp As String


Dim d, m As Integer

Private Sub Cmbsex_GotFocus()
If txtage = "" Then
MsgBox "PLEASE ENTER PATIENT AGE", vbOKOnly, "USER INFORMATION"
txtage.SetFocus
End If
End Sub
Private Sub Cmdsave_Click()
If Txtname = "" Or txtage = "" Or Txtdoa = "" Or txtadd = "" Or Cmbsex.Text = "" Or txtph = "" Then
MsgBox "PLEASE ENTER PATIENT NAME ,AGE,SEX,ADDRESS,DATE OF ADMISSION,PHONE NO.", vbOKOnly, "USER INFORMATION"
Exit Sub
End If

txt2rs

If MsgBox("DO YOU WANT REGISTRATION SLIP", vbYesNo) = vbYes Then

Dim str As String
Crpt.Reset
Crpt.DiscardSavedData = True
Crpt.ReportFileName = App.Path & "\rep3.rpt"
str = " {patient.pcode} = '" & temp & "'"
Crpt.ReplaceSelectionFormula str
Crpt.WindowState = crptMaximized
If MsgBox(" CLICK YES FOR PRINT OR NO FOR PREVIEW", vbYesNo, "USER INFORMATION") = vbYes Then

Crpt.Destination = crptToPrinter
On Error GoTo e1
Else
Crpt.Destination = crptToWindow
End If
Crpt.Action = 1
End If

Exit Sub
e1:
MsgBox "    THERE IS NO DEFAULT PRINTER", vbCritical, "PRINTER ERROR"
End Sub
Private Sub Cmdcancel_Click()
txtempty
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
Private Sub Form_Load()
main.Enabled = False

cn.CursorLocation = adUseClient
If cn.State = 1 Then cn.Close
cn.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\agrawal.mdb;Persist Security Info=False"

If rs.State = 1 Then rs.Close
rs.Open "select * from patient", cn, adOpenDynamic, adLockOptimistic
If rs3.State = 1 Then rs3.Close
rs3.Open "select *  from patient", cn, adOpenDynamic, adLockOptimistic

If rs1.State = 1 Then rs1.Close
rs1.Open "select * from bed_tra", cn, adOpenDynamic, adLockOptimistic
If rs2.State = 1 Then rs2.Close
rs2.Open "select * from bed", cn, adOpenDynamic, adLockOptimistic
Txtdoa = Date
Txtdoa.Enabled = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
If rs.State = 1 Then rs.Close
If rs1.State = 1 Then rs1.Close
If rs2.State = 1 Then rs2.Close
If rs3.State = 1 Then rs3.Close
If cn.State = 1 Then cn.Close
End Sub

Private Sub txtatten_GotFocus()
If Txtname.Text = "" Then
MsgBox "PLEASE ENTER PATIENT NAME", vbOKOnly, "USER INFORMATION"
Txtname.SetFocus
End If
End Sub

Private Sub txtage_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
MsgBox "PLEASE ENTER CORRECT PATIENT AGE", vbOKOnly, "USER INFORMATION"
KeyAscii = 0
txtage = ""
End If
End Sub
Private Sub txtname_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
MsgBox "PLEASE ENTER CORRECT PATIENT'S NAME", vbOKOnly
KeyAscii = 0
Txtname = ""
Exit Sub
End If
End Sub
Private Sub txtph_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
MsgBox "PLEASE ENTER CORRECT PATIENT'S PHONE NO.", vbOKOnly
KeyAscii = 0
txtph = ""
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

Private Sub txtempty()
Txtname.Text = ""
txtage.Text = ""
txtph.Text = ""
txtadd.Text = ""
txtrfa.Text = ""
txtatten.Text = ""
txtedu.Text = ""
End Sub
Private Sub txt2rs()
If cmbward.Text = "" Then
 MsgBox "PLEASE SELECT WARD", vbOKOnly
 Exit Sub
End If
  rs2.Filter = "category" & " like '" & cmbward.Text & "*'"
  
If Not (rs2!empty = 0) Then
    rs2!booked = (rs2!booked + 1)
    rs2!empty = ((rs2!empty) - 1)
    rs.AddNew
   
    rs!pname = Txtname & ""
   
If rs3.RecordCount > 0 Then
           If rs3.State = 1 Then rs3.Close
           rs3.Open "select *  from patient", cn, adOpenDynamic, adLockOptimistic
rs3.Requery
rs3.Filter = 0
           rs3.MoveLast
           
           z = rs3!id
        Else
             z = 0
      End If
   d = z + 1
   
     rs!pcode = "A" & "/" & Year(Date) & "/" & Month(Date) & "/" & d
    temp = "A" & "/" & Year(Date) & "/" & Month(Date) & "/" & d

    rs!age = Val(txtage)
    rs!address = txtadd & ""
    rs!ph = txtph & ""
    rs!education = txtedu & ""
    rs!doa = Date
    rs!toa = Time
    rs!pccomplaint = txtrfa & ""
    rs!paname = txtatten & ""
    m = Cmbsex.ListIndex
    Select Case m
     Case 0
         rs!sex = "male" & ""
     Case 1
         rs!sex = "female" & ""
     Case Else
         rs!sex = "--" & ""
     End Select
    rs1.AddNew
    Select Case cmbward.ListIndex
     Case 0
         rs1!category = "ICU"
     Case 1
         rs1!category = "Delux"
     Case 2
         rs1!category = "General"
     Case 3
         rs1!category = "Private"
     Case 4
         rs1!category = "Semiprivate"
     End Select
    rs1!pcode = "A" & "/" & Year(Date) & "/" & Month(Date) & "/" & d
    rs1!Date = Date
    rs2.Filter = "category" & " like '" & cmbward.Text & "*'"
    rs1!Rate = rs2!Rate
    rs1!tos = rs!toa
    rs2.Update
    rs.Update
    rs1.UpdateBatch
    Cmdcancel_Click
    rs3.Close
     rs3.Open "select *  from patient", cn, adOpenDynamic, adLockOptimistic


Else
  MsgBox "THIS WARD IS FULL", vbOKOnly
  cmbward.SetFocus
  Exit Sub
End If
If rs3.RecordCount > 0 Then
           If rs3.State = 1 Then rs3.Close
           rs3.Open "select *  from patient", cn, adOpenDynamic, adLockOptimistic
            rs3.Requery
rs3.Filter = 0
           rs3.MoveLast
           
           z = rs3!id
        Else
             z = 0
      End If






End Sub

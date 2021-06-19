VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form staff 
   BackColor       =   &H00FFC0C0&
   Caption         =   "staff form"
   ClientHeight    =   6495
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8175
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   8175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   37
      Top             =   5760
      Width           =   855
   End
   Begin VB.CommandButton cmdcanle 
      Caption         =   "Cancle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   36
      Top             =   5760
      Width           =   975
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
      Index           =   3
      Left            =   6720
      TabIndex        =   8
      Top             =   5760
      Width           =   975
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
      Index           =   2
      Left            =   3120
      TabIndex        =   5
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton Cmddelete 
      Caption         =   "Delete"
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
      Index           =   1
      Left            =   1920
      TabIndex        =   6
      Top             =   5760
      Width           =   975
   End
   Begin VB.CommandButton Cmdnew 
      Caption         =   "Add  "
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
      Index           =   0
      Left            =   720
      TabIndex        =   7
      Top             =   5760
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   4695
      Left            =   3600
      TabIndex        =   13
      Top             =   840
      Width           =   4215
      Begin VB.TextBox txtecode 
         Height          =   285
         Left            =   1680
         TabIndex        =   35
         Top             =   4200
         Width           =   1575
      End
      Begin VB.TextBox txtdoj 
         Height          =   375
         Left            =   1680
         TabIndex        =   34
         Top             =   3720
         Width           =   1455
      End
      Begin VB.TextBox txtph 
         Height          =   285
         Left            =   1560
         TabIndex        =   33
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox txtadd 
         Height          =   285
         Left            =   1680
         TabIndex        =   32
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox txtedu 
         Height          =   285
         Left            =   1680
         TabIndex        =   31
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtage 
         Height          =   285
         Left            =   1560
         TabIndex        =   30
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtdob 
         Height          =   285
         Left            =   1560
         TabIndex        =   29
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtfname 
         Height          =   285
         Left            =   1560
         TabIndex        =   28
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtname 
         Height          =   285
         Left            =   1560
         TabIndex        =   27
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox Cmbdegic 
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
         ItemData        =   "staff.frx":0000
         Left            =   1800
         List            =   "staff.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2460
         Width           =   1935
      End
      Begin VB.ComboBox Cmbsex 
         Height          =   315
         ItemData        =   "staff.frx":003E
         Left            =   2880
         List            =   "staff.frx":004B
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Employee code"
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
         TabIndex        =   25
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFC0C0&
         Caption         =   "date of joining"
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
         TabIndex        =   24
         Top             =   3690
         Width           =   1575
      End
      Begin VB.Label Label12 
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
         Left            =   2400
         TabIndex        =   22
         Top             =   1560
         Width           =   375
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Phone no"
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
         TabIndex        =   21
         Top             =   3300
         Width           =   1575
      End
      Begin VB.Label Label9 
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
         Left            =   120
         TabIndex        =   20
         Top             =   2790
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Degicnation"
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
         TabIndex        =   19
         Top             =   2415
         Width           =   1575
      End
      Begin VB.Label Label7 
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
         Left            =   120
         TabIndex        =   18
         Top             =   2025
         Width           =   1575
      End
      Begin VB.Label Label6 
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
         Left            =   120
         TabIndex        =   17
         Top             =   1635
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Date of birth"
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
         TabIndex        =   16
         Top             =   1245
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Father's name"
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
         TabIndex        =   15
         Top             =   870
         Width           =   1575
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
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   4695
      Left            =   600
      TabIndex        =   9
      Top             =   840
      Width           =   3015
      Begin VB.TextBox txtsel 
         Height          =   285
         Left            =   1440
         TabIndex        =   26
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Cmdclear 
         Caption         =   "clear"
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
         Left            =   1680
         TabIndex        =   2
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Cmdserch 
         Caption         =   "serch"
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
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   975
      End
      Begin VB.ComboBox Cmbserch 
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
         ItemData        =   "staff.frx":0060
         Left            =   1680
         List            =   "staff.frx":006A
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid Dgsearch 
         Height          =   2775
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   4895
         _Version        =   393216
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
         Caption         =   "Select "
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
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Staff information"
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
      Left            =   2160
      TabIndex        =   23
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "staff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim j As Integer
Dim i As Integer
Dim st As String

Private Sub Cmbserch_Click()
Select Case Cmbserch.ListIndex
Case 0
st = "ecode"
Case 1
st = "ename"
Case Else
st = ""
End Select
End Sub

Private Sub Cmddelete_Click(Index As Integer)
If rs.RecordCount = 0 Then
MsgBox "NO RECORDS", vbOKOnly
Exit Sub
End If
If MsgBox("CONFIRM RECORD UPDATED", vbYesNo, "USER INFORMATION") = 6 Then
rs.Delete
rs.Update
txtempty
Else
Exit Sub
End If
End Sub

Private Sub CMDEXIT_Click(Index As Integer)
main.Enabled = True

If rs.State = 1 Then rs.Close
If cn.State = 1 Then cn.Close
Unload Me
Load main
main.Show

End Sub

Private Sub cmdnew_Click(Index As Integer)
cmdsave.Enabled = True

txtempty

End Sub

Private Sub Cmdsave_Click()
If txtname = "" Or txtfname = "" Or txtage = "" Or txtadd = "" Or Cmbdegic.Text = "" Or txtph = "" Then
     MsgBox "PLEASE ENTER  NAME ,AGE,SEX,ADDRESS,DATE OF ADMISSION", vbOKOnly, "USER INFORMATION"
Exit Sub
End If
i = rs.RecordCount
rs.AddNew
rs!ecode = "A" & "/" & "N " & "/" & Year(Date) & "/" & Month(Date) & "/" & i
rs!ename = "" & txtname
rs!gname = "" & txtfname
rs!dob = DateValue(txtdob)
rs!doj = Date
rs!address = "" & txtadd
rs!sex = "" & Cmbsex.Text
rs!phone = "" & txtph
rs!edu = "" & txtedu
rs!degn = "" & Cmbdegic.Text
rs!age = Val(txtage)
rs!ecode = "A" & "/" & " N" & "/" & Year(Date) & "/" & Month(Date) & "/" & i
rs.Update
txtempty
End Sub

Private Sub Cmdserch_Click()
If Cmbserch.Text = "" Or Txtsl = "" Then
MsgBox "PLEASE SELECT  FIELD TO SEARCH", vbInformation, "USER INFORMATION"
Exit Sub
End If
If Txtsl = "" And Cmbserch.Text = "ecode" Then
MsgBox "Please Enter code to Search", vbInformation, "User Information"
        Txtsl.SetFocus
  ElseIf Txtsl = "" And Cmbserch.ListIndex = 1 Then
  MsgBox "Please Enter code to Search", vbInformation, "User Information"
        Txtsl.SetFocus
Else
    
        Set Dgsearch.DataSource = Nothing
        txtempty
        rs.Filter = st & " like '" & Txtsl & "*'"
        Set Dgsearch.DataSource = rs
          
        If rs.RecordCount = 0 Then
            MsgBox "No Records Found", vbInformation, "User Information"
            txtempty
        Else
              rs2txt
        End If
    End If
End Sub

Private Sub Cmdupdate_Click(Index As Integer)
If rs.RecordCount = 0 Then
MsgBox "NO RECORDS", vbOKOnly
Exit Sub
End If
If MsgBox("CONFIRM RECORD UPDATED", vbYesNo, "USER INFORMATION") = 6 Then
rs!ecode = "A" & "/" & "N " & "/" & Year(Date) & "/" & Month(Date) & "/" & i
rs!ename = "" & txtname
rs!gname = "" & txtfname
rs!dob = txtdob & ""
rs!doj = Date
rs!address = "" & txtadd
rs!sex = "" & Cmbsex.Text
rs!phone = "" & txtph
rs!edu = "" & txtedu
rs!degn = "" & Cmbdegic.Text
rs!age = Val(txtage)
rs!ecode = "A" & "/" & " N" & "/" & Year(Date) & "/" & Month(Date) & "/" & i
rs.Update
Else
Exit Sub
End If

End Sub

Private Sub dgsearch_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
rs2txt
End Sub

Private Sub Form_Load()
main.Enabled = False

cn.CursorLocation = adUseClient
cn.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\agrawal.mdb;Persist Security Info=False"
If rs.State = 1 Then rs.Close
rs.Open "select * from staff", cn, adOpenDynamic, adLockOptimistic
Set Dgsearch.DataSource = rs
rs2txt
cmdsave.Enabled = False

End Sub

Private Sub txtempty()
txtname = ""
txtfname = ""
txtdob = ""
txtage = ""
txtedu = ""
txtadd = ""
txtph = ""
txtdoj = ""
txtecode = ""
End Sub

Private Sub rs2txt()
If rs.RecordCount > 0 And Not (rs.EOF Or rs.BOF) Then
txtname = rs!ename
txtfname = rs!gname
txtdob = rs!dob
txtdoj = rs!doj
txtadd = rs!address
Cmbsex.Text = rs!sex
txtph = rs!phone
txtedu = rs!edu
Cmbdegic.Text = rs!degn
txtage = rs!age
txtecode = rs!ecode
Else
txtempty
Exit Sub
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If rs.State = 1 Then rs.Close
If cn.State = 1 Then cn.Close
End Sub

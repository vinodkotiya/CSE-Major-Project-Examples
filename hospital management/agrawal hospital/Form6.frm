VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form updatservice 
   BackColor       =   &H00FFC0C0&
   Caption         =   "update service"
   ClientHeight    =   4920
   ClientLeft      =   2025
   ClientTop       =   2820
   ClientWidth     =   8070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   LockControls    =   -1  'True
   ScaleHeight     =   4920
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Height          =   3855
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   3255
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
         Left            =   1920
         TabIndex        =   2
         Top             =   960
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
         TabIndex        =   1
         Top             =   960
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid dgsearch 
         Height          =   2055
         Left            =   0
         TabIndex        =   12
         Top             =   1680
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   3625
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ItemData        =   "Form6.frx":0000
         Left            =   1680
         List            =   "Form6.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Select Category"
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
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
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
      Left            =   6720
      TabIndex        =   8
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "DELETE"
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
      Left            =   5160
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
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
      Left            =   3600
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
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
      Height          =   405
      Left            =   5760
      TabIndex        =   4
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox txtitem 
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
      Left            =   5760
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Update Service form"
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
      Left            =   1800
      TabIndex        =   13
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Rate/Rate"
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
      Left            =   4200
      TabIndex        =   9
      Top             =   2520
      Width           =   1335
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
      Left            =   4200
      TabIndex        =   6
      Top             =   1440
      Width           =   1455
   End
End
Attribute VB_Name = "updatservice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim cn As New ADODB.Connection
Dim i As Integer
Dim rs1 As New ADODB.Recordset
Dim a As Integer

Private Sub Cmdupdate_Click()
If MsgBox("CONFIRM RECORD UPDATED", vbYesNo, "USER QUESTION") = 6 Then
    If txtitem = "" Then
        MsgBox "PLEASE ENTER ITEM NAME", vbInformation
    Exit Sub
    End If
rs!itemname = "" & txtitem
rs!Rate = Val(txtrate)
rs.Update
Else
Exit Sub
End If
End Sub

Private Sub cmddel_Click()
If MsgBox("CONFIRM RECORD DELETED", vbYesNo, "USER QUESTION") = 6 Then
  rs.Delete
rs.Update
Else
Exit Sub
End If
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

Private Sub cmdsearch_Click()
If cmbcat.Text = "" Then
MsgBox "PLEASE SELECT  FIELD TO SEARCH", vbInformation, "USER INFORMATION"
Exit Sub

Else
    
        Set dgsearch.DataSource = Nothing
        txtitem = ""
        txtrate = ""
        
        rs.Filter = "category" & " like '" & cmbcat.Text & "*'"
        Set dgsearch.DataSource = rs
          
        If rs.RecordCount = 0 Then
            MsgBox "No Records Found", vbInformation, "User Information"
            cmdclear_Click
        Else
        txtitem = rs!itemname
        txtrate = rs!Rate
        End If
    End If
    

End Sub

Private Sub cmdclear_Click()



    rs.Filter = 0
    rs.Requery
    Set dgsearch.DataSource = rs
  
End Sub

Private Sub dgsearch_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If rs.RecordCount > 0 And Not (rs.BOF Or rs.EOF) Then
txtitem = rs!itemname
txtrate = rs!Rate
End If
End Sub

Private Sub Form_Load()
main.Enabled = False

cn.CursorLocation = adUseClient
cn.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\agrawal.mdb;Persist Security Info=False"
If rs.State = 1 Then rs.Close
rs.Open "select * from service_master", cn, adOpenDynamic, adLockOptimistic
Set dgsearch.DataSource = rs
If rs.EOF = True Then
txtitem = ""
txtrate = ""
Exit Sub
End If
txtitem = rs!itemname
txtrate = rs!Rate
If rs1.State = 1 Then rs1.Close
rs1.Open "select cat from catgri", cn, adOpenDynamic, adLockOptimistic
rs2cmb
End Sub

Private Sub Form_Unload(Cancel As Integer)
If rs.State = 1 Then rs.Close
If rs1.State = 1 Then rs1.Close
If cn.State = 1 Then cn.Close
End Sub

Private Sub txtitem_KeyPress(KeyAscii As Integer)
If KeyAscii >= 48 And KeyAscii <= 57 Then
MsgBox "PLEASE ENTER CORRECT NAME", vbOKOnly, "USER INFORMATION"

KeyAscii = 0
txtitem = ""
End If
End Sub

Private Sub txtrate_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
MsgBox "PLEASE ENTER RATE IN NUMBERS", vbOKOnly, "USER INFORMATION"
KeyAscii = 0
txtrate = ""
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

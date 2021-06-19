VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form bedmodify 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Bed modify form"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ForeColor       =   &H00000080&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
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
      Left            =   5880
      TabIndex        =   14
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton Cmdcancle 
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
      TabIndex        =   13
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton Cmdsave 
      Caption         =   "Save"
      Enabled         =   0   'False
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
      Left            =   1560
      TabIndex        =   12
      Top             =   5400
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
      Left            =   3000
      TabIndex        =   11
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton Cmdnew 
      Caption         =   "New"
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
      Left            =   120
      TabIndex        =   10
      Top             =   5400
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   3495
      Left            =   3360
      TabIndex        =   3
      Top             =   1440
      Width           =   3495
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
         Height          =   375
         Left            =   1800
         TabIndex        =   9
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtbed 
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
         TabIndex        =   8
         Top             =   1740
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
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rate/bed"
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
         Left            =   360
         TabIndex        =   7
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "No of Beds"
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
         Left            =   360
         TabIndex        =   6
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ward name"
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
         Left            =   360
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ward selection"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   3015
      Begin MSDataGridLib.DataGrid dgward 
         Height          =   2775
         Left            =   120
         TabIndex        =   2
         Top             =   480
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
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Bed modify form"
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
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "bedmodify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Dim a As Integer
Dim b As Integer



Private Sub Cmdcancle_Click()
txtempty
End Sub

Private Sub CMDEXIT_Click()
main.Enabled = True
If rs.State = 1 Then rs.Close
If cn.State = 1 Then cn.Close
Unload Me
Load main
main.Show

End Sub

Private Sub cmdnew_Click()
  txtempty
Cmdsave.Enabled = True
Cmdupdate.Enabled = False
End Sub

Private Sub Cmdsave_Click()
rs.AddNew
rs!category = "" & txtward
rs!total = Val(txtbed)
rs!Rate = Val(txtrate)
rs!booked = 0
rs!empty = Val(txtbed)
rs.Update
End Sub

Private Sub Cmdupdate_Click()
Dim c As Integer


If MsgBox("CONFIRM RECORD UPDATED", vbYesNo, "USER QUESTION") = 6 Then
    If txtward = "" Then
        MsgBox "PLEASE ENTER WARD NAME", vbInformation
    Exit Sub
    End If
 If txtrate = "" Then
        MsgBox "PLEASE ENTER RATE", vbInformation
    Exit Sub
    End If
 If txtbed = "" Then
        MsgBox "PLEASE ENTER NO. OF BED", vbInformation
    Exit Sub
    End If
a = rs!total
c = rs!booked
If txtbed < c Then
MsgBox "INVALID NUMBER OF BEDS", vbOKOnly
Exit Sub
End If
rs!category = "" & txtward
rs!Rate = Val(txtrate)
rs!total = Val(txtbed)
If Not (rs!total = a) Then
b = rs!total - a
rs!empty = rs!empty + b
End If

rs.UpdateBatch

Else
Exit Sub
End If
End Sub

Private Sub dgward_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If rs.RecordCount > 0 And Not (rs.BOF Or rs.EOF) Then
   rs2txt
   End If
   
End Sub

Private Sub Form_Load()
main.Enabled = False

cn.CursorLocation = adUseClient
If cn.State = 1 Then cn.Close
cn.Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\agrawal.mdb;Persist Security Info=False"
If rs.State = 1 Then rs.Close
rs.Open "select * from bed", cn, adOpenDynamic, adLockOptimistic

 Set dgward.DataSource = rs
 txtward = rs!category
 txtbed = rs!total
 txtrate = rs!Rate
 End Sub
Private Sub rs2txt()
If rs.EOF Then
txtempty
Exit Sub
End If
txtward = rs!category
txtbed = rs!total
txtrate = rs!Rate

End Sub
Private Sub txtempty()
txtward = ""
txtbed = ""
txtrate = ""

End Sub

Private Sub Form_Unload(Cancel As Integer)
If rs.State = 1 Then rs.Close
If cn.State = 1 Then cn.Close
End Sub

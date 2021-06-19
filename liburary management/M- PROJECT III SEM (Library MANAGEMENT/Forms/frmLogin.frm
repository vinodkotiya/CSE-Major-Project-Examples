VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   Icon            =   "frmLogin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Login"
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   720
      Top             =   1320
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   360
      Left            =   2100
      TabIndex        =   3
      Tag             =   "Cancel"
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   495
      TabIndex        =   2
      Tag             =   "OK"
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txt 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1305
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   525
      Width           =   2325
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   1305
      TabIndex        =   0
      Top             =   135
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   248
      Index           =   1
      Left            =   105
      TabIndex        =   4
      Tag             =   "&Password:"
      Top             =   540
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   248
      Index           =   0
      Left            =   105
      TabIndex        =   5
      Tag             =   "&User Name:"
      Top             =   150
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancel_Click()
    Screen.MousePointer = 0
    End
End Sub

Private Sub cmdOK_Click()
   On Error Resume Next
    Dim i As Integer
    Dim BL As Boolean
    BL = True
    For i = 0 To txt.Count - 1
        If txt(i).text = "" Then
            BL = False
            MsgBox "Please enter a value", vbCritical, "Blank Field"
            txt(i).SetFocus
           Exit For
        Else
           BL = True
        End If
    Next
    
    If BL Then
        '================================================
        'checking whether username and password are valid
        '================================================
            Dim query As String
            Dim objrs As New ADODB.Recordset
       
            FileDSN = "filedsn=" & App.Path & "\connection\library.dsn;uid=;pwd=;"
            
            With ObjCon
                
               .Open FileDSN
               query = "Select UserName from Login where username = '" & txt(0).text & "' and password  = '" & txt(1).text & "'"
               Set objrs = .Execute(query)
               
              If Not objrs.EOF Then
                 LoginPass = txt(0)
                 
                   
                   'MsgBox LoginPass
                    Call wait(1)
                    Load frmMain
                    frmMain.Show
                    Call wait(0)

                    Unload Me
                Else
                   MsgBox "Invalid Login name or Password", vbExclamation, "Invalid Login"
                    txt(0).text = ""
                    txt(1).text = ""
                    txt(0).SetFocus
                End If
                Set objrs = Nothing
                .Close
            End With
           Set ObjCon = Nothing
    End If
End Sub

Private Sub txt_Change(Index As Integer)

End Sub

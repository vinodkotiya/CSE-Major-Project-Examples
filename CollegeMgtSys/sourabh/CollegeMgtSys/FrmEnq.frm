VERSION 5.00
Begin VB.Form FrmEnq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enquiry Module"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "FrmEnq.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6120
   Begin VB.Frame Frame1 
      Caption         =   "Enquiry Module"
      Height          =   3420
      Left            =   30
      TabIndex        =   8
      Top             =   0
      Width           =   6045
      Begin VB.ComboBox CmbStatus 
         Height          =   315
         Index           =   4
         ItemData        =   "FrmEnq.frx":0442
         Left            =   1515
         List            =   "FrmEnq.frx":044C
         TabIndex        =   4
         Text            =   "Paid"
         Top             =   2250
         Width           =   2625
      End
      Begin VB.TextBox TxtDetails 
         Height          =   285
         Index           =   3
         Left            =   1515
         TabIndex        =   3
         Top             =   1890
         Width           =   2610
      End
      Begin VB.TextBox TxtDetails 
         Height          =   285
         Index           =   2
         Left            =   1515
         TabIndex        =   2
         Top             =   1515
         Width           =   2610
      End
      Begin VB.TextBox TxtDetails 
         Height          =   285
         Index           =   1
         Left            =   1515
         TabIndex        =   1
         Top             =   1155
         Width           =   2610
      End
      Begin VB.TextBox TxtDetails 
         Height          =   285
         Index           =   0
         Left            =   1515
         TabIndex        =   0
         Top             =   780
         Width           =   2610
      End
      Begin VB.Frame Frame2 
         Caption         =   "Operations"
         Height          =   735
         Left            =   120
         TabIndex        =   15
         Top             =   2610
         Width           =   5790
         Begin VB.CommandButton CmdOperation 
            Caption         =   "Ca&ncel"
            Height          =   270
            Index           =   2
            Left            =   4065
            TabIndex        =   7
            ToolTipText     =   "Close Window"
            Top             =   330
            Width           =   885
         End
         Begin VB.CommandButton CmdOperation 
            Caption         =   "&Modify"
            Height          =   270
            Index           =   1
            Left            =   2295
            TabIndex        =   6
            ToolTipText     =   "Modify Transaction"
            Top             =   315
            Width           =   885
         End
         Begin VB.CommandButton CmdOperation 
            Caption         =   "&Ok"
            Height          =   270
            Index           =   0
            Left            =   690
            TabIndex        =   5
            ToolTipText     =   "Commit Transaction"
            Top             =   285
            Width           =   885
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "payment Status"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   14
         Top             =   2355
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Issue Date"
         Height          =   195
         Index           =   3
         Left            =   510
         TabIndex        =   13
         Top             =   1971
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Form Cost "
         Height          =   195
         Index           =   2
         Left            =   525
         TabIndex        =   12
         Top             =   1589
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Form No"
         Height          =   195
         Index           =   1
         Left            =   675
         TabIndex        =   11
         Top             =   1207
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Enquiry No"
         Height          =   195
         Index           =   0
         Left            =   495
         TabIndex        =   10
         Top             =   825
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Enquiry Module"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   360
         Left            =   1890
         TabIndex        =   9
         Top             =   240
         Width           =   2010
      End
   End
End
Attribute VB_Name = "FrmEnq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub DataStore()
    Dim i As Integer 'local variable
    For i = 0 To 3 Step 1s
        SeqGen.RecSet1.Fields(i) = TxtDetails(i).Text
    Next i
    SeqGen.RecSet1.Fields(4) = CmbStatus(4).Text
End Sub
Private Sub CmdOperation_Click(Index As Integer)
    'select appropiate button
    
    Select Case Index
        Case 0: 'ok button is presed
                Dim i As Integer 'local variable
                
                'retriving data from enquiry table
                
                DataStore
                
                'asssigning the value of txtfields to recordset
                
'                For i = 0 To 3
'                    RecSet1.Fields(i) = TxtDetails(i).Text
'                Next i

'                RecSet1!enq_no = TxtDetails(0).Text
                'assigning combobox value to recordset
'                SeqGen.RecSet1.Fields(0) = TxtDetails(0).Text
'                SeqGen.RecSet1.Fields(1) = TxtDetails(1).Text
'                SeqGen.RecSet1.Fields(2) = TxtDetails(2).Text
'                SeqGen.RecSet1.Fields(3) = TxtDetails(3).Text
'                SeqGen.RecSet1.Fields(4) = CmbStatus(4).Text
                SeqGen.RecSet1.Update 'updateing the recordset
                SeqGen.RecSet1.Close
                For i = 0 To 3 Step 1
                    TxtDetails(i).Text = ""
                Next i
                CmbStatus(4).Text = "Paid"
        Case 1: 'Modify is pressed
        
        Case 2: 'cancel is pressed
            Unload FrmEnq
    End Select
End Sub

Private Sub Form_Load()

    'Do display Windows in center
    FrmEnq.Top = (CFrmMain.ScaleHeight - FrmEnq.ScaleHeight) / 2
    FrmEnq.Left = (CFrmMain.ScaleWidth - FrmEnq.ScaleWidth) / 2
    
    'to keep windows unresizeable
    FrmEnq.Height = 3840
    FrmEnq.Width = 6210
    
    SeqGen.dataConnect 'connect to database
    SeqGen.RecSet1.Open "select * from enquiry"
    SeqGen.RecSet1.AddNew 'adding new record
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SeqGen.RecSet1.Close 'closeing recordset
End Sub

Private Sub TxtDetails_GotFocus(Index As Integer)
    'display the sequence no at load time
    Select Case Index
        Case 0:
                'local variables
                
                Dim StrLength As Integer 'to store strlength
                Dim Convert As String 'to store converted string
                Dim Retrive As String 'to store retrived value
                
                SeqGen.dataConnect
                SeqGen.RecSet1.Open "Select max(to_number(substr(enq_no,1,5)))+1 as NewNum from enquiry"
                
                'checking for null value
                If IsNull(SeqGen.RecSet1!NewNum) Then
                    TxtDetails(0).Text = "E00001"
                Else
                    Convert = CStr(SeqGen.RecSet1!NewNum) 'convert from integer to str
                    StrLength = Len(Convert) ' finding length of str
                    Retrive = SeqGen.SeqGen(StrLength) 'passing parameter to function
                    
                    TxtDetails(0).Text = "E" & Retrive & Convert 'Putting the retrived value to text box
                    
                End If
                
                SeqGen.RecSet1.Close 'close the recordset
        Case 3: 'display current date to text box
                TxtDetails(3).Text = Format(Date, "dd/mmm/yy")
                TxtDetails(3).Enabled = False
    End Select
                    
End Sub

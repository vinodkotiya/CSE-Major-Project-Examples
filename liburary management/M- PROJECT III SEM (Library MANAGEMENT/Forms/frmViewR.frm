VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmViewR 
   Caption         =   "View :- Return Details"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin MSFlexGridLib.MSFlexGrid MFG 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1296
      _Version        =   393216
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmViewR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CROWS As Integer
Private Sub cmdExit_Click()
    Screen.MousePointer = 0
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
    Call Form_Load
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    With ObjCon
        .Open FileDSN
            
            '=============================================
            'Retreiving all the records from Desired Table
            '=============================================
            query = "select * from ReturnDetails order by LCardNo"
            Set objrs = .Execute(query)
            
            
            MFG.Cols = 5
            
            MFG.Row = 0
            MFG.Col = 0
            MFG.text = "No."
            
            MFG.Col = 1
            MFG.ColWidth(MFG.Col) = 2500
            MFG.text = "AccessionNo"
            
            MFG.Col = 2
            MFG.ColWidth(MFG.Col) = 2500
            MFG.text = "LCardNo"
            
            MFG.Col = 3
            MFG.ColWidth(MFG.Col) = 2500
            MFG.text = "ReturnDate"
            
            MFG.Col = 4
            MFG.ColWidth(MFG.Col) = 1000
            MFG.text = "Fine"
            
           
           
            If Not objrs.EOF Then
            
                '=============================================
                'Retreiving all the records from Desired Table
                '=============================================
                Dim objrs1 As Recordset
                
                
                query = "select count(AccessionNo) from ReturnDetails"
                Set objrs1 = .Execute(query)
                
                MFG.Rows = objrs1(0) + 1
                
                MFG.Col = 0
                MFG.Row = 1
                For i = 1 To CInt(objrs1(0))
                    MFG.text = i
                    MFG.Row = i
                Next
                
                CROWS = 1
                Set objrs1 = Nothing
                
                MFG.Row = CROWS
                
                While Not objrs.EOF
                
                    For i = 1 To 4
                        
                        MFG.Col = i
                        MFG.text = objrs(i - 1)
                        
                    Next
                    
                    MFG.Row = CROWS + 1
                    objrs.MoveNext
                    
                Wend
            Else
                Beep
                MsgBox "No Records Exist", vbExclamation, "Warning"
                .Close
                Screen.MousePointer = 0
                Exit Sub
            End If
        .Close
    End With
    Screen.MousePointer = 0
End Sub

Private Sub MSFlexGrid1_Click()
    
End Sub


Private Sub Form_Resize()
    MFG.Height = Me.Height - 500
    MFG.Width = Me.Width
    
End Sub

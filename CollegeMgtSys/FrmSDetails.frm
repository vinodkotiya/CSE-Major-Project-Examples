VERSION 5.00
Begin VB.Form FrmSDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Entry Module"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   Icon            =   "FrmSDetails.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   8700
   Begin VB.Frame Frame1 
      Caption         =   "Student Details"
      Height          =   6780
      Left            =   0
      TabIndex        =   25
      Top             =   -15
      Width           =   8685
      Begin VB.ComboBox CmbSex 
         Height          =   315
         Index           =   5
         ItemData        =   "FrmSDetails.frx":0442
         Left            =   1710
         List            =   "FrmSDetails.frx":044C
         TabIndex        =   5
         Text            =   "Male"
         Top             =   2580
         Width           =   1950
      End
      Begin VB.ComboBox Cmbic 
         Height          =   315
         Index           =   12
         ItemData        =   "FrmSDetails.frx":045E
         Left            =   1710
         List            =   "FrmSDetails.frx":0468
         TabIndex        =   12
         Text            =   "Issued"
         Top             =   6330
         Width           =   1950
      End
      Begin VB.ComboBox CmbLC 
         Height          =   315
         Index           =   11
         ItemData        =   "FrmSDetails.frx":0480
         Left            =   1710
         List            =   "FrmSDetails.frx":048A
         TabIndex        =   11
         Text            =   "Issued"
         Top             =   5955
         Width           =   1950
      End
      Begin VB.ComboBox CmbFaculty 
         Height          =   315
         Index           =   1
         ItemData        =   "FrmSDetails.frx":04A2
         Left            =   4215
         List            =   "FrmSDetails.frx":04B8
         TabIndex        =   1
         Text            =   "BCA I"
         Top             =   1230
         Width           =   2220
      End
      Begin VB.Frame Frame3 
         Caption         =   "&Operation"
         Height          =   3540
         Left            =   6540
         TabIndex        =   40
         Top             =   1485
         Width           =   2055
         Begin VB.CommandButton CmdOperation 
            Caption         =   "&Close"
            Height          =   285
            Index           =   7
            Left            =   1050
            TabIndex        =   20
            ToolTipText     =   "Close Window"
            Top             =   3045
            Width           =   915
         End
         Begin VB.CommandButton CmdOperation 
            Caption         =   "&Ok"
            Height          =   285
            Index           =   6
            Left            =   75
            TabIndex        =   19
            ToolTipText     =   "Save Record"
            Top             =   3045
            Width           =   915
         End
         Begin VB.CommandButton CmdOperation 
            Caption         =   "Ca&ncel"
            Height          =   285
            Index           =   5
            Left            =   1050
            TabIndex        =   18
            ToolTipText     =   "Cancel Operation"
            Top             =   2190
            Width           =   915
         End
         Begin VB.CommandButton CmdOperation 
            Caption         =   "&Display"
            Height          =   285
            Index           =   4
            Left            =   75
            TabIndex        =   17
            ToolTipText     =   "Display Record"
            Top             =   2190
            Width           =   915
         End
         Begin VB.CommandButton CmdOperation 
            Caption         =   "&Find"
            Height          =   285
            Index           =   3
            Left            =   1050
            TabIndex        =   16
            ToolTipText     =   "find Record"
            Top             =   1215
            Width           =   915
         End
         Begin VB.CommandButton CmdOperation 
            Caption         =   "&Modify"
            Height          =   285
            Index           =   2
            Left            =   75
            TabIndex        =   15
            ToolTipText     =   "modify Record"
            Top             =   1215
            Width           =   915
         End
         Begin VB.CommandButton CmdOperation 
            Caption         =   "D&elete"
            Height          =   285
            Index           =   1
            Left            =   1050
            TabIndex        =   14
            ToolTipText     =   "Delete Record"
            Top             =   330
            Width           =   915
         End
         Begin VB.CommandButton CmdOperation 
            Caption         =   "&Add"
            Height          =   285
            Index           =   0
            Left            =   75
            TabIndex        =   13
            ToolTipText     =   "Add Record"
            Top             =   330
            Width           =   915
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "&Movement"
         Height          =   1530
         Left            =   6540
         TabIndex        =   39
         Top             =   5070
         Width           =   2040
         Begin VB.CommandButton CmdMovement 
            Height          =   360
            Index           =   3
            Left            =   1200
            Picture         =   "FrmSDetails.frx":04F2
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Go Next"
            Top             =   645
            Width           =   390
         End
         Begin VB.CommandButton CmdMovement 
            Height          =   360
            Index           =   2
            Left            =   825
            Picture         =   "FrmSDetails.frx":0934
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Go Last"
            Top             =   990
            Width           =   390
         End
         Begin VB.CommandButton CmdMovement 
            Height          =   360
            Index           =   1
            Left            =   420
            Picture         =   "FrmSDetails.frx":0D76
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Go Previous"
            Top             =   645
            Width           =   390
         End
         Begin VB.CommandButton CmdMovement 
            Height          =   375
            Index           =   0
            Left            =   795
            Picture         =   "FrmSDetails.frx":11B8
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Go First"
            Top             =   285
            Width           =   405
         End
      End
      Begin VB.TextBox TxtSDetails 
         Height          =   285
         Index           =   9
         Left            =   1710
         TabIndex        =   9
         Top             =   4815
         Width           =   1950
      End
      Begin VB.TextBox TxtSDetails 
         Height          =   285
         Index           =   6
         Left            =   1710
         TabIndex        =   6
         Top             =   2925
         Width           =   1950
      End
      Begin VB.TextBox TxtSDetails 
         Height          =   735
         Index           =   10
         Left            =   1710
         TabIndex        =   10
         Top             =   5160
         Width           =   4710
      End
      Begin VB.TextBox TxtSDetails 
         Height          =   735
         Index           =   8
         Left            =   1710
         TabIndex        =   8
         Top             =   4035
         Width           =   4710
      End
      Begin VB.TextBox TxtSDetails 
         Height          =   735
         Index           =   7
         Left            =   1710
         TabIndex        =   7
         Top             =   3240
         Width           =   4710
      End
      Begin VB.TextBox TxtSDetails 
         Height          =   285
         Index           =   4
         Left            =   1710
         TabIndex        =   4
         Top             =   2250
         Width           =   4710
      End
      Begin VB.TextBox TxtSDetails 
         Height          =   285
         Index           =   3
         Left            =   1710
         TabIndex        =   3
         Top             =   1905
         Width           =   4725
      End
      Begin VB.TextBox TxtSDetails 
         Height          =   285
         Index           =   0
         Left            =   1710
         TabIndex        =   0
         Top             =   1245
         Width           =   1860
      End
      Begin VB.TextBox TxtSDetails 
         Height          =   285
         Index           =   2
         Left            =   1695
         TabIndex        =   2
         Top             =   1590
         Width           =   4740
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Faculty"
         Height          =   195
         Left            =   3675
         TabIndex        =   41
         Top             =   1335
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Li&b Card Status"
         Height          =   195
         Index           =   12
         Left            =   540
         TabIndex        =   38
         Top             =   6045
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "I-&Card Status"
         Height          =   195
         Index           =   11
         Left            =   690
         TabIndex        =   37
         Top             =   6405
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "&Education"
         Height          =   195
         Index           =   10
         Left            =   900
         TabIndex        =   36
         Top             =   5715
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Date Of &Admission"
         Height          =   195
         Index           =   9
         Left            =   300
         TabIndex        =   35
         Top             =   4905
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "&Permanent Add"
         Height          =   195
         Index           =   8
         Left            =   510
         TabIndex        =   34
         Top             =   4560
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "&Local Add"
         Height          =   195
         Index           =   7
         Left            =   900
         TabIndex        =   33
         Top             =   3765
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "D&ate Of Birth"
         Height          =   195
         Index           =   6
         Left            =   690
         TabIndex        =   32
         Top             =   3015
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "&Sex"
         Height          =   195
         Index           =   5
         Left            =   1350
         TabIndex        =   31
         Top             =   2685
         Width           =   285
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "&Father's Name"
         Height          =   195
         Index           =   4
         Left            =   600
         TabIndex        =   30
         Top             =   2340
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "E-Mail ID"
         Height          =   195
         Index           =   3
         Left            =   960
         TabIndex        =   29
         Top             =   2010
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Student &Name"
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   28
         Top             =   1695
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Student &Code"
         Height          =   195
         Index           =   1
         Left            =   630
         TabIndex        =   27
         Top             =   1335
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Student Details Module"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   360
         Index           =   0
         Left            =   2565
         TabIndex        =   26
         Top             =   405
         Width           =   3240
      End
   End
End
Attribute VB_Name = "FrmSDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Flag As String 'variable to check which button is clicked

'procedure for hide and display buttons

Private Sub HideShowButtons(Index As Integer)
    Dim i As Integer
    Select Case Index
        Case 0, 1, 2, 3, 4:  'Hide the buttons
        
            For i = 0 To 4
                CmdOperation(i).Enabled = False
            Next i
        Exit Sub
            
        Case 5, 6, 7 ' Show the buttons
            For i = 0 To 4
                CmdOperation(i).Enabled = True
            Next i
        Exit Sub
    End Select
End Sub

Private Sub CmdOperation_Click(Index As Integer)
    Select Case Index
        Case 0: 'Add
                Dim i As Integer 'variable for looping
                
                HideShowButtons (Index)
                For i = 0 To 12 Step 1
                        If Not i = 1 And Not i = 5 And Not i = 11 And Not i = 12 Then
                            TxtSDetails(i).Text = ""
                        End If
                Next i
                
                Flag = "Add"
                'Open record set
                SeqGen.dataConnect
                SeqGen.RecSet1.Open "select * from student_details"
                SeqGen.RecSet1.AddNew
                    
                MsgBox "Add is Clicked"
        Case 1: 'Delete
                Flag = "Delete"
                HideShowButtons (Index)
        Case 2: 'Modify
                Flag = "Modify"
                HideShowButtons (Index)
        Case 3: 'Search
                Flag = "Find"
                HideShowButtons (Index)
        Case 4: 'Display
                Flag = "Display"
                HideShowButtons (Index)
        Case 5: 'Cancel
                HideShowButtons (Index)
        Case 6: 'Commit
                HideShowButtons (Index)
                Select Case Flag
                    Case "Add": 'add button is pressed
                        Dim j As Integer 'variable for looping
                        For j = 0 To 12 Step 1
                            If Not j = 1 And Not j = 5 And Not j = 11 And Not j = 15 Then
                            'assigning record to record set
'                                SeqGen.RecSet1.Fields(j) = TxtSDetails(j).Text
                            End If
                        Next j
                        'SeqGen.RecSet1!
                        'adding records to database
                        SeqGen.RecSet1.Update
                        Flag = ""
                        SeqGen.RecSet1.Close
                    Case "Delete": 'Delete Key is pressed
                        Flag = ""
                    Case "Display": 'Display Key is pressed
                        Flag = ""
                    Case "Modify": 'Modify key is pressed
                        Flag = ""
                    Case "Find":   'Find key is pressed
                        Flag = ""
                End Select
                    
        Case 7: 'Close The Windows
                HideShowButtons (Index)
                Unload FrmSDetails
    End Select
End Sub

Private Sub Form_Load()
    FrmSDetails.Height = 7200
    FrmSDetails.Width = 8790
    FrmSDetails.Top = (CFrmMain.ScaleHeight - FrmSDetails.ScaleHeight) / 2
    FrmSDetails.Left = (CFrmMain.ScaleWidth - FrmSDetails.ScaleWidth) / 2
End Sub


'Private Sub TxtSDetails_GotFocus(Index As Integer)
'    'display the sequence no
'    Dim code As String
'    Select Case Index
'
'        Case 0:
'                SeqGen.dataConnect
'                SeqGen.RecSet1.Open "select max(substr(course_code,2,5))+1 into code from course_details"
'                If code = Null Then
'                    TxtSDetails(0).Text = "C00001"
'                Else
'
'
'
'    End Select
'
'End Sub

Private Sub TxtSDetails_KeyPress(Index As Integer, KeyAscii As Integer)
    'display all the lower case letter into upper case

Select Case Index
        Case 0: 'unique id
        
                If KeyAscii > 92 And KeyAscii < 123 Then
                    KeyAscii = KeyAscii - 32
                End If
                
        Case 1: 'faculty
        
                If KeyAscii > 92 And KeyAscii < 123 Then
                    KeyAscii = KeyAscii - 32
                End If
                
        Case 2: 'student name
        
                If KeyAscii > 92 And KeyAscii < 123 Then
                    KeyAscii = KeyAscii - 32
                End If
                If KeyAscii > 46 And KeyAscii < 55 Then
                    KeyAscii = 0
                End If
                
        Case 3: 'Email
                
        Case 4: 'Father name
        
                If KeyAscii > 92 And KeyAscii < 123 Then
                    KeyAscii = KeyAscii - 32
                End If
                
        Case 5: 'sex
        
                If KeyAscii > 92 And KeyAscii < 123 Then
                    KeyAscii = KeyAscii - 32
                End If
                
        Case 6: 'date of birth
        
                If KeyAscii > 64 And KeyAscii < 123 Then
                    KeyAscii = 0
                End If
                
        Case 7: 'local Add
                If KeyAscii > 92 And KeyAscii < 123 Then
                    KeyAscii = KeyAscii - 32
                End If
                
        Case 8: 'Permanent Add
        
                If KeyAscii > 92 And KeyAscii < 123 Then
                    KeyAscii = KeyAscii - 32
                End If
                
        Case 9: 'date of adm
        
                If KeyAscii > 64 And KeyAscii < 123 Then
                    KeyAscii = 0
                End If
                
        Case 10: 'Education
                If KeyAscii > 92 And KeyAscii < 123 Then
                    KeyAscii = KeyAscii - 32
                End If
                
        Case 11: 'Lib card
        
                If KeyAscii > 92 And KeyAscii < 123 Then
                    KeyAscii = KeyAscii - 32
                End If
                
        Case 12: 'i-card
        
                If KeyAscii > 92 And KeyAscii < 123 Then
                    KeyAscii = KeyAscii - 32
                End If
                
End Select
        
End Sub

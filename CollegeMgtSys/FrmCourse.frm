VERSION 5.00
Begin VB.Form FrmCourse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Course Entry Module"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6765
   Icon            =   "FrmCourse.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   6765
   Begin VB.Frame Frame1 
      Caption         =   "New Course Entry Module"
      Height          =   4770
      Left            =   -15
      TabIndex        =   13
      Top             =   -30
      Width           =   6780
      Begin VB.ComboBox CmbCtype 
         Height          =   315
         Index           =   4
         ItemData        =   "FrmCourse.frx":0442
         Left            =   1890
         List            =   "FrmCourse.frx":044C
         TabIndex        =   4
         Text            =   "Yearly"
         Top             =   2250
         Width           =   1545
      End
      Begin VB.Frame Frame2 
         Caption         =   "Operations"
         Height          =   825
         Left            =   135
         TabIndex        =   23
         Top             =   3765
         Width           =   6525
         Begin VB.CommandButton CmdOperations 
            Caption         =   "Ca&ncel"
            Height          =   285
            Index           =   4
            Left            =   5265
            TabIndex        =   12
            Top             =   300
            Width           =   915
         End
         Begin VB.CommandButton CmdOperations 
            Caption         =   "&Ok"
            Height          =   285
            Index           =   3
            Left            =   4050
            TabIndex        =   11
            Top             =   300
            Width           =   915
         End
         Begin VB.CommandButton CmdOperations 
            Caption         =   "&Mordify"
            Height          =   285
            Index           =   2
            Left            =   2835
            TabIndex        =   10
            Top             =   300
            Width           =   915
         End
         Begin VB.CommandButton CmdOperations 
            Caption         =   "D&elete"
            Height          =   285
            Index           =   1
            Left            =   1575
            TabIndex        =   9
            Top             =   315
            Width           =   915
         End
         Begin VB.CommandButton CmdOperations 
            Caption         =   "&Add"
            Height          =   285
            Index           =   0
            Left            =   330
            TabIndex        =   8
            Top             =   300
            Width           =   915
         End
      End
      Begin VB.TextBox TxtDetails 
         Height          =   285
         Index           =   7
         Left            =   1890
         TabIndex        =   7
         Top             =   3345
         Width           =   2490
      End
      Begin VB.TextBox TxtDetails 
         Height          =   285
         Index           =   6
         Left            =   1890
         TabIndex        =   6
         Top             =   2985
         Width           =   4710
      End
      Begin VB.TextBox TxtDetails 
         Height          =   285
         Index           =   5
         Left            =   1890
         TabIndex        =   5
         Top             =   2625
         Width           =   1530
      End
      Begin VB.TextBox TxtDetails 
         Height          =   285
         Index           =   3
         Left            =   1890
         TabIndex        =   3
         Top             =   1920
         Width           =   1530
      End
      Begin VB.TextBox TxtDetails 
         Height          =   285
         Index           =   2
         Left            =   1890
         TabIndex        =   2
         Top             =   1560
         Width           =   4710
      End
      Begin VB.TextBox TxtDetails 
         Height          =   285
         Index           =   1
         Left            =   1890
         TabIndex        =   1
         Top             =   1215
         Width           =   4710
      End
      Begin VB.TextBox TxtDetails 
         Height          =   285
         Index           =   0
         Left            =   1890
         TabIndex        =   0
         Top             =   855
         Width           =   1530
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "No Of Installment"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   22
         Top             =   3420
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Course Faculty"
         Height          =   195
         Index           =   7
         Left            =   285
         TabIndex        =   21
         Top             =   3060
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Course Fees"
         Height          =   195
         Index           =   6
         Left            =   450
         TabIndex        =   20
         Top             =   2700
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Course Type"
         Height          =   195
         Index           =   5
         Left            =   435
         TabIndex        =   19
         Top             =   2340
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Duration"
         Height          =   195
         Index           =   4
         Left            =   735
         TabIndex        =   18
         Top             =   1995
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Eligibility"
         Height          =   195
         Index           =   3
         Left            =   750
         TabIndex        =   17
         Top             =   1620
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Course Name"
         Height          =   195
         Index           =   2
         Left            =   375
         TabIndex        =   16
         Top             =   1305
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Course Code"
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   15
         Top             =   945
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "New Course Entry Module"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Index           =   0
         Left            =   1620
         TabIndex        =   14
         Top             =   255
         Width           =   3330
      End
   End
End
Attribute VB_Name = "FrmCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Flag As String

Private Sub CHideShow(Index As Integer)
    Dim i As Integer
    'show and hide the buttons
    Select Case Index
        Case 0, 1, 2: 'hide is called
            
            
            'disable add,delete,modify button
            
            For i = 0 To 2 Step 1
                CmdOperations(i).Enabled = False
            Next i
        
        Case 3, 4: 'ok is pressed
            
            'enable add, delete, modify button
            
            For i = 0 To 2 Step 1
                CmdOperations(i).Enabled = True
            Next i
        
    End Select
        
End Sub

Private Sub CmdOperations_Click(Index As Integer)
    'selecting appropiate button
    Select Case Index
        Case 0: 'adding
            CHideShow (Index) 'hiding the command button
            Flag = "Add"
            'local variable declare
                            
            Dim i As Integer
            'CmbCtype(4).Text = ""
            
                For i = 0 To 3 Step 1
                    TxtDetails(i).Text = ""
                Next i
                For i = 5 To 7 Step 1
                    TxtDetails(i).Text = ""
                Next i
            SeqGen.dataConnect
            SeqGen.RecSet1.Open "select * from course_details" 'open the record set
            SeqGen.RecSet1.AddNew
            
            'MsgBox "Add button is clicked"
            
        Case 1: 'Deleting
            Flag = "Delete"
            CHideShow (Index)
            MsgBox "Delete Button is clicked"
        Case 2: 'Modifing
            Flag = "Modify"
            CHideShow (Index)
            MsgBox "Modify Button is clocked"
        Case 3: 'Commit the transaction
                    'perform the work of ok according to flag value
            
                    Select Case Flag
                        Case "Add": 'perform add work
                        
                            Dim j As Variant 'local variable
'                                RecSet1.Fields(4) = CmbCtype(4).Text
                            
                                For Each j In TxtDetails
                                'For j = 0 To 6 Step 1
                                    RecSet1.Fields(Index) = TxtDetails(j).Text
                                Next j
                                
'                                For j = 5 To 7 Step 1
'                                    RecSet1.Fields(j) = TxtDetails(j).Text
'                                Next j
                            
                                
                                RecSet1.Update 'updating the record set
                                RecSet1.Close 'closeing the record set
                                Flag = ""
                                
                        Case "Delete": 'perform delete work
                        Case "Modify": 'perform modify work
                    End Select
            CHideShow (Index)
        Case 4: 'Close The Windows
            CHideShow (Index)
            MsgBox "Close is Clicked"
    End Select
End Sub


Private Sub CmdOperations_KeyPress(Index As Integer, KeyAscii As Integer)
'display the letters in upper case.
    Select Case Index
        Case 0: 'course code
        
            If KeyAscii > 92 And KeyAscii < 123 Then
                KeyAscii = KeyAscii - 32
            End If
            
        Case 1: 'course name
        
            If KeyAscii > 92 And KeyAscii < 123 Then
                KeyAscii = KeyAscii - 32
            End If
            
            If KeyAscii > 46 And KeyAscii < 64 Then
                KeyAscii = 0
            End If
            
        Case 2: 'eligibility
        
            If KeyAscii > 92 And KeyAscii < 123 Then
                KeyAscii = KeyAscii - 32
            End If
            
        Case 3: 'duration
        
            If KeyAscii > 92 And KeyAscii < 123 Then
                KeyAscii = KeyAscii - 32
            End If
            
            If KeyAscii > 46 And KeyAscii < 64 Then
                KeyAscii = 0
            End If
                    
        Case 4: 'course type lived
        Case 5: 'course fees
        
            If KeyAscii < 64 And KeyAscii < 123 Then
                KeyAscii = 0
            End If
            
        Case 6: 'course faculty
        
            If KeyAscii > 92 And KeyAscii < 123 Then
                KeyAscii = KeyAscii - 32
            End If
            
        Case 7: 'no of installment
        
            If KeyAscii > 92 And KeyAscii < 123 Then
                KeyAscii = KeyAscii - 32
            End If
            
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Integer 'Local Variable
    'displaying the values at loading time
'    SeqGen.dataConnect
'    SeqGen.RecSet1.Open "Select course_code as AllRec from course_details"
'    If IsNull(SeqGen.RecSet1!AllRec) Then 'If no value found
'        SeqGen.RecSet1.Close
'        Exit Sub
'    Else
'        'SeqGen.RecSet1.MoveFirst
'        For i = 0 To 7 Step 1
'            TxtDetails(i).Text = SeqGen.RecSet1.Fields(i)
'        Next i
'    End If
    
End Sub

Private Sub TxtDetails_GotFocus(Index As Integer)
    'display the sequence no
    
    Select Case Index
            
            
'        Case 0:
'                If Flag = "Add" Then
'                        Dim StrLength As Integer 'to store string length
'                        Dim Retrive As String 'to store result from module
'                        Dim Convert As String 'to convert form integer to string
'
'                        SeqGen.dataConnect
'                        SeqGen.RecSet1.Open "select max(to_number(substr(course_code,2,5)))+1 as  code from course_details"
'                        If IsNull(RecSet1!code) Then
'                            TxtDetails(0).Text = "C00001"
'                        Else
'                            Convert = CStr(SeqGen.RecSet1!code)
'                            StrLength = Len(Convert)
'                            Retrive = SeqGen.SeqGen(StrLength)
'                            TxtDetails(0).Text = "C" & Retrive & Convert
'                        End If
'                End If
    End Select
End Sub

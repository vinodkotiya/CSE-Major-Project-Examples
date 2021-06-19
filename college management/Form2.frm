VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Course Modification/ Course Deletion"
   ClientHeight    =   3390
   ClientLeft      =   4125
   ClientTop       =   2595
   ClientWidth     =   5070
   LinkTopic       =   "Form2"
   ScaleHeight     =   3390
   ScaleWidth      =   5070
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Confirm"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Text            =   "(None)"
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Course  -  Id"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Mainform.Show
End Sub

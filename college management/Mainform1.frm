VERSION 5.00
Begin VB.MDIForm Mainform 
   BackColor       =   &H8000000C&
   Caption         =   "AMIT COLLEGE OF MANAGEMENT"
   ClientHeight    =   6330
   ClientLeft      =   135
   ClientTop       =   360
   ClientWidth     =   8100
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Menu cou 
      Caption         =   "Course"
      Begin VB.Menu couinf 
         Caption         =   "Information"
      End
   End
   Begin VB.Menu stu 
      Caption         =   "Student"
      Begin VB.Menu stuadm 
         Caption         =   "Admission"
      End
   End
   Begin VB.Menu emp 
      Caption         =   "Employee"
      Begin VB.Menu empinf 
         Caption         =   "Information"
         Begin VB.Menu empinfper 
            Caption         =   "Permanent"
         End
         Begin VB.Menu empinfvis 
            Caption         =   "Visiting"
         End
      End
   End
   Begin VB.Menu fee 
      Caption         =   "Fees"
      Begin VB.Menu feeinf 
         Caption         =   "Information"
      End
   End
   Begin VB.Menu enq 
      Caption         =   "Enquiry"
      Begin VB.Menu enqent 
         Caption         =   "Entry"
      End
      Begin VB.Menu enqinf 
         Caption         =   "Information"
         Begin VB.Menu enqinfenqid 
            Caption         =   "Enquiry-Id"
         End
         Begin VB.Menu enqinfcouid 
            Caption         =   "Course-Id"
         End
         Begin VB.Menu enqinfmon 
            Caption         =   "Month"
         End
      End
   End
   Begin VB.Menu rep 
      Caption         =   "Report"
   End
   Begin VB.Menu exi 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub couinf_Click()
form1.Show
End Sub

Private Sub empdel_Click()
Form8.Show
Form8.Caption = "Employee Deletion"
End Sub

Private Sub empinfper_Click()
Form9.Show
Form9.Caption = "Employee Information (Permanent)"
End Sub

Private Sub empinfvis_Click()
Form9.Show
Form9.Caption = "Employee Information (Visiting)"
End Sub

Private Sub empjoi_Click()
Form7.Show
End Sub

Private Sub empmod_Click()
Form8.Show
Form8.Caption = "Employee Modification"
End Sub

Private Sub enqent_Click()
Form12.Show
End Sub

Private Sub enqinfcouid_Click()
Form14.Show
End Sub

Private Sub enqinfenqid_Click()
Form13.Show
End Sub

Private Sub enqinfmon_Click()
Form15.Show
End Sub

Private Sub exi_Click()
End
End Sub

Private Sub feeent_Click()
Form10.Show
End Sub

Private Sub feeinf_Click()
Form11.Show
End Sub

Private Sub MDIForm_Load()
'Set db = OpenDatabase("A:\project\college.mdb")
Set db = OpenDatabase("C:\My Documents\visual basic projects\college management\college.mdb")
Set rs1 = db.OpenRecordset("course", dbOpenDynaset)
Set rs2 = db.OpenRecordset("student", dbOpenDynaset)
End Sub

Private Sub seapg_Click()
Form17.Show
End Sub

Private Sub seaug_Click()
Form16.Show
End Sub

Private Sub stuadm_Click()
Form4.Show
End Sub





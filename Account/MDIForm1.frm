VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Exit"
            ImageKey        =   "Exit"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6240
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3240
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0000
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0544
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0A88
            Key             =   "New"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0FCC
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1510
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1A54
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F98
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":24DC
            Key             =   "Stop"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2930
            Key             =   "Exit"
         EndProperty
      EndProperty
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu new 
         Caption         =   "&new"
         Shortcut        =   ^N
      End
      Begin VB.Menu open 
         Caption         =   "&open"
         Begin VB.Menu cust 
            Caption         =   "customer"
         End
         Begin VB.Menu inv 
            Caption         =   "invoice"
         End
         Begin VB.Menu inv_det 
            Caption         =   "invoice detail"
         End
         Begin VB.Menu pro 
            Caption         =   "product"
         End
         Begin VB.Menu int 
            Caption         =   "interest"
         End
      End
      Begin VB.Menu save 
         Caption         =   "&save"
         Shortcut        =   ^S
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu cut 
         Caption         =   "&cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "&copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "&paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu Repo 
      Caption         =   "&Report"
      Begin VB.Menu custrep 
         Caption         =   "Customer"
      End
      Begin VB.Menu intrep 
         Caption         =   "Interest"
      End
      Begin VB.Menu invrep 
         Caption         =   "Invoice detail"
      End
   End
   Begin VB.Menu wind 
      Caption         =   "&Window"
      Begin VB.Menu cascade 
         Caption         =   "Cascade"
      End
      Begin VB.Menu tile_hori 
         Caption         =   "Tile horizontal"
      End
      Begin VB.Menu tile_ver 
         Caption         =   "Tile verticle"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cascade_Click()
Me.Arrange vbCascade
End Sub
Private Sub copy_Click()
 On Error Resume Next
    Clipboard.SetText ActiveForm.ActiveControl.SelText
End Sub
Private Sub custrep_Click()
DataReport1.Show
End Sub
Private Sub cut_Click()
Clipboard.Clear
Clipboard.SetText ActiveForm.ActiveControl.SelText
ActiveForm.ActiveControl.SelText = ""
End Sub
Private Sub help_Click()
CommonDialog1.HelpFile = "windows.hlp"
CommonDialog1.HelpCommand = cdlHelpContext
CommonDialog1.ShowHelp
End Sub
Private Sub intrep_Click()
DataReport2.Show
End Sub
Private Sub invrep_Click()
DataReport3.Show
End Sub
Private Sub tile_hori_Click()
Me.Arrange vbTileHorizontal
End Sub
Private Sub paste_Click()
On Error Resume Next
    ActiveForm.ActiveControl.SelText = Clipboard.GetText()
End Sub
Private Sub tile_ver_Click()
Me.Arrange vbTileVertical
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    Select Case Button.Key
        Case "New"
            new_Click
        Case "Cut"
            cut_Click
        Case "Exit"
            exit_Click
        Case "Copy"
            copy_Click
        Case "Paste"
            paste_Click
    End Select
End Sub
Private Sub cust_Click()
form4.Show
Form1.Hide
Form2.Hide
Form3.Hide
Form6.Hide
End Sub
Private Sub exit_Click()
Unload Me
End Sub
Private Sub int_Click()
Form6.Show
Form1.Hide
Form2.Hide
form4.Hide
Form3.Hide
End Sub
Private Sub inv_Click()
Form1.Show
form4.Hide
Form2.Hide
Form3.Hide
Form6.Hide
End Sub
Private Sub inv_det_Click()
Form2.Show
Form1.Hide
form4.Hide
Form3.Hide
Form6.Hide
End Sub
Private Sub MDIForm_Load()
Form3.Show
Form1.Hide
Form2.Hide
form4.Hide
Form6.Hide
End Sub
Private Sub new_Click()
Form1.txtno.Text = ""
Form1.txtdate.Text = ""
Form1.txtprocode.Text = ""
Form1.txtcust.Text = ""
Form1.txtnam.Text = ""
Form1.txtunit.Text = ""
Form1.txtdes.Text = ""
Form1.txtnounit.Text = ""
Form1.txtinamt.Text = ""
Form1.txttax.Text = ""
Form1.txtamt.Text = ""
Form1.txtcash.Text = ""
Form2.txt_no.Text = ""
Form2.txt_date.Text = ""
Form2.txtcode.Text = ""
Form2.txtamt.Text = ""
Form2.Text2.Text = ""
Form2.Text1.Text = ""
Form3.Combo1.Text = ""
Form3.txtdes.Text = ""
Form3.txtprice.Text = ""
form4.txtcode.Text = ""
form4.txtadd.Text = ""
form4.txtname.Text = ""
form4.txtphone.Text = ""
form4.Combo1.Text = ""
form4.txtlimit.Text = ""
form4.txtdue.Text = ""
Form6.Text1.Text = ""
Form6.Combo1.Text = ""
Form6.txtint.Text = ""
Form6.txtper.Text = ""
Form6.Text2.Text = ""
Form6.txtlimit.Text = ""
Form6.Text3.Text = ""
Form6.Text4.Text = ""
Form3.Combo1.SetFocus
End Sub
Private Sub pro_Click()
Form3.Show
Form1.Hide
Form2.Hide
form4.Hide
Form6.Hide
End Sub



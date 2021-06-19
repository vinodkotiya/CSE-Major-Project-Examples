Attribute VB_Name = "Module2"

Option Explicit
'-----------------
'Global variables.
'=================

Global random_time As Integer ' Get timer info from form for my randomizer
Global strGuestSearchfor As String 'Reservations GuestSearch
Global NextFile As Integer 'Freefile Variable
Global Source, target As String 'COPYFILE function variables
Global log_msg As String 'File logger variables
Global log_file As String

'-----------------
'API Declarations.
'=================
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long

 '***Function to get current display settings
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwFlags As Long) As Long
'Restart computer for new display setting (optional)
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
  
'Use kernel32 for wait function
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

'-----------------
'Public constants.
'=================
' Horizontal bar const for listbox
Public Const LB_SETHORIZONTALEXTENT = &H194

'Constants for synchronize listboxes
'Note: Reservations tab, Reports
Public Const LB_SETTOPINDEX = &H197
Public Const LB_GETTOPINDEX = &H18E

'Encrypted Master Reset Key
Public Const MasterResetKey = "Š®´–³­"

Public Const log_sys = "log_system.txt"

'--------
'CONTROLS
'--------
'-----------------------------------------------------------------
'Allows only 0-9, "." and backspace to be keypressed in a textbox
'-----------------------------------------------------------------
Public Sub AllowOnlyIntegers(KeyAscii As Integer)
Const Numbers$ = "0123456789."
    If KeyAscii <> 8 Then
       If InStr(Numbers, Chr(KeyAscii)) = 0 Then
            MsgBox _
   "Only numbers allowed.", _
   vbOKOnly + vbInformation, _
   " "
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub
'-----------------------------------------------------------------
'Allows a-z, both cases and baclspace to be keypressed in a textbox
'-----------------------------------------------------------------
Public Sub AllowOnlyalpha(KeyAscii As Integer)
Const Numbers$ = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
    If KeyAscii <> 8 Then
       If InStr(Numbers, Chr(KeyAscii)) = 0 Then
            MsgBox _
   "Only alphabets allowed.", _
   vbOKOnly + vbInformation, _
   " "
            KeyAscii = 0
            Exit Sub
        End If
    End If
End Sub

'------------------------------------
'Clears all textboxes on certain form
'------------------------------------
Public Sub ClearTextBoxes(frmTarget As Form)
Dim i, ctrltarget
    For i = 0 To (frmTarget.Controls.count - 1)
        Set ctrltarget = frmTarget.Controls(i)


        If TypeOf ctrltarget Is TextBox Then
            ctrltarget.Text = ""
        End If
    Next i
End Sub

'----------------------
'Search item in listbox
'----------------------
Function searchfor(lst As ListBox, target As String) As Integer
Dim counter As Integer, found As Boolean

Let found = False

For counter = -1 To lst.ListCount
    If target = UCase(lst.List(counter)) Then
    searchfor = counter
    found = True
    End If
Next counter

If found = False Then searchfor = -9999

End Function

'-------------------------
'Add Scroll Bar to listbox
'-------------------------
Public Sub AddScroll(List As ListBox)
    Dim i As Integer, intGreatestLen As Integer, lngGreatestWidth As Long

    For i = 0 To List.ListCount - 1
        If Len(List.List(i)) > Len(List.List(intGreatestLen)) Then
            intGreatestLen = i
        End If
    Next i

    lngGreatestWidth = List.Parent.TextWidth(List.List(intGreatestLen) + Space(1))
    lngGreatestWidth = lngGreatestWidth \ Screen.TwipsPerPixelX
    SendMessage List.hwnd, LB_SETHORIZONTALEXTENT, lngGreatestWidth, 0
    
End Sub

'---------------------------
'Move 1 field up in list box
'---------------------------
Public Sub MoveFldUp(lb As ListBox)
Dim i As Integer
    i = lb.ListIndex


    If i > 0 And i < lb.ListCount Then
        lb.Selected(i - 1) = True
        lb.Selected(i) = False
    End If
End Sub

'---------------------------
'Move 1 field up in list box
'---------------------------
Public Sub MoveFldDown(lb As ListBox)
Dim i As Integer
    i = lb.ListIndex


    If i > -1 And i < lb.ListCount - 1 Then
        lb.Selected(i + 1) = True
        lb.Selected(i) = False
    End If
End Sub



'------------------------
'FILES AND DISK FUNCTIONS
'------------------------
'------------------------------------------
'Get filename from full path.
'truncate the path, and gets the filename.
'------------------------------------------
Public Function GetFile(ByVal s As String) As String
   Dim i As Integer
   Dim j As Integer
   
   i = 0
   j = 0
   
   i = InStr(s, "\")
   Do While i <> 0
      j = i
      i = InStr(j + 1, s, "\")
   Loop
   
   If j = 0 Then
      GetFile = ""
   Else
      GetFile = Right$(s, Len(s) - j)
   End If
End Function

'---------------------
'Checks if file exists
'----------------------
Public Function FileExist(ByVal fullpath As String) As Boolean
  On Error GoTo nofile
     Open fullpath For Binary Access Read As #1
     Close #1
     FileExist = True
     Exit Function
nofile:
     FileExist = False
End Function

'-----------
'File Copy
'---------
Public Sub CopyFile(ByVal sourcefile As String, ByVal destfile As String)
    Dim Bytearray() As Byte
    Dim filesize As Long
    
    Open sourcefile For Binary Access Read As #1
'    Open destfile For Binary Access Write As #2
    filesize = LOF(1)
    ReDim Bytearray(filesize)
    Get #1, , Bytearray
     'Put #2, , Bytearray
    Close 1
    Close 2
End Sub

'--------------
'MISC FUNCTIONS
'--------------
'-------
'Logger
'------
'Logs events into log file for future reference.

Public Sub logger(log_file As String, log_msg As String)
Open App.Path & "\" & log_file For Append As #1
       Print #1, log_msg
Close #1

End Sub

'-----------------------
'Random Number Generator
'-----------------------
Public Function random_number(max, random_time) As Integer
Dim count As Integer
  
  For count = 0 To max
     random_number = Int(count * Rnd)
     Randomize (random_time)
     Next count
End Function

'-------------------------------------------------------------
'Generate temp 8 digit password and username after MasterReset
'-------------------------------------------------------------
Function random_userpass(random_time) As Long
Dim i, n As Integer

random_userpass = random_number(9, random_time)
For i = 1 To 8
n = random_number(9, random_time)
random_userpass = random_userpass & n
Next i

End Function

'----------------
'A wait function
'----------------
'This method is pretty resource hungry, :P

Public Function Wait(ByVal TimeToWait As Long) 'Time In seconds
    Dim EndTime As Long
    EndTime = GetTickCount + TimeToWait * 1000 '* 1000 Cause u give seconds and GetTickCount uses Milliseconds
    Do Until GetTickCount > EndTime
        DoEvents
        Loop
    End Function


'-----------------------------------
'Error Handler
'-----------------------------------
Sub ErrHandler()
Dim errdesc As String, errnum As Integer

    errdesc = Err.Description
    errnum = Err.Number

'Log Errors
If errnum = 0 Then 'Error 0 is nothing, so don't do anything if there is an error 0
Else
log_msg = "ERROR: " & errnum & " - " & errdesc & " at " & Now
log_file = "Log_error.txt"
Call logger(log_file, log_msg)
End If

'Display error msg to user
Select Case Err.Number
      
    'File and Disk Errors
    Case 52    'Bad filename
        MsgBox "Bad filename! (erro 52)", vbExclamation
        
        
    Case 53     'File not found
        MsgBox "File not found! (erro 53)", vbExclamation
        
        
    Case 57     'Device I/O error
        If MsgBox("Destiny disk not ready! (erro 57)", vbCritical + vbYesNo) = vbYes Then Resume Next
    
    Case 58
        MsgBox "File already exists.", vbCritical
        
    
    Case 61     'Disk full
        If MsgBox("Destiny disk full! (error 61)", vbExclamation + vbYesNo) = vbYes Then Resume Next

    Case 68    'Drive not found
        MsgBox "The selected drive is not available.", vbCritical
        
    
    Case 70    'Permission denied
        If MsgBox("Destiny directory or drive protected! (error 70)", vbCritical + vbYesNo) = vbYes Then Resume Next
           
    Case 71    'Disk not ready
        If MsgBox("Destiny disk not ready! (error 71)", vbCritical + vbYesNo) = vbYes Then Resume Next
   
            
    Case 75
        MsgBox "Path/File access error.", vbCritical
        
        
    Case 76     'Path not found
        If MsgBox("Destiny directory unavailable! (error 76)", vbCritical + vbYesNo) = vbYes Then Resume Next
    
    Case 482
        MsgBox "Printer Error.", vbCritical
        
    
    'Database Errors
    Case 3004, 3024
        MsgBox "File not found.", vbCritical
       
    Case 3021
        MsgBox "Table is empty.", vbCritical
    
    Case 3022
        MsgBox "Duplicate key fields.", vbCritical
        
        
    
    Case 3058, 3315
        MsgBox "No entry in Keyfield.", vbCritical
        
        
    End Select
    
End Sub


Public Function EncryptText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String

    strPwd = UCase$(strPwd)

    'Encrypt string
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c + Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    EncryptText = strBuff
End Function


Public Function valid_Rs(txt As TextBox) As Boolean
Dim cnt
Dim cnt1
Dim stri
Dim j

cnt = 0
cnt1 = 0
    valid_Rs = True
   cnt = InStr(1, txt.Text, "..")
  If cnt <> 0 Then
    MsgBox "Two successive dots are not allowed.", vbCritical, "Validation"
    txt.SetFocus
    valid_Rs = False
    Exit Function
  End If

  stri = Mid(txt.Text, Len(txt.Text), 1)
  If stri = "." Then
    MsgBox "Wrong amount.", vbCritical, "Validation"
    txt.SetFocus
    valid_Rs = False
    Exit Function
  End If

cnt1 = 0
  For j = 1 To Len(txt.Text)
     stri = Mid(txt.Text, j, 1)
      If stri = "." Then
        cnt1 = cnt1 + 1
     End If
  Next
  If cnt1 > 1 Then
     MsgBox "Wrong amount", vbCritical, "Validation"
     cnt1 = 0
     txt.SetFocus
     valid_Rs = False
    Exit Function
 End If

End Function





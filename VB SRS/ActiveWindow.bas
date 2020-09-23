Attribute VB_Name = "ActiveWindow"
'Api for Getting the active window
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public sActiveWindow As String
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd _
    As Long) As Long

'Declaration for Speech Recognizer
Public speakar As New SpeechLib.SpVoice

'Find Window Api
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long


'Delcaration for Closing Window
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_CLOSE = &H10
'Const for Sending Keystroke and text to Yahoo
Private Const WM_SETTEXT = &HC
Private Const WM_KEYDOWN = &H100
Private Const VK_RETURN = &HD


'Starting run,find ,explore

  Public shlShell As New shell32.Shell
  
'Api for Shutdown, Reboot and Logoff
Const LOGOFF = 0
Const Shutdown = 1
Const REBOOT = 2
Const FORCE = 4
Const POWEROFF = 8
'APi for Opening Cd rom
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long




'Variables for Storing the Voice Commands
Public vCommands As Variant
Public SCommands As Variant
'Public vCommands() As String
'Public sCommands() As String
Public Sub ExecuteSystemCommands(sComm As String)
On Error GoTo erh
Select Case sComm

Case "run"
    shlShell.FileRun
Case "find"
    shlShell.FindFiles
Case "explorer"
    shlShell.Explore "c:\"
Case "shutdown"
    For Each objPC In GetObject("winmgmts:{(shutdown)}").ExecQuery("Select * from Win32_OperatingSystem ")
    objPC.Win32Shutdown Shutdown + FORCE
    Next

Case "reboot"
    For Each objPC In GetObject("winmgmts:{(shutdown)}").ExecQuery("Select * from Win32_OperatingSystem ")
    objPC.Win32Shutdown REBOOT + FORCE
    Next

Case "logoff"
    For Each objPC In GetObject("winmgmts:{(shutdown)}").ExecQuery("Select * from Win32_OperatingSystem ")
    objPC.Win32Shutdown LOGOFF + FORCE
    Next
Case "open cd-rom"
     mciSendString "set CDAudio door open", vbNullString, ByVal 0, ByVal 0
Case "close cd-rom"
    mciSendString "Set CDAudio Door Closed Wait", vbNullString, ByVal 0, ByVal 0
End Select
Exit Sub
erh:
'MsgBox Err.Description
End Sub

'Sub for Sending the text to Yahoo Messenger

Public Sub ToYahoo(x As String)
    
    Dim lParent As Long, lChild As Long
    Static sText As String
    
    'Get the Handle of Instant Messenger
    lParent = FindWindow("IMClass", vbNullString)
    If (lParent <> 0) Then
           'Get the Handle of Rich text Box of Instant Messenger
        lChild = FindWindowEx(lParent, 0, "RichEdit20A", "")
         If LCase$(x) = "backspace" Then sText = GetWordofString(sText): SendMessage lChild, WM_SETTEXT, ByVal 0, ByVal sText: Exit Sub
         If LCase$(x) = "send" Then SendMessage lChild, WM_KEYDOWN, VK_RETURN, ByVal 0: sText = "": Exit Sub
         
         sText = sText & " " & x
         
         'If the Number chars exceed 255 then automatically press Enter key
         If Len(sText) > 255 Then sText = "": SendMessage lChild, WM_KEYDOWN, VK_RETURN, ByVal 0
         
          SendMessage lChild, WM_SETTEXT, ByVal 0, ByVal sText
            
        
        Else
       ' MsgBox "Could not Find Yahoo Instant Messenger"
    End If
End Sub


Function GetWordofString(txt As String) As String

Dim lPos As Long

txt = StrReverse(txt)
lPos = InStr(1, txt, " ", vbBinaryCompare)
txt = Mid$(txt, lPos + 1, Len(txt))
txt = StrReverse(txt)
GetWordofString = txt
End Function






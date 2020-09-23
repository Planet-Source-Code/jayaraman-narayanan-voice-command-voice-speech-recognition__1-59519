VERSION 5.00
Begin VB.Form frmSpeech 
   Caption         =   "About Voice Commander"
   ClientHeight    =   2625
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   4980
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSpeech.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4980
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   4215
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Contact snj501@rediffmail.com"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   3735
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Jayaraman.N"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Voice Commander"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00984E00&
         Height          =   495
         Left            =   720
         TabIndex        =   1
         Top             =   480
         Width           =   3255
      End
   End
   Begin VB.Timer timActiveWindow 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3000
      Top             =   2520
   End
   Begin VB.Menu mnuSpeech 
      Caption         =   "&Speech"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mnuopen 
         Caption         =   "&About"
      End
      Begin VB.Menu mnudash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuchangeprofile 
         Caption         =   "&Change Profile"
      End
      Begin VB.Menu mnumictraining 
         Caption         =   "&Start Microphone &Traning"
      End
      Begin VB.Menu mnutrainingwizard 
         Caption         =   "Start Voice Training &Wizard"
      End
      Begin VB.Menu mnudash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuvoicecommands 
         Caption         =   "&Modify Voice Commands"
      End
      Begin VB.Menu mnuselectvoice 
         Caption         =   "&Select &Voice"
      End
      Begin VB.Menu mnudictation 
         Caption         =   "&Start &Dictation"
      End
      Begin VB.Menu mnudsh 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnableListening 
         Caption         =   "&Listen to Commands"
      End
      Begin VB.Menu mnuEnablewrite 
         Caption         =   "&Speak to Yahoo Messenger"
      End
      Begin VB.Menu mnuActiveWindow 
         Caption         =   "&Read Active Window Titles"
      End
      Begin VB.Menu mnudsh1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit SpeechReco"
      End
   End
End
Attribute VB_Name = "frmSpeech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents RC As SpSharedRecoContext
Attribute RC.VB_VarHelpID = -1
Dim myGrammar As ISpeechRecoGrammar
Dim sRecoString As String





Private Sub Form_Load()
'On Error GoTo erh

'Code for Adding icon to the System Tray
  With nid 'with system tray
            .cbSize = Len(nid)
            .hwnd = Me.hwnd
            .uID = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallbackMessage = WM_MOUSEMOVE
            .hIcon = Me.Icon 'use form's icon in tray
            .szTip = "Voice Commander" & vbNullChar 'tooltip text
        End With
        
    Shell_NotifyIcon NIM_ADD, nid 'add to tray
    
    
    'Start Speech Engine
    
    Set RC = New SpSharedRecoContext
    Set myGrammar = RC.CreateGrammar
    myGrammar.DictationSetState SGDSActive
    
    'RC.creat
    ' Set Voice for Speech
    'speakar.
                    
    'MsgBox GetSetting("SpeeechReco", "Voices", "Selected")
    Set speakar.Voice = speakar.GetVoices().Item(Val(GetSetting("SpeeechReco", "Voices", "Selected")))
    
    'When Run for the first Time
    If GetSetting("SpeeechReco", "SysCmds", "EnableListening") = "" Then SaveSetting "SpeeechReco", "SysCmds", "EnableListening", "false"
    If GetSetting("SpeeechReco", "SysCmds", "ReadActiveWindow") = "" Then SaveSetting "SpeeechReco", "SysCmds", "ReadActiveWindow", "false"
    If GetSetting("SpeeechReco", "SysCmds", "Yahoowrite") = "" Then SaveSetting "SpeeechReco", "SysCmds", "Yahoowrite", "false"
    
    
    'Get the Voice Commands and SystemCommands from Registry
    SCommands = GetAllSettings("SpeeechReco", "SysCmds")
    vCommands = GetAllSettings("SpeeechReco", "VCommands")
    
    'Check if the read out Action is Set
    Dim lCount As Long
    If Not IsEmpty(SCommands) Then
    For lCount = LBound(SCommands) To UBound(SCommands)
    
        If SCommands(lCount, 0) = "ReadActiveWindow" And SCommands(lCount, 1) = "True" Then timActiveWindow.Enabled = True: mnuActiveWindow.Checked = True: Exit For
        If SCommands(lCount, 0) = "Yahoowrite" And SCommands(lCount, 1) = "True" Then mnuEnablewrite.Checked = True
        If SCommands(lCount, 0) = "EnableListening" And SCommands(lCount, 1) = "True" Then mnuEnableListening.Checked = True
    Next
    End If
    'MsgBox UBound(SCommands)
    'Me.Hide
    Exit Sub
'erh:
 '   MsgBox Err.Description
End Sub


  Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim Result, Action As Long
    
    'Activate Popup only when clicked on the system Tray
    If Me.ScaleMode = vbPixels Then
        Action = x
    Else
        Action = x / Screen.TwipsPerPixelX
    End If
    
Select Case Action

    Case WM_LBUTTONDBLCLK 'Left Button Double Click
        Me.WindowState = vbNormal 'put into taskbar
            Result = SetForegroundWindow(Me.hwnd)
        Me.Show 'show form
    
    Case WM_RBUTTONUP 'Right Button Up
        Result = SetForegroundWindow(Me.hwnd)
        PopupMenu mnuSpeech 'popup menu, cool eh?
    
    End Select
    
End Sub

Private Sub Form_Resize()
If Me.WindowState = vbMinimized Then
Me.Hide
End If
End Sub

Private Sub Form_Unload(Cancel As Integer) 'on form unload
 Me.Hide
 Cancel = 1
'Call CloseAll
End Sub

Private Sub mnuActiveWindow_Click()
mnuActiveWindow.Checked = Not (mnuActiveWindow.Checked)
SaveSetting "SpeeechReco", "SysCmds", "ReadActiveWindow", mnuActiveWindow.Checked
SCommands = GetAllSettings("SpeeechReco", "SysCmds")

If mnuActiveWindow.Checked = True Then
 timActiveWindow.Enabled = True
Else
timActiveWindow.Enabled = False
End If

End Sub

Private Sub mnuchangeprofile_Click()
'Shell App.Path & "\sapi.cpl", vbNormalFocus
ShellExecute vbNull, "", App.Path & "\sapi.cpl", "", "", SW_SHOWNORMAL
End Sub

Private Sub mnudictation_Click()
frmDictation.Show
End Sub

Private Sub mnuEnableListening_Click()
mnuEnableListening.Checked = Not (mnuEnableListening.Checked)
SaveSetting "SpeeechReco", "SysCmds", "EnableListening", mnuEnableListening.Checked
If mnuEnableListening.Checked = True Then
  mnuEnablewrite.Checked = False
  SaveSetting "SpeeechReco", "SysCmds", "YahooWrite", "false"
End If
SCommands = GetAllSettings("SpeeechReco", "SysCmds")

End Sub

Private Sub mnuEnablewrite_Click()

mnuEnablewrite.Checked = Not (mnuEnablewrite.Checked)
SaveSetting "SpeeechReco", "SysCmds", "Yahoowrite", mnuEnablewrite.Checked
If mnuEnablewrite.Checked = True Then
  mnuEnableListening.Checked = False
  SaveSetting "SpeeechReco", "SysCmds", "EnableListening", "false"
End If
SCommands = GetAllSettings("SpeeechReco", "SysCmds")


End Sub

Private Sub mnuExit_Click() 'exit
If MsgBox("If you Exit Voice Commander, Voice features will be turned off. Do you want to do this?", vbQuestion + vbYesNo, "Voice Commander") = vbYes Then Call CloseAll
End Sub

Private Sub mnumictraining_Click()
RunUI SpeechMicTraining
End Sub

Private Sub mnuopen_Click()
Me.WindowState = 0: Me.Show
End Sub




Private Sub mnuselectvoice_Click()
frmVoice.Show
End Sub

Private Sub mnutrainingwizard_Click()
RunUI SpeechUserTraining
End Sub

Private Sub mnuvoicecommands_Click()
frmVoiceCommands.Show
End Sub

Private Sub timActiveWindow_Timer()
    Static lHwnd As Long
    Dim lCurHwnd As Long
    Dim sText As String * 255
    
    lCurHwnd = GetForegroundWindow
    If lCurHwnd = lHwnd Then Exit Sub
    lHwnd = lCurHwnd
    If lHwnd <> hwnd Then
    
        GetWindowText lHwnd, ByVal sText, 255
        
        If sActiveWindow <> sText Then
        
            speakar.Speak sText
            
         sActiveWindow = sText
        End If
        
    End If
    
End Sub

Private Sub RC_Recognition(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, ByVal RecognitionType As SpeechLib.SpeechRecognitionType, ByVal Result As SpeechLib.ISpeechRecoResult)
    
    'If in AppWrite Mode then
    If GetSetting("SpeeechReco", "SysCmds", "Yahoowrite") = "True" Then
    'Dim Vc As Variant, Sc As Variant
    
    sRecoString = Result.PhraseInfo.GetText
    ToYahoo sRecoString
    Exit Sub
    End If
    
    
    'Check if Listen to Command Enabled
    If GetSetting("SpeeechReco", "SysCmds", "EnableListening") <> "True" Then Exit Sub
    
    Dim lCurHwnd As Long, lCounter As Long
    Dim sText As String
    
    sRecoString = Result.PhraseInfo.GetText
    
       
    
    'Notify Command in the Task bar Display the Balloon
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE 'Or NIF_TIP 'NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Voice Commander" & vbNullChar
        .szInfo = sRecoString & Chr(0)
        .uTimeout = 10
        .szInfoTitle = "Recognised Word" & Chr(0)
        .dwInfoFlags = NIIF_INFO
        
         End With
   Shell_NotifyIcon NIM_MODIFY, nid
    
    
    'check if word is in Voice command list
    'MsgBox vCommands
    'MsgBox LBound(vCommands)
     SCommands = GetAllSettings("SpeeechReco", "SysCmds")
    vCommands = GetAllSettings("SpeeechReco", "VCommands")
    If Not IsEmpty(vCommands) Then
    For lCounter = Val(LBound(vCommands)) To Val(UBound(vCommands))
    If LCase$(sRecoString) = LCase$(vCommands(lCounter, 0)) Then
        Shell vCommands(lCounter, 1), vbNormalFocus
    End If
    Next
    End If
    'Check if Word is in System Commands list

    If Not IsEmpty(SCommands) Then
    For lCounter = Val(LBound(SCommands)) To Val(UBound(SCommands))
     If LCase$(sRecoString) = LCase$(SCommands(lCounter, 0)) And Val(SCommands(lCounter, 1)) = 1 Then
            ExecuteSystemCommands LCase$(sRecoString)

     End If
    Next
    End If
    ' Code for Closing a Window
    If sRecoString = "close" Then

    Static sTemp As String


    lCurHwnd = GetForegroundWindow
    If lCurHwnd = lHwnd Then Exit Sub
    lHwnd = lCurHwnd


        GetWindowText lHwnd, ByVal sText, 255

        'If sActiveWindow <> sText Then

            speakar.Speak "You are about to Close.   " & sText
    sTemp = sRecoString
    End If

    If sTemp = "close" And sRecoString = "yes" Then
    sTemp = ""
    lCurHwnd = GetForegroundWindow
    Debug.Print lCurHwnd
   Debug.Print SendMessage(lCurHwnd, WM_CLOSE, 0, 0)

    End If

End Sub

Private Sub RC_RecognitionForOtherContext(ByVal StreamNumber As Long, ByVal StreamPosition As Variant)
    'Label4.Caption = "For another context"
    'Label1.Caption = ""
End Sub
Sub CloseAll()

timActiveWindow.Enabled = False
Set speakar = Nothing
Set RC = Nothing
Set myGrammar = Nothing

Shell_NotifyIcon NIM_DELETE, nid 'remove from tray

End

End Sub

Private Function RunUI(theUI As String)
On Error GoTo erh
    If RC.Recognizer.IsUISupported(theUI) = True Then
        RC.Recognizer.DisplayUI Me.hwnd, "Voice Commander Additional Training", theUI, vbNullString
    End If
    Exit Function
erh:
    'MsgBox Err.Description
End Function


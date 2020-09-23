VERSION 5.00
Begin VB.Form frmVoice 
   Caption         =   "Select a Voice"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVoice.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   6030
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Test Voice"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   4560
      Width           =   1935
   End
   Begin VB.ListBox lstVoice 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Select a voice"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
End
Attribute VB_Name = "frmVoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private V As SpeechLib.SpVoice
Private T As SpeechLib.ISpeechObjectToken


Private Sub cmdSave_Click()

If lstVoice.ListIndex <> -1 Then
 SaveSetting "SpeeechReco", "Voices", "Selected", lstVoice.ListIndex
 MsgBox "Selection Saved.", vbInformation, "Voice Commander"
End If
End Sub

Private Sub cmdTest_Click()
On Error GoTo erh
If lstVoice.ListIndex > -1 Then
    
        'Set voice object to voice name selected in list box
        'The new voice speaks its own name
        
        Set V.Voice = V.GetVoices().Item(lstVoice.ListIndex)
        V.Speak V.Voice.GetDescription
        
    Else
        MsgBox "Please select a voice from the listbox", vbExclamation
    End If
Exit Sub
erh:
    'MsgBox Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo erh
    Dim strVoice As String
    
    Set V = New SpVoice
    
    'Get each token in the collection returned by GetVoices
    For Each T In V.GetVoices
        strVoice = T.GetDescription     'The token's name
        lstVoice.AddItem strVoice          'Add to listbox
    Next
    Exit Sub
erh:
    'MsgBox Err.Description
End Sub


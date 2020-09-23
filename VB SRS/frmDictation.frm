VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmDictation 
   Caption         =   "Reader"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10005
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDictation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   10005
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox rtbReader 
      Height          =   4095
      Left            =   600
      TabIndex        =   4
      Top             =   960
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7223
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDictation.frx":15162
   End
   Begin MSComDlg.CommonDialog cdDictate 
      Left            =   8640
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton txtLLoadText 
      Caption         =   "&Load Text File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton txtSaveText 
      Caption         =   "&Save Text File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   0
      Top             =   5640
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "SpeechReco Dictation"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   4695
   End
End
Attribute VB_Name = "frmDictation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents RC As SpSharedRecoContext
Attribute RC.VB_VarHelpID = -1
Private myGrammar As ISpeechRecoGrammar
Dim sRecoString As String


Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo erh
Set RC = New SpSharedRecoContext
Set myGrammar = RC.CreateGrammar
myGrammar.DictationSetState SGDSActive
Exit Sub
erh:
'MsgBox Err.Description

End Sub

Private Sub txtLLoadText_Click()
cdDictate.ShowOpen
rtbReader.LoadFile cdDictate.FileName
End Sub

Private Sub txtSaveText_Click()
cdDictate.ShowSave
rtbReader.SaveFile cdDictate.FileName, 1
End Sub



Private Sub RC_Recognition(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, ByVal RecognitionType As SpeechLib.SpeechRecognitionType, ByVal Result As SpeechLib.ISpeechRecoResult)
    
    
     rtbReader.Text = rtbReader.Text & " " & Result.PhraseInfo.GetText
    
        
End Sub

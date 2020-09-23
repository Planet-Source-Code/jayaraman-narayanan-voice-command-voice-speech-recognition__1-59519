VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmVoiceCommands 
   Caption         =   "Voice Commands"
   ClientHeight    =   7305
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9570
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVoiceCommands.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   9570
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Voice Commands"
      Height          =   6855
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5535
      Begin VB.ListBox lstCommands 
         Height          =   3570
         Left            =   840
         TabIndex        =   12
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtFilepath 
         Height          =   375
         Left            =   840
         TabIndex        =   11
         Top             =   5280
         Width           =   3855
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "..."
         Height          =   375
         Left            =   4920
         TabIndex        =   10
         Top             =   5280
         Width           =   375
      End
      Begin VB.TextBox txtCommand 
         Height          =   375
         Left            =   840
         TabIndex        =   9
         Top             =   4680
         Width           =   3855
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Settings"
         Height          =   495
         Left            =   3120
         TabIndex        =   8
         Top             =   5880
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Command"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Path"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Commands"
         Height          =   375
         Left            =   840
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "System Commands"
      Height          =   6735
      Left            =   5880
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.CheckBox chkCloseCdrom 
         Caption         =   "Close CDROM"
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Tag             =   "close cd-rom"
         Top             =   3000
         Width           =   1335
      End
      Begin VB.CheckBox chkOpenCdrom 
         Caption         =   "Open CDROM"
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Tag             =   "open cd-rom"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.CheckBox chkShutdown 
         Caption         =   "S&hutdown"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Tag             =   "shutdown"
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox chkReboot 
         Caption         =   "&Reboot"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Tag             =   "reboot"
         Top             =   855
         Width           =   1335
      End
      Begin VB.CheckBox chkLogoff 
         Caption         =   "&Logoff"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Tag             =   "logoff"
         Top             =   1230
         Width           =   1335
      End
      Begin VB.CheckBox chkRun 
         Caption         =   "&Run"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Tag             =   "run"
         Top             =   1605
         Width           =   1335
      End
      Begin VB.CheckBox chkFind 
         Caption         =   "&Find"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Tag             =   "find"
         Top             =   1980
         Width           =   1335
      End
      Begin VB.CheckBox chkExplore 
         Caption         =   "&Explorer"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Tag             =   "explorer"
         Top             =   2280
         Width           =   1335
      End
   End
   Begin MSComDlg.CommonDialog cdVoicemd 
      Left            =   0
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Executables Files (*.exe)|*.exe"
   End
End
Attribute VB_Name = "frmVoiceCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkCloseCdrom_Click()
SaveSysCmdSetting chkCloseCdrom
End Sub

Private Sub chkExplore_Click()
SaveSysCmdSetting chkExplore
End Sub

Private Sub chkFind_Click()
SaveSysCmdSetting chkFind
End Sub

Private Sub chkLogoff_Click()
SaveSysCmdSetting chkLogoff
End Sub


Private Sub chkOpenCdrom_Click()
SaveSysCmdSetting chkOpenCdrom
End Sub

Private Sub chkReboot_Click()
SaveSysCmdSetting chkReboot
End Sub

Private Sub chkRun_Click()
SaveSysCmdSetting chkRun
End Sub

Private Sub chkShutdown_Click()
'MsgBox chkShutdown.Value
SaveSysCmdSetting chkShutdown
End Sub



Private Sub cmdBrowse_Click()
cdVoicemd.ShowOpen
txtFilepath = cdVoicemd.FileName
End Sub

Private Sub cmdSave_Click()

If txtCommand <> "" And txtFilepath <> "" Then SaveSetting "SpeeechReco", "VCommands", txtCommand, txtFilepath
Call GetallCommands
End Sub

Sub GetallCommands()

Dim SCommands As Variant, iSet As Integer
lstCommands.Clear
SCommands = GetAllSettings("SpeeechReco", "VCommands")
If Not IsEmpty(SCommands) Then
For iSet = LBound(SCommands, 1) To UBound(SCommands, 1)
   lstCommands.AddItem SCommands(iSet, 0)
Next iSet

End If
End Sub

Private Sub Form_Load()
Call GetallCommands
Call LoadSysCmdSetting
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Refresh Activated Commands
SCommands = GetAllSettings("SpeeechReco", "SysCmds")
vCommands = GetAllSettings("SpeeechReco", "VCommands")
End Sub

Private Sub lstCommands_Click()
txtCommand = lstCommands.List(lstCommands.ListIndex)
txtFilepath = GetSetting("SpeeechReco", "VCommands", lstCommands.List(lstCommands.ListIndex))
End Sub

Sub SaveSysCmdSetting(chk As CheckBox)

SaveSetting "SpeeechReco", "SysCmds", chk.Tag, chk.Value

End Sub
Sub LoadSysCmdSetting()
On Error Resume Next
For Each Control In Me
If TypeOf Control Is CheckBox Then

Control.Value = GetSetting("SpeeechReco", "SysCmds", Control.Tag)

End If

Next
End Sub

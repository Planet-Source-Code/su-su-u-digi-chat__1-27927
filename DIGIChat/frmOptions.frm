VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfig 
   BackColor       =   &H00404040&
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   Picture         =   "frmOptions.frx":0000
   ScaleHeight     =   5325
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "O.K."
      Height          =   495
      Left            =   5280
      TabIndex        =   33
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton cmdEditBGColors 
      Caption         =   "Edit Background Colors"
      Height          =   615
      Left            =   5280
      TabIndex        =   32
      Top             =   120
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog cdBGPicture 
      Left            =   5160
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2295
      Left            =   240
      TabIndex        =   21
      Top             =   120
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   4048
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BackColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Chat Window"
      TabPicture(0)   =   "frmOptions.frx":0342
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame5"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Main Window"
      TabPicture(1)   =   "frmOptions.frx":035E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame6 
         Caption         =   "Main Window Background Color"
         Height          =   1815
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   4575
         Begin VB.TextBox txtBGPicture 
            Height          =   285
            Left            =   120
            TabIndex        =   31
            Top             =   1440
            Width           =   3615
         End
         Begin VB.CommandButton cmdBGPicture 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3840
            TabIndex        =   30
            Top             =   1440
            Width           =   615
         End
         Begin VB.OptionButton optMainWinBG 
            Caption         =   "Select a picture :"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   29
            Top             =   960
            Width           =   1695
         End
         Begin VB.OptionButton optMainWinBG 
            Caption         =   "Random Gradient"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   1935
         End
         Begin VB.OptionButton optMainWinBG 
            Caption         =   "Standard"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Chat Window Background Color"
         Height          =   1815
         Left            =   -74880
         TabIndex        =   22
         Top             =   360
         Width           =   4575
         Begin VB.TextBox txtChatWinColor 
            Height          =   285
            Left            =   120
            TabIndex        =   25
            Top             =   1080
            Width           =   2535
         End
         Begin VB.CommandButton cmdRTBBackColor 
            Caption         =   "Change Color"
            Height          =   375
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Chosen Color :"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   1335
         End
      End
   End
   Begin VB.CheckBox chkMinimizeAlert 
      BackColor       =   &H00404040&
      Caption         =   "Incoming message alert when chat window is minimized."
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   2040
      Width           =   4455
   End
   Begin VB.ComboBox cmbVoices 
      Height          =   315
      Left            =   1080
      TabIndex        =   13
      Top             =   3960
      Width           =   3855
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   1920
      TabIndex        =   11
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Options"
      Height          =   375
      Left            =   5400
      TabIndex        =   7
      Top             =   3240
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog cd3 
      Left            =   5160
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "Popup Message Alert"
      ForeColor       =   &H00000040&
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   4815
      Begin VB.TextBox txtAlertSound 
         Height          =   345
         Left            =   960
         TabIndex        =   12
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkMsgAlert 
         BackColor       =   &H00404040&
         Caption         =   "On"
         ForeColor       =   &H8000000F&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdChangeSound 
         Caption         =   "Change"
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Text-To-Speech"
      ForeColor       =   &H00000040&
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   3360
      Width           =   4815
      Begin VB.HScrollBar hsSpeed 
         Height          =   255
         Left            =   840
         Max             =   200
         Min             =   20
         TabIndex        =   14
         Top             =   1080
         Value           =   20
         Width           =   3855
      End
      Begin VB.CheckBox chkTTS 
         BackColor       =   &H00404040&
         Caption         =   "Enable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Speed :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label lblVoices 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Voices :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.TextBox txtNick 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Text            =   "Anonymous"
      Top             =   480
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      Caption         =   "Incoming Message Alert"
      ForeColor       =   &H00000040&
      Height          =   615
      Left            =   240
      TabIndex        =   16
      Top             =   1800
      Width           =   4815
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00404040&
      Caption         =   "Personal Details"
      ForeColor       =   &H00000040&
      Height          =   1575
      Left            =   240
      TabIndex        =   17
      Top             =   120
      Width           =   4815
      Begin VB.Label lblNick 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Nick Name :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Email          :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Local IP      :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkMsgAlert_Click()
'Sound alert or not?
If chkMsgAlert.Value = 1 Then
    txtAlertSound.Enabled = True
Else
    txtAlertSound.Enabled = False
End If

End Sub

Private Sub cmdBGPicture_Click()
'Let the user choose his/her own pic to be the background
On Error Resume Next
If optMainWinBG(2).Value <> True Then
    MsgBox "You have to choose the 'Custom Picture' option to select a picture.", vbInformation
Else
    cdBGPicture.Filter = "Picture Files (*.bmp, *.jpg, *.gif) | *.bmp;*.jpg;*.gif"
    cdBGPicture.ShowOpen
    frmMain.imgBg.Visible = False
    frmMain.imgCustomBG.Picture = LoadPicture(cdBGPicture.FileName)
    txtBGPicture.Text = cdBGPicture.FileName
    If txtBGPicture.Text <> "" Then
        Call CopyFile(txtBGPicture.Text, App.Path & "/" & cdBGPicture.FileTitle, False)
        txtBGPicture.Text = Right(txtBGPicture.Text, Len(txtBGPicture) - InStr(txtBGPicture.Text, "/"))
        txtBGPicture.Refresh
        
    End If
    
End If

    
End Sub

Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub cmdChangeSound_Click()
'Change the sound to play when msg is received
Dim WaveFile
cd3.Filter = "WAV files(*.wav) | *.wav"
cd3.ShowOpen
If cd3.FileName = "" Then
    WaveFile = App.Path & "/beep.wav"
Else
    WaveFile = cd3.FileName
End If
txtAlertSound.Text = WaveFile

End Sub

Private Sub cmdEditBGColors_Click()

'Different options on choosing the background of frmMain

SSTab1.Visible = True

If optMainWinBG(0).Value = True Then
    optMainWinBG(1).Value = False
    optMainWinBG(2).Value = False
    txtBGPicture.Enabled = False
ElseIf optMainWinBG(1).Value = True Then
    optMainWinBG(0).Value = False
    optMainWinBG(2).Value = False
    txtBGPicture.Enabled = False
ElseIf optMainWinBG(2).Value = True Then
    optMainWinBG(0).Value = False
    optMainWinBG(1).Value = False
    txtBGPicture.Enabled = True
End If

Frame1.Visible = False
Frame2.Visible = False
cmdSave.Visible = False
cmdCancel.Visible = False
cmbVoices.Visible = False
cmdEditBGColors.Visible = False
cmdOK.Visible = True


End Sub

'Back to main configurations
Private Sub cmdOK_Click()
SSTab1.Visible = False
Frame1.Visible = True
Frame2.Visible = True
cmdSave.Visible = True
cmdCancel.Visible = True
cmdEditBGColors.Visible = True
cmdOK.Visible = False
cmbVoices.Visible = True


End Sub

Private Sub cmdSave_Click()
'Select the path to save the settings in
iniPath$ = App.Path & "/Settings.ini"
Dim WaveFile
If cd3.FileName = "" Then
    WaveFile = App.Path & "/beep.wav"
Else
    WaveFile = cd3.FileName
End If

'Go through each entry of settings on the form and save
'them to .ini file
entry$ = txtNick.Text
RetVal = WritePrivateProfileString("DIGIChat", "NickName", entry$, iniPath$)
entry$ = txtEmail.Text
RetVal = WritePrivateProfileString("DIGIChat", "Email", entry$, iniPath$)
entry$ = cmbVoices.Text
RetVal = WritePrivateProfileString("DIGIChat", "TTSVoice", entry$, iniPath$)
entry$ = CStr(hsSpeed.Value)
RetVal = WritePrivateProfileString("DIGIChat", "TTSSpeed", entry$, iniPath$)
entry$ = CStr(chkMsgAlert.Value)
RetVal = WritePrivateProfileString("DIGIChat", "MsgAlert", entry$, iniPath$)
entry$ = CStr(chkTTS.Value)
RetVal = WritePrivateProfileString("DIGIChat", "TTSOn", entry$, iniPath$)
entry$ = WaveFile
RetVal = WritePrivateProfileString("DIGIChat", "WaveFile", entry$, iniPath$)
entry$ = CStr(chkMinimizeAlert.Value)
RetVal = WritePrivateProfileString("DIGIChat", "MinimizeAlert", entry$, iniPath$)
entry$ = txtChatWinColor.Text
RetVal = WritePrivateProfileString("DIGIChat", "ChatWinBGColor", entry$, iniPath$)
entry$ = optMainWinBG(0).Value
RetVal = WritePrivateProfileString("DIGIChat", "optMainWinBG(0)", entry$, iniPath$)
entry$ = optMainWinBG(1).Value
RetVal = WritePrivateProfileString("DIGIChat", "optMainWinBG(1)", entry$, iniPath$)
entry$ = txtBGPicture.Text
RetVal = WritePrivateProfileString("DIGIChat", "CustomPictureFile", entry$, iniPath$)


Me.Hide

End Sub
Private Sub Form_Load()


SSTab1.Visible = False
cmdOK.Visible = False

'Go through the settings saved in the .ini file and load
'them when frmConfig is shown

Dim intVoices As Integer
Dim intCountEngine As Integer
iniPath$ = App.Path & "/Settings.ini"
txtNick.Text = GetFromINI("DIGIChat", "NickName", iniPath$)
txtEmail.Text = GetFromINI("DIGIChat", "Email", iniPath$)
txtIP.Text = frmMain.sckConnect.LocalIP
chkMsgAlert.Value = Val(GetFromINI("DIGIChat", "MsgAlert", iniPath$))
txtAlertSound.Text = GetFromINI("DIGIChat", "WaveFile", iniPath$)
If Len(Dir$(txtAlertSound.Text)) <= 0 Then txtAlertSound.Text = App.Path & "/Beep.wav"
chkTTS.Value = Val(GetFromINI("DIGIChat", "TTSOn", iniPath$))
intVoices = frmMain.TTS.CountEngines
For intCountEngine = 1 To intVoices
    cmbVoices.AddItem frmMain.TTS.ModeName(intCountEngine), intCountEngine - 1
Next intCountEngine
cmbVoices.Text = GetFromINI("DIGIChat", "TTSVoice", iniPath$)


hsSpeed.Value = CInt(GetFromINI("DIGIChat", "TTSSpeed", iniPath$))


If chkMsgAlert.Value = 0 Then
    txtAlertSound.Enabled = False
Else
    txtAlertSound.Enabled = True
End If

txtIP.Enabled = False



chkMinimizeAlert.Value = Val(GetFromINI("DIGIChat", "MinimizeAlert", iniPath$))
txtChatWinColor.Text = GetFromINI("DIGIChat", "ChatWinBGColor", iniPath$)
optMainWinBG(0).Value = GetFromINI("DIGIChat", "optMainWinBG(0)", iniPath$)
optMainWinBG(1).Value = GetFromINI("DIGIChat", "optMainWinBG(1)", iniPath$)
txtBGPicture.Text = GetFromINI("DIGIChat", "CustomPictureFile", iniPath$)


    


End Sub


Private Sub cmdRTBBackColor_Click()
'Choose the bg color for the chat window
cd3.ShowColor
frmMain.rtbChat.BackColor = cd3.Color
txtChatWinColor.Text = cd3.Color


End Sub
Private Sub optMainWinBG_Click(Index As Integer)
    'Handle results of different clicks
If optMainWinBG(0).Value = True Then
    optMainWinBG(1).Value = False
    optMainWinBG(2).Value = False
    txtBGPicture.Enabled = False
    frmMain.BackColor = &H400000
    frmMain.imgCustomBG.Visible = False
    frmMain.imgBg.Visible = True
    GradientBG = False
    
ElseIf optMainWinBG(1).Value = True Then
    optMainWinBG(0).Value = False
    optMainWinBG(2).Value = False
    txtBGPicture.Enabled = False
    frmMain.imgCustomBG.Visible = False
    frmMain.imgBg.Visible = False
    GradientBG = True
    
    
ElseIf optMainWinBG(2).Value = True Then
    optMainWinBG(0).Value = False
    optMainWinBG(1).Value = False
    txtBGPicture.Enabled = True
    frmMain.BackColor = &H400000
    frmMain.imgCustomBG.Picture = LoadPicture(txtBGPicture.Text)
    frmMain.BackColor = &H400000
    GradientBG = False
    
End If

    
End Sub

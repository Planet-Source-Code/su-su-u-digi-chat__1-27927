VERSION 5.00
Object = "{2398E321-5C6E-11D1-8C65-0060081841DE}#1.0#0"; "VTEXT.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00400000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DIGI Chat"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   5865
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtbText 
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":08CA
   End
   Begin RichTextLib.RichTextBox rtbChat 
      Height          =   2055
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3625
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmMain.frx":094C
   End
   Begin VB.PictureBox picSysTrayIcon 
      Height          =   495
      Left            =   5400
      Picture         =   "frmMain.frx":09CE
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   6
      Top             =   3480
      Width           =   495
   End
   Begin HTTSLibCtl.TextToSpeech TTS 
      Height          =   375
      Left            =   3360
      OleObjectBlob   =   "frmMain.frx":1298
      TabIndex        =   7
      Top             =   0
      Width           =   495
   End
   Begin VB.CheckBox chkBold 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3600
      Picture         =   "frmMain.frx":12BC
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.CheckBox chkItalic 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   3900
      Picture         =   "frmMain.frx":15FE
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3120
      Width           =   315
   End
   Begin VB.CheckBox chkUnderline 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4200
      Picture         =   "frmMain.frx":1940
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3120
      Width           =   315
   End
   Begin VB.CommandButton cmdColors 
      Height          =   315
      Left            =   3240
      Picture         =   "frmMain.frx":1C82
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3120
      Width           =   315
   End
   Begin VB.ComboBox cmbFonts 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3120
      Width           =   3000
   End
   Begin MSWinsockLib.Winsock sckConnect 
      Left            =   2880
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1300
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2280
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      Caption         =   "Connection Options"
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   5535
      Begin VB.OptionButton optServerClient 
         BackColor       =   &H00400000&
         Caption         =   "Client"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtIP 
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   480
         Width           =   2175
      End
      Begin VB.OptionButton optServerClient 
         BackColor       =   &H00400000&
         Caption         =   "Host"
         ForeColor       =   &H8000000E&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.Label cmdConnect 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Connect"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4080
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblIP 
         BackColor       =   &H00400000&
         Caption         =   "IP Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Image imgMin 
      Height          =   255
      Left            =   5400
      Picture         =   "frmMain.frx":1FC4
      Top             =   0
      Width           =   255
   End
   Begin VB.Image imgClose 
      Height          =   255
      Left            =   5640
      Picture         =   "frmMain.frx":237A
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblSend 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Dialogue :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Image imgBg 
      Height          =   3975
      Left            =   0
      Picture         =   "frmMain.frx":2730
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8655
   End
   Begin VB.Image imgCustomBG 
      Height          =   3975
      Left            =   0
      Top             =   0
      Width           =   5895
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu itmOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu itmSaveAs 
         Caption         =   "&Save As"
      End
      Begin VB.Menu itmSep1 
         Caption         =   "-"
      End
      Begin VB.Menu itmPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu itmSep2 
         Caption         =   "-"
      End
      Begin VB.Menu itmExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu itmCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu itmCut 
         Caption         =   "Cut"
      End
      Begin VB.Menu itmPaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu itmSep8 
         Caption         =   "-"
      End
      Begin VB.Menu itmClear 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu itmView 
      Caption         =   "View"
      Begin VB.Menu itmOptions 
         Caption         =   "Options"
      End
      Begin VB.Menu itmSep6 
         Caption         =   "-"
      End
      Begin VB.Menu itmMinimize 
         Caption         =   "Minimize"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu itmFileTransfer 
         Caption         =   "File Transfer"
      End
      Begin VB.Menu itmSep3 
         Caption         =   "-"
      End
      Begin VB.Menu itmWhiteboard 
         Caption         =   "Whiteboard"
      End
      Begin VB.Menu itmSep9 
         Caption         =   "-"
      End
      Begin VB.Menu itmSendPopupMessage 
         Caption         =   "Send Popup Message"
      End
      Begin VB.Menu itmSep14 
         Caption         =   "-"
      End
      Begin VB.Menu itmGetProfile 
         Caption         =   "Get Other Side's Profile"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu itmViewHelpFile 
         Caption         =   "View Help File"
      End
      Begin VB.Menu itmSep7 
         Caption         =   "-"
      End
      Begin VB.Menu itmAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "^_^"
      Visible         =   0   'False
      Begin VB.Menu itmRestore 
         Caption         =   "Restore"
      End
      Begin VB.Menu itmMin 
         Caption         =   "Minimize"
      End
      Begin VB.Menu itmSep10 
         Caption         =   "-"
      End
      Begin VB.Menu itmConnect 
         Caption         =   "Connect"
      End
      Begin VB.Menu itmListen 
         Caption         =   "Listen"
      End
      Begin VB.Menu itmDisconnect 
         Caption         =   "Disconnect"
      End
      Begin VB.Menu itmSep11 
         Caption         =   "-"
      End
      Begin VB.Menu itmSendFile 
         Caption         =   "Send a File"
      End
      Begin VB.Menu itmSendPopup 
         Caption         =   "Send Popup Message"
      End
      Begin VB.Menu itmSep13 
         Caption         =   "-"
      End
      Begin VB.Menu itmQuit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Constants for displaying colors in RTB
Private Const vbDarkRed = &H80&
Private Const vbDarkGreen = &H8000&

'Constants for error handling for the Cut, Copy and paste Procedures
Const iCB_CLEAR_ERR = 3 + vbObjectError + 512
Const iCB_SET_ERR = 4 + vbObjectError + 512
Const iCB_PASTE_ERR = 5 + vbObjectError + 512




Sub AddChat(strNick As String, strRTF As String)


''Adds someone's nick and what they said to rtbChat

  Dim lngLastLen As Long
  
  ''set selected position to length of text
  rtbChat.SelStart = Len(rtbChat.Text)
  rtbChat.SelLength = 0
  
  ''set the seltext to a new line plus "Nick:" and tab character
  rtbChat.SelText = vbCrLf & strNick$ & ":" & vbTab
  
  ''change color, size, font name, and font styles
  rtbChat.SelStart = Len(rtbChat.Text) - (Len(strNick$) + 4) '4 = Length of vbCrLf + ':' + vbTab
  rtbChat.SelLength = Len(strNick$) + 4
  rtbChat.SelColor = vbBlue
  rtbChat.SelFontSize = 8
  rtbChat.SelFontName = "Arial"
  rtbChat.SelBold = True
  rtbChat.SelUnderline = False
  rtbChat.SelItalic = False
  
  ''store length of text so we can have a hangingindent later
  lngLastLen& = Len(rtbChat.Text)
  
  ''set selstart & sellength then add the rtf string
  rtbChat.SelStart = lngLastLen&
  rtbChat.SelLength = 0
  rtbChat.SelRTF = strRTF$
  
  ''now set the hanging indent
  rtbChat.SelStart = lngLastLen&
  rtbChat.SelLength = Len(rtbChat.Text) - lngLastLen&
  rtbChat.SelHangingIndent = 1400
  
  ''scroll textbox down
  rtbChat.SelStart = Len(rtbChat.Text)
  rtbChat.SelLength = 0
  
  ''set focus to rtbText
  If frmMain.Visible = True Then
    rtbText.SetFocus
  End If
  End Sub
Sub ParseData(strData As String)

'Declare the commonly used variables:
Dim strCommand As String
Dim strArgument As String
Dim strBuffer1 As String
Dim strBuffer2 As String

'Set strArgument the actual data and strCommand the command sent from the other side
strCommand = Left(strData, InStr(strData, " ") - 1)
strArgument = Right(strData$, Len(strData) - InStr(strData, " "))

'Take action depending on command received
Select Case strCommand

    'Normal msg
    Case "[MESSAGE]"
   
        strBuffer1 = Left(strArgument, InStr(strArgument, ":") - 1)
        strBuffer2 = Right(strArgument, Len(strArgument) - InStr(strArgument, ":"))
        AddChat strBuffer1, strBuffer2
        On Error Resume Next
        If frmConfig.chkTTS.Value = 1 Then
            TTS.Select (frmConfig.cmbVoices.ListIndex + 1)
            TTS.Speed = frmConfig.hsSpeed.Value
            TTS.Speak strBuffer2
        End If
        
        Dim RetVal
        If frmMain.Visible = False Then
            RetVal = FlashWindow(Me.hWnd, 1)
            Sleep 500
            RetVal = FlashWindow(Me.hWnd, 0)
            Sleep 500
            RetVal = FlashWindow(Me.hWnd, 1)
            Sleep 500
            RetVal = FlashWindow(Me.hWnd, 0)
            Sleep 500
            
            If frmConfig.chkMinimizeAlert.Value = 1 Then
                frmPopup.lblMsgDisplay.Caption = strArgument
                frmPopup.Show
            End If
        End If
    
    'New connection received
    Case "[JOIN]"
        If strArgument = frmConfig.txtNick.Text Then
            sckConnect.SendData "[ERR_NICKINUSE] "
            Exit Sub
        End If
        AddSysMessage vbCrLf & "*** " & strArgument & " has joined the chat.", RGB(15, 181, 0)
    Case "[LEAVE]"
        sckConnect.Close
        AddSysMessage vbCrLf & "*** " & strArgument & " has left the chat.", RGB(15, 181, 0)
        If optServerClient(0).Value = True Then
            cmdConnect.Caption = "Listen"
        Else
            cmdConnect.Caption = "Connect"
        End If
        optServerClient(0).Enabled = True
        optServerClient(1).Enabled = True
        txtIP.Enabled = True
        
    'Nickname used by another person in chat session
    Case "[ERR_NICKINUSE]"
        Do
            strBuffer1 = InputBox("The nickname " & strArgument & "is currently used by the other person, please choose another one:", "Error : Nickname already in use")
        Loop Until Trim(strBuffer1) <> ""
              
        frmConfig.txtNick.Text = strBuffer1
        
        sckConnect.SendData "[JOIN] " & strBuffer1
   
   
        
    'File transfer requested by other side
    Case "[SENDFILE]"
    
       
        
        FileSize = Val(Mid(strArgument, 1, Len(strArgument) - Len(Left(strArgument, InStr(strArgument, "_") - 1))))
        FileName = Mid(strArgument, InStr(1, "_") + 1)
        
        Dim Question As String
        Dim Answer As VbMsgBoxResult
      
    
        RetVal = sndPlaySound(App.Path & "/IncomingFileTransfer.wav", SND_ASYNC)
    
      
      Question = "The remote computer wishes to send you this file:" & vbCrLf & _
               FileName & " (" & FileSize & " bytes)" & vbCrLf & vbCrLf & _
               "Recieve this file? "
      Answer = MsgBox(Question, vbInformation Or vbYesNo)
      
     
      If Answer = vbYes Then
         'Prepare for the file transfer
         frmFileTransfer.fraSending.Caption = "Receiving file " & FileName
         frmFileTransfer.cmdSend.Enabled = False
          frmFileTransfer.lblTransferStatus.Caption = "Received 0 bytes (0%)"
          frmFileTransfer.ProgressBar1.Max = FileSize
          frmFileTransfer.ProgressBar1.Value = 0
          frmFileTransfer.fraSending.Visible = True
         RecievedFile = ""
       
         sckConnect.SendData "[SENDFILE_Y] "
         BeginTransfer = Timer
      Else
        
         sckConnect.SendData "[SENDFILE_N] "
      End If
    
    'Other side accepted the file, send it!
    Case "[SENDFILE_Y]"
        SendingFile = True
        frmFileTransfer.Show
        
    'Other side refused to accept the file
    Case "[SENDFILE_N]"
        MsgBox "The other side refuses to accept the file.", vbInformation
        AbortFile = False
        
    'Other side wants to draw something
    Case "[DRAW]"
        Dim DrawPic As String
        Dim a
        
        
        On Error Resume Next
        
        
        'Get rid of the command "[DRAW]"
        DrawPic = strArgument
        
        Dim drawtheline
        drawtheline = Split(DrawPic, ",") 'split the string in to each section
    
        For a = 0 To (UBound(drawtheline) - 1) 'for each seperation in the whole string
   
            Dim drawit
            Dim Size
                drawit = Split(drawtheline(a), "$") 'split it in to little sections to decipher
                frmWhiteboard.p1.Line (drawit(0), drawit(1))-(drawit(2), drawit(3)), drawit(4)  'this is the format of drawing the line (you should recognise it from the top)
                Size = drawit(5) 'this alters the size on the other computer
                frmWhiteboard.p1.DrawWidth = Size
        Next a
        
    'Clear the whiteboard
    Case "[WHITEBOARD_CLS]"
    
        frmWhiteboard.p1.Cls
    
    'Change the whiteboard's bg color
    Case "[WHITEBOARD_BACKCOLOR]"
        frmWhiteboard.p1.BackColor = strArgument
    
    'Other side's closed the whiteboard, close this one too
    Case "[CLOSE_WHITEBOARD]"
        'Closes the form
'       frmWhiteboard.timerUnloader.Enabled = True

        
    'Display a popup msg(bit like MSN Messenger)
    Case "[POPUP_MESSAGE]"
        frmPopup.lblMsgDisplay.Caption = strArgument
        frmPopup.Show
        
        If frmConfig.chkMsgAlert.Value = 1 Then
            If Len(Dir$(App.Path & "/Beep.wav")) <> 0 Then
                RetVal = sndPlaySound(frmConfig.txtAlertSound.Text, SND_ASYNC)
            End If
        End If
            
    'Request other side's info
    Case "[GET_PROFILE]"
        sckConnect.SendData "[PROFILE_ARRIVE] " & frmConfig.txtNick.Text & "," & frmConfig.txtEmail.Text & "," & frmConfig.txtIP.Text
    
    'Info arrived, display them
    Case "[PROFILE_ARRIVE]"
        Dim strArrayData
        
        strArrayData = Split(strArgument, ",")
        
        frmProfile.txtNick.Text = strArrayData(0)
        frmProfile.txtEmail.Text = strArrayData(1)
        frmProfile.txtRemoteIP.Text = strArrayData(2)
        frmProfile.lblTitle.Caption = strArrayData(0) & " 's Profile"
        
        frmProfile.Show
        
         
    'This is the actual data of a file from the 'file transfer' feature
    Case Else
      'if this is data from the actual file transfer then
      'add it to the variable that contains the data already sent.
      RecievedFile = RecievedFile & strData
        If FileSize <> 0 Or BeginTransfer <> Timer Then
            frmFileTransfer.lblTransferStatus.Caption = "Received " & Len(RecievedFile) & " bytes (" & Format((Len(RecievedFile) * 100) / FileSize, "00.0") & "%) - " & Format(Len(RecievedFile) / (Timer - BeginTransfer) / 1000, "0.0") & " kbps"
        End If
       frmFileTransfer.ProgressBar1.Value = Len(RecievedFile)
       frmFileTransfer.ProgressBar1.Refresh
       frmFileTransfer.lblTransferStatus.Refresh
   
      'check if the file transfer is complete
      If Len(RecievedFile) = FileSize Then
         frmFileTransfer.cd4.FileName = FileName
         frmFileTransfer.cd4.ShowSave
         'prompt the user for a path to save the file in
         Open frmFileTransfer.cd4.FileName For Binary As #1
         Put #1, 1, RecievedFile
         Close
      End If
      DoEvents
    Exit Sub
                
End Select

End Sub


Private Sub chkBold_Click()

   'toggle bold
   rtbText.SelBold = Not rtbText.SelBold
   rtbText.SetFocus
   
End Sub

Private Sub chkItalic_Click()

   'toggle italic
   rtbText.SelItalic = Not rtbText.SelItalic
   rtbText.SetFocus
   
End Sub

Private Sub chkUnderline_Click()

   'toggle underline
   rtbText.SelUnderline = Not rtbText.SelUnderline
   rtbText.SetFocus
   
End Sub

Private Sub cmbFonts_Click()
  
  On Error Resume Next

  'set the font
  rtbText.SelFontName = cmbFonts.List(cmbFonts.ListIndex)
  rtbText.SetFocus

End Sub

Private Sub cmdColors_Click()

On Error Resume Next

'Code below changes the forecolor of the main Chat Window
cd1.CancelError = True
cd1.ShowColor
rtbText.SelColor = cd1.Color
rtbText.SetFocus

End Sub

Private Sub cmdConnect_Click()

'If cmdConnect.caption = "connect", then it means that
'we'not currently not connected. Therefore, connect to
'the remote host
If cmdConnect.Caption = "Connect" Then
    If txtIP.Text = "" Then
        MsgBox "Please enter a valid IP Address before proceeding.", vbCritical
        Exit Sub
    Else
        optServerClient(0).Enabled = False
        optServerClient(1).Enabled = False
        frmConfig.txtNick.Enabled = False
        txtIP.Enabled = False
        cmdConnect.Caption = "Disconnect"
    
        sckConnect.RemotePort = 1300
    
        sckConnect.RemoteHost = txtIP.Text
    End If
    
    sckConnect.Connect txtIP.Text, 1300
    
    If sckConnect.State = 6 Then
        AddSysMessage vbCrLf & "*** Connecting..."
    End If
    Exit Sub
'If we ar the server, 'Listen' for incoming connecion requests
ElseIf cmdConnect.Caption = "Listen" Then
    optServerClient(0).Enabled = False
    optServerClient(1).Enabled = False
    optServerClient(0).Value = True
    cmdConnect.Caption = "Disconnect"
    sckConnect.LocalPort = 1300
    sckConnect.Listen
    AddSysMessage vbCrLf & "*** Waiting for connection..."
    Exit Sub
'We want to disconnect, close winsock
ElseIf cmdConnect.Caption = "Disconnect" Then
   
    If sckConnect.State = sckConnected Then
        sckConnect.SendData "[LEAVE] " & frmConfig.txtNick.Text
    End If

    
    'Tell the user that he/she is disconnected
    AddSysMessage (vbCrLf & "*** Disconnected")
    sckConnect.Close
    optServerClient(0).Enabled = True
    optServerClient(1).Enabled = True
    
    'Get the captions right
    If optServerClient(0).Value = True Then
        cmdConnect.Caption = "Listen"
    Else
        cmdConnect.Caption = "Connect"
    End If
    txtIP.Enabled = True
    Exit Sub
End If

    
End Sub
Sub AddSysMessage(strText As String, Optional lngColor As Long = vbRed)

  'A system message is something like '*** Disconnected'

  rtbChat.SelStart = Len(rtbChat.Text)
  rtbChat.SelLength = 0
  rtbChat.SelText = strText$
  rtbChat.SelStart = Len(rtbChat.Text) - Len(strText$)
  rtbChat.SelLength = Len(strText$)
  rtbChat.SelColor = lngColor&
  rtbChat.SelBold = False
  rtbChat.SelFontName = "Courier New"
  rtbChat.SelFontSize = 10
  rtbChat.SelStart = Len(rtbChat.Text)
  rtbChat.SelLength = 0


End Sub

Private Sub cmdConnect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Looks good
cmdConnect.BackColor = RGB(32, 32, 32)
cmdConnect.ForeColor = vbWhite
End Sub

Private Sub cmdConnect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Looks good
cmdConnect.BackColor = &HE0E0E0
cmdConnect.ForeColor = &H400000
End Sub


Private Sub Form_paint()

'If Gradient Background option is chosen at frmConfig,
'then draw a random gradient background
If GradientBG = True Then
    Randomize

    Dim RandomInt
    RandomInt = Rand(1, 7)
    Select Case RandomInt
     Case 1
            FadeFormPurple Me
        Case 2
            FadeFormRed Me
        Case 3
            FadeFormBlue Me
        Case 4
            FadeFormYellow Me
        Case 5
            FadeFormGreen Me
        Case 6
            FadeFormGrey Me
        Case 7
            FadeFormGrey Me
    End Select
Else
    If frmConfig.optMainWinBG(1).Value <> True Then frmMain.BackColor = &H400000
    
End If
    
End Sub



Private Sub imgClose_Click()
'Just a small sub showing a small animation just before
'closing the app
Startrek Me
'Wait for sometime
Sleep 300
'Quit
Unload Me
End
End Sub

Private Sub imgMin_Click()
'Minimize the form
frmMain.Visible = False

End Sub

Private Sub itmAbout_Click()
'Show frmAbout
frmAbout.Show

End Sub

Private Sub itmClear_Click()
'Clear the chat window
rtbChat.Text = ""

End Sub

Private Sub itmConnect_Click()
'Connect to the other side
Call cmdConnect_Click

End Sub

Private Sub itmCopy_Click()
'Copy
EditCopy

End Sub

Private Sub itmCut_Click()
'Cut
EditCut

End Sub

Private Sub itmDisconnect_Click()
'Refer back to cmdConnect_Click
Call cmdConnect_Click

End Sub

Private Sub itmExit_Click()

'Take the icon out of the System Tray
RemoveIconFromTray

'Perform cool animation
Startrek Me

'Quit
Unload Me

End Sub

Private Sub itmFileTransfer_Click()
'Show the file transfer form
frmFileTransfer.Show

End Sub

Private Sub itmGetProfile_Click()
'If we ar connected, tell the other side that we want
'his/her info
If sckConnect.State = sckConnected Then
    sckConnect.SendData "[GET_PROFILE] "
Else
    'Not connected - error msg - inform the user
    MsgBox "You are not connected to the other side.", vbInformation
    frmProfile.Show
End If


End Sub

Private Sub itmListen_Click()
'Refer cmdConnect_Click
Call cmdConnect_Click

End Sub

Private Sub itmMinimize_Click()
'Hide the form
frmMain.Visible = False

End Sub

Private Sub itmOpen_Click()
'Open a text file and input it into the chat window

Dim Buffer
cd1.Filter = "Text Files (*.txt) | *.txt"
cd1.ShowOpen
Open cd1.FileName For Input As #1
While Not EOF(1)
    Line Input #1, Buffer
Wend
Close #1
rtbChat.Text = rtbChat.Text & Buffer
End Sub

Private Sub itmOptions_Click()
'Show frmConfig
frmConfig.Show
End Sub

Private Sub itmPaste_Click()
'Paste
EditPaste

End Sub

Private Sub itmPrint_Click()
'Print the contents of the RichTextBox with a one inch margin
PrintRTF rtbChat, 1440, 1440, 1440, 1440 ' 1440 Twips = 1 Inch

End Sub

Private Sub itmQuit_Click()
'Remove icon from System Tray
RemoveIconFromTray
Unload Me
End
End Sub

Private Sub itmRestore_Click()
'Make frmMain visible again
If frmMain.Visible = False Then frmMain.Visible = True

End Sub

Private Sub itmSaveAs_Click()
'Save the conversation in a text file
Dim Buffer
cd1.ShowSave
Open cd1.FileName & ".txt" For Output As #2
Buffer = rtbChat.Text
Print #2, Buffer
Close


End Sub

Private Sub itmSendFile_Click()

'Send a file
frmFileTransfer.Show

End Sub

Private Sub itmSendPopup_Click()
'Send popup msg
Call itmSendPopupMessage_Click

End Sub

Private Sub itmSendPopupMessage_Click()
'Send popup msg
Dim Message
Message = InputBox("Enter the message you wish to send : ", "Instant Popup Message")
If Message = "" Then
    Exit Sub
Else
    sckConnect.SendData "[POPUP_MESSAGE] " & frmConfig.txtNick.Text & " Says: " & vbCrLf & Message
End If


End Sub

Private Sub itmViewHelpFile_Click()
Call ShellExecute(hWnd, "Open", "ReadMe.txt", "", App.Path, 3)



End Sub

Private Sub itmwhiteboard_Click()
'Show the whiteboard
frmWhiteboard.Show
'If sckConnect.State = sckConnected Then
'    sckConnect.SendData "[OPEN_WHITEBOARD] "
'End If

End Sub

Private Sub lblSend_Click()
'Send a msg

If sckConnect.State = sckConnected Then
    sckConnect.SendData "[MESSAGE] " & frmConfig.txtNick.Text & ":" & rtbText.Text
End If
AddChat frmConfig.txtNick.Text, rtbText.TextRTF
rtbText.Text = ""
End Sub


Private Sub Form_Load()



'Determine the bg color of chat window by looking at Settings.ini
If frmConfig.txtChatWinColor.Text <> "" Then
    rtbChat.BackColor = frmConfig.txtChatWinColor.Text
Else
    rtbChat.BackColor = &H80000005
End If

'Put icon in system tray
picSysTrayIcon.Visible = False

'Disable the cross(close button) on the form
EnableCloseButton Me.hWnd, False


Hook Me.hWnd   ' Set up our event handler
AddIconToTray Me.hWnd, frmMain.picSysTrayIcon.Picture, Me.Icon.Handle, "Access many program functions here."

'Hide that ugly mouth
TTS.Visible = False

'The following command line....no comment
rtbChat.DataChanged = False

'Load different fonts
Dim intBuffer As Integer, strFont As String
 
   'load printer fonts to combobox
   If Dir$(App.Path & "\fonts.dat") = "" Then
        'font file doesnt exist. Create it.
        Open App.Path & "\fonts.dat" For Output As #1
             For intBuffer% = 0 To Printer.FontCount - 1
                Call cmbFonts.AddItem(Printer.Fonts(intBuffer%))
                Print #1, Printer.Fonts(intBuffer%)
             Next intBuffer%
        Close #1
   Else
        'load fonts from file
        Open App.Path & "\fonts.dat" For Input As #1
             While Not EOF(1)
                Input #1, strFont$
                Call cmbFonts.AddItem(strFont$)
             Wend
        Close #1
   End If
   
 cmbFonts.ListIndex = 0
 ''cmbFonts.Sorted = True 'Alphabetize list


  'set combobox to "Arial"
  For intBuffer% = 0 To cmbFonts.ListCount - 1
    If cmbFonts.List(intBuffer%) = "Arial" Then cmbFonts.ListIndex = intBuffer%: Exit For
  Next intBuffer%
  

  'set rtbText's font-styles
  rtbText.SelBold = False
  rtbText.SelUnderline = False
  rtbText.SelItalic = False
  rtbText.SelColor = vbBlack
  rtbText.SelFontName = cmbFonts.List(cmbFonts.ListIndex)
  rtbText.SelFontSize = 10
  
'Make 'client' the default option
optServerClient(1).Value = True

'Get the captions of the button(s) right
If optServerClient(1).Value = True Then
    cmdConnect.Caption = "Connect"
    optServerClient(0).Value = False
Else
    cmdConnect.Caption = "Listen"
End If

'Test to see if theres a Settings.ini file, if not
'prompt the user to create one, very important!
If Len(Dir$(App.Path & "/Settings.ini")) <= 0 Then
    frmConfig.Show
    MsgBox "Please fill in all your details before proceeding.", vbInformation
    Ontop frmConfig
    Exit Sub
End If

'Lines below determine the form's bg by looking at the
'Settings.ini file
iniPath$ = App.Path & "/Settings.ini"

On Error GoTo Err:

If GetFromINI("DIGIChat", "optMainWinBG(0)", iniPath$) = True Then
    imgBg.Visible = True
    imgCustomBG.Visible = False
    Me.BackColor = &H400000
    GradientBG = False
ElseIf GetFromINI("DIGIChat", "optMainWinBG(1)", iniPath$) = True Then
    GradientBG = True
    imgBg.Visible = False
    imgCustomBG.Visible = False
ElseIf GetFromINI("DIGIChat", "optMainWinBG(0)", iniPath$) = False And GetFromINI("DIGIChat", "optMainWinBG(1)", iniPath$) = False Then
    imgBg.Visible = False
    imgCustomBG.Picture = LoadPicture(GetFromINI("DIGIChat", "CustomPictureFile", iniPath$))
    imgCustomBG.Visible = True
    GradientBG = False
End If
    
'Just another error handler
Err:
    imgBg.Visible = True
    frmConfig.optMainWinBG(0).Value = True
    

End Sub
Private Sub Form_Unload(Cancel As Integer)

'If connected tell other side that u left the chat
If sckConnect.State = sckConnected Then sckConnect.SendData "[LEAVE] " & frmConfig.txtNick.Text
sckConnect.Close
sckConnect.LocalPort = 0
   
'Remove icon from system tray
RemoveIconFromTray

'Kick away the event handler
Unhook


End Sub
Private Sub lblSend_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Looks nice...
lblSend.BackColor = RGB(32, 32, 32)
lblSend.ForeColor = vbWhite
End Sub
Private Sub lblSend_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Looks nice...
lblSend.BackColor = &HE0E0E0
lblSend.ForeColor = &H400000
End Sub

Private Sub sckConnect_Close()

   'if user is guest then display "Disconnected".
   If optServerClient(1).Value = True Then
      AddSysMessage (vbCrLf & "*** Disconnected")
   End If
   
optServerClient(0).Enabled = True
optServerClient(1).Enabled = True
txtIP.Enabled = True

'Get the captions right
If optServerClient(0).Value = True Then
    cmdConnect.Caption = "Listen"
Else
    cmdConnect.Caption = "Connect"
End If



End Sub

Private Sub sckConnect_Connect()

'Tell user that he/she is currently connected to other side
AddSysMessage (vbCrLf & "*** Connected")

'Tell other side that u joined the chat
If optServerClient(1).Value = True Then
    sckConnect.SendData "[JOIN] " & frmConfig.txtNick.Text
End If

'Ready to type
rtbText.SetFocus


End Sub


  
Private Sub optServerClient_Click(Index As Integer)

'Decide whether u ar a client or a server

    Select Case Index
       Case 0: 'Server
        
          cmdConnect.Caption = "Listen"
          txtIP.Visible = False
          lblIP.Visible = False
          
       Case 1: 'Client
          txtIP.Visible = True
          lblIP.Visible = True
          txtIP.Text = ""
          cmdConnect.Caption = "Connect"
          
    End Select

End Sub


Private Sub rtbText_KeyPress(KeyAscii As Integer)
'If the key 'Enter' is pressed, send the msg
   
   If KeyAscii = 13 Then 'If user pressed 'Enter'
      lblSend_Click 'click 'Send' button
      KeyAscii = 0 'Make sure it doesnt write enter to rtbText
   End If
    
End Sub
Private Sub sckConnect_ConnectionRequest(ByVal requestID As Long)
'Accept the connection after closing any previous connection
If sckConnect.State <> sckClosed Then
    sckConnect.Close
    DoEvents
End If

sckConnect.Accept requestID

End Sub

Private Sub sckConnect_DataArrival(ByVal bytesTotal As Long)

Dim strData

'Assign the data to variable strData
sckConnect.GetData strData, vbString

'Leave the work to sub ParseData
ParseData (strData)

End Sub

Private Sub sckConnect_SendComplete()
'File transfer is completed
If SendingFile Then
      frmFileTransfer.lblTransferStatus.Caption = "Transfer Complete"
      SendingFile = False
End If

End Sub

Private Sub sckConnect_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
'File sending in progress.....

   If SendingFile Then
        Dim BytesAlreadySent As Long
        BytesAlreadySent = Val(Mid(frmFileTransfer.lblTransferStatus.Caption, 5, InStr(6, frmFileTransfer.lblTransferStatus.Caption, "b")))
        BytesAlreadySent = BytesAlreadySent + bytesSent
        frmFileTransfer.ProgressBar1.Value = BytesAlreadySent
        If FileSize <> 0 Or BeginTransfer <> Timer Then
            frmFileTransfer.lblTransferStatus.Caption = "Sent " & BytesAlreadySent & " bytes (" & Format((BytesAlreadySent * 100) / (BytesAlreadySent + bytesRemaining), "00.0") & "%) - " & Format(BytesAlreadySent / (Timer - BeginTransfer) / 1000, "0.0") & " kbps"
        End If
        frmFileTransfer.ProgressBar1.Refresh
        frmFileTransfer.lblTransferStatus.Refresh
   End If
   
End Sub

Public Sub SysTrayMouseEventHandler()
'Code below displays a popup menu when a right mouse button
'click is detected on the icon in the SysTray

If optServerClient(0).Value = True Then
    If sckConnect.State = sckConnected Then
        itmConnect.Enabled = False
        itmListen.Enabled = False
        itmDisconnect.Enabled = True
    Else
        itmConnect.Enabled = False
        itmListen.Enabled = True
        itmDisconnect.Enabled = False
    End If
Else
    If sckConnect.State = sckConnected Then
        itmConnect.Enabled = False
        itmListen.Enabled = False
        itmDisconnect.Enabled = True
    Else
        itmConnect.Enabled = True
        itmListen.Enabled = False
        itmDisconnect.Enabled = False
    End If
End If


SetForegroundWindow Me.hWnd
PopupMenu mnuPopup, vbPopupMenuRightButton

End Sub

Sub Startrek(frm As Form)

'Small animation when the form is closed

Dim GotoVal
Dim GoInto


    GotoVal = frm.Height / 2
    For GoInto = 1 To GotoVal
        DoEvents
        frm.Height = frm.Height - 100
        frm.Top = (Screen.Height - frm.Height) \ 2
        If frm.Height <= 500 Then Exit For
    Next GoInto
horiz:
    frm.Height = 30
    GotoVal = frm.Width / 2
    For GoInto = 1 To GotoVal
        DoEvents
        frm.Width = frm.Width - 100
        frm.Left = (Screen.Width - frm.Width) \ 2
        If frm.Width <= 2000 Then Exit For
    Next GoInto

End Sub





' Standard Copy procedure
' This could be enhanced to support alternative clipboard types
Public Sub EditCopy()



' Set the error handler
On Error Resume Next

' Clear the clipboard
Clipboard.Clear
If Err.Number Then
   Err.Raise iCB_CLEAR_ERR, "CApplication", "Could not clear clipboard."
Else
   ' Place selected text on the clipboard
   Clipboard.SetText Screen.ActiveControl.SelText
   If Err.Number Then
    Err.Raise iCB_SET_ERR, "CApplication", "Could not set text from active control to clipboard."
   End If
End If
End Sub


' Standard Cut procedure
' This could be enhanced to support alternative clipboard types
Public Sub EditCut()
' Set the error handler
On Error Resume Next

' Clear the clipboard
Clipboard.Clear
If Err.Number Then
   Err.Raise iCB_CLEAR_ERR, "CApplication", "Could not clear clipboard."
Else
   ' Place selected text on the clipboard
   Clipboard.SetText Screen.ActiveControl.SelText
   Screen.ActiveControl.SelText = ""
   If Err.Number Then
    Err.Raise iCB_SET_ERR, "CApplication", "Could not set text from active control to clipboard."
   End If
End If

End Sub





' Standard Paste procedure
' This could be enhanced to support alternative clipboard types
Public Sub EditPaste()
' Set the error handler
On Error Resume Next

' Place the text from the clipboard
Screen.ActiveControl.SelText = Clipboard.GetText()
If Err.Number Then
   Err.Raise iCB_PASTE_ERR, "CApplication", "Could not paste text from clipboard to active control."
End If

End Sub


VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFileTransfer 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "ShapedForm"
   ClientHeight    =   4770
   ClientLeft      =   1500
   ClientTop       =   2115
   ClientWidth     =   8985
   ControlBox      =   0   'False
   ForeColor       =   &H80000008&
   MaxButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   4800
      TabIndex        =   10
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtRemoteIP 
      Height          =   315
      Left            =   4680
      TabIndex        =   6
      Top             =   1200
      Width           =   2505
   End
   Begin VB.TextBox txtLocalIP 
      BackColor       =   &H8000000A&
      Height          =   315
      Left            =   4680
      TabIndex        =   5
      Top             =   1560
      Width           =   2505
   End
   Begin MSComDlg.CommonDialog cd4 
      Left            =   1320
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7320
      TabIndex        =   3
      Top             =   2040
      Width           =   585
   End
   Begin VB.TextBox txtFileToSend 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2040
      Width           =   5055
   End
   Begin VB.Frame fraSending 
      BackColor       =   &H00E0E0E0&
      Height          =   1335
      Left            =   1320
      TabIndex        =   0
      Top             =   2640
      Width           =   7335
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   375
         Left            =   840
         TabIndex        =   11
         Top             =   720
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Remote IP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   2160
      TabIndex        =   8
      Top             =   1200
      Width           =   2445
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Local IP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   2160
      TabIndex        =   7
      Top             =   1560
      Width           =   2445
   End
   Begin VB.Label lblTransferStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   2160
      TabIndex        =   4
      Top             =   720
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "File to send :"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Image imgClose 
      Height          =   255
      Left            =   7920
      Picture         =   "frmFileTransfer.frx":0000
      Top             =   840
      Width           =   255
   End
End
Attribute VB_Name = "frmFileTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreatePolygonRgn Lib "GDI32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "GDI32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "GDI32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "GDI32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "GDI32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "GDI32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "USER32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "USER32" () As Long
Private Type POINTAPI
   X As Long
   Y As Long
End Type
Private Const RGN_COPY = 5
Private Const CreatedBy = "VBSFC 6.2"
Private Const RegisteredTo = "Not Registered"
Private ResultRegion As Long

Private Function CreateFormRegion(ScaleX As Single, ScaleY As Single, OffsetX As Integer, OffsetY As Integer) As Long
    Dim HolderRegion As Long, ObjectRegion As Long, nRet As Long, Counter As Integer
    Dim PolyPoints() As POINTAPI
    ResultRegion = CreateRectRgn(0, 0, 0, 0)
    HolderRegion = CreateRectRgn(0, 0, 0, 0)

'This procedure was generated by VB Shaped Form Creator.  This copy has
'NOT been registered for commercial use.  It may only be used for non-
'profit making programs.  If you intend to sell your program, I think
'it's only fair you pay for mine.  Commercial registration costs $30,
'and can be performed online.  See "Registration" item on the help menu
'for details.

'Latest versions of VB Shaped Form Creator can be found at my website at
'http://www.comports.com/AlexV/VBSFC.html or you can visit my main site
'with many other free programs and utilities at http://www.comports.com/AlexV

'Lines starting with '! are required for reading the form shape using the
'Import Form command in VB Shaped Form Creator, but are not necessary for
'Visual Basic to display the form correctly.

'!Shaped Form Region Definition
'!2,51,59,599,318,77,133,1
    ObjectRegion = CreateRoundRectRgn(59 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 51 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 599 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 318 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 154 * ScaleX * 15 / Screen.TwipsPerPixelX, 266 * ScaleY * 15 / Screen.TwipsPerPixelY)
    nRet = CombineRgn(ResultRegion, ObjectRegion, ObjectRegion, RGN_COPY)
    DeleteObject ObjectRegion
    DeleteObject HolderRegion
    CreateFormRegion = ResultRegion
End Function

Private Sub cmdQuit_Click()
'Confirm quit
Dim Answer As VbMsgBoxResult
   Answer = MsgBox("Do you really want to quit?", vbInformation Or vbYesNo)
   If Answer = vbYes Then
      Unload Me
   End If
End Sub

Private Sub cmdSend_Click()
'Prepare to send the file
SendFile txtFileToSend.Text

End Sub

Private Sub Command1_Click()
'if not connected then tell the user
If frmMain.sckConnect.State <> sckConnected Then
    MsgBox "You must be connected to send a file", vbCritical
    Exit Sub
End If

'Let user choose which file to send
On Error Resume Next
cd4.Filter = "All Files | *.*"
cd4.ShowOpen
txtFileToSend = cd4.FileName


End Sub

Private Sub Form_Load()
    Dim nRet As Long
    nRet = SetWindowRgn(Me.hWnd, CreateFormRegion(1, 1, 0, 0), True)
    'If the above two lines are modified or moved a second copy of
    'them may be added again if the form is later Modified by VBSFC.
    Ontop Me
    If frmMain.sckConnect.State = sckConnected Then
        lblTransferStatus.Caption = "Status : Connected"
    Else
        lblTransferStatus.Caption = "Status : Not connected"
    End If
txtLocalIP.Text = frmMain.sckConnect.LocalIP
If frmMain.sckConnect.State = sckConnected Then
    txtRemoteIP.Text = frmMain.sckConnect.RemoteHostIP
End If

    
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Next two lines enable window drag from anywhere on form.  Remove them
'to allow window drag from title bar only.
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub
Private Sub Form_Unload(Cancel As Integer)
    DeleteObject ResultRegion
    'If the above line is modified or moved a second copy of it
    'may be added again if the form is later Modified by VBSFC.
End Sub
Private Sub imgClose_Click()

Unload Me

End Sub

'procedure   :SendFile
'inptuts     :The full path & filename of the file to send
'what it does:
'              a)It reads the file into a string variable
'              b)It sends the file information (filename and size)
'                to the other side and it waits for a response.
'              c)If the response is a yes,it sends the file
Public Sub SendFile(ByVal FileName As String)

   Dim FileData As String
   Dim ByteData As Byte
   Dim Counter As Long
   
   If frmMain.sckConnect.State = sckConnected Then
   Open FileName For Binary As #1
   
   ProgressBar1.Max = LOF(1)
   ProgressBar1.Value = 0
   
   lblTransferStatus.Caption = "Reading file into memmory...Please be patient..."
   
   'Read the file into the variable FileData
   FileData = Input(LOF(1), 1)
   
   DoEvents
   
   lblTransferStatus.Caption = "Initiating file transfer..."
   
   Close #1
   
   SendingFile = False
   AbortFile = False
   
   If MsgBox(FileTitle(FileName) & " (" & Len(FileData) & " bytes)" & vbCrLf & _
   "Begin the file transfer?", vbInformation Or vbYesNo) <> vbYes Then
      Exit Sub
   End If
   
    frmMain.sckConnect.SendData "[SENDFILE] " & Len(FileData) & "_" & FileTitle(FileName)
        
   
   'This loop suspends the program until the other side
   Do Until SendingFile Or AbortFile Or DoEvents = 0
      DoEvents
   Loop
   
   lblTransferStatus.Caption = "Sent 0 bytes (0%)"
   
   'This command begins the file transfer.The whole file is stored
   'in the string variable FileData.
   BeginTransfer = Timer
   frmMain.sckConnect.SendData FileData
   Else
   MsgBox "You must be connected in order to begin the transfer.", vbCritical
   End If

End Sub


'function  : FileTitle
'inputs    : A string containing a full filename (path & file title)
'returns   : The file title
'example   : FileTitle("c:\windows\desktop\readme.txt")
'            returns "readme.txt"
Private Function FileTitle(ByVal FileName As String) As String

   Dim i As Integer
   Dim Temp As String
   
   'if the string includes a path
   If InStr(FileName, "\") <> 0 Then
      'then begin the proccess of parsing the file title.
      i = Len(FileName)
      Do Until Left(Temp, 1) = "\"
         i = i - 1
         Temp = Mid(FileName, i)
      Loop
      FileTitle = Mid(Temp, 2)
   Else
      'If it's already a file title,just return the same string.
      FileTitle = FileName
   End If

End Function




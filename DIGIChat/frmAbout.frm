VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "About MyApp"
   ClientHeight    =   3375
   ClientLeft      =   2295
   ClientTop       =   1605
   ClientWidth     =   5820
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2329.484
   ScaleMode       =   0  'User
   ScaleWidth      =   5465.281
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   480
      Left            =   5160
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   2160
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   600
      Top             =   2880
   End
   Begin VB.PictureBox p1 
      BackColor       =   &H8000000C&
      ForeColor       =   &H8000000E&
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2355
      ScaleWidth      =   5760
      TabIndex        =   3
      Top             =   360
      Width           =   5814
      Begin VB.PictureBox p2 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   3615
         Left            =   0
         ScaleHeight     =   3615
         ScaleWidth      =   5685
         TabIndex        =   4
         Top             =   2280
         Width           =   5685
         Begin VB.Label lblAuthorsEmail 
            BackStyle       =   0  'Transparent
            Caption         =   "danielyh@optushome.com.au"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   2640
            MousePointer    =   10  'Up Arrow
            TabIndex        =   9
            Top             =   3240
            Width           =   2535
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Questions?        : Email "
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   8
            Top             =   3240
            Width           =   2415
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmAbout.frx":08CA
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   2535
            Index           =   2
            Left            =   120
            TabIndex        =   7
            Top             =   840
            Width           =   5055
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Written on       : 06/10/2001"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   4095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Written by       : Daniel Ho"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   4095
         End
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1560
      TabIndex        =   0
      Top             =   3000
      Width           =   1260
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&System Info..."
      Height          =   345
      Left            =   2880
      TabIndex        =   2
      Top             =   3000
      Width           =   1245
   End
   Begin VB.Image imgClose 
      Height          =   225
      Left            =   5520
      Picture         =   "frmAbout.frx":0A8E
      Top             =   101
      Width           =   225
   End
   Begin VB.Image imgMinimize 
      Height          =   225
      Left            =   5280
      Picture         =   "frmAbout.frx":0EBE
      Top             =   101
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   0
      Picture         =   "frmAbout.frx":12C2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5820
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1936.06
      Y2              =   1936.06
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1946.413
      Y2              =   1946.413
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const conSwNormal = 1

Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub



Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Try To Get System Info Program Path Only From Registry...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Validate Existance Of Known 32 Bit File Version
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - File Can Not Be Found...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Registry Entry Can Not Be Found...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Private Sub Form_Load()

'Start scrolling the Credits...upwards...
p2.Top = 2280
Timer1.Enabled = True


'Make bottom part of form transparent, COOL!!!
Call MakeTranslucent(Me, tColor)



End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Refresh the transparency, or enable the user to drag
'the form
If Button = vbLeftButton Then
    Call DragForm(Me)
    Call MakeTranslucent(Me, tColor)
End If

'Enable timer in case it's stopped by placing cursor over
'my email address
Timer1.Enabled = True
End Sub

Private Sub Form_Resize()
'Refresh form's transparency after being resized
Call MakeTranslucent(Me, tColor)
End Sub
Private Sub Form_Unload(Cancel As Integer)
'Disable the timer just to make sure everything works OK
Timer1.Enabled = False

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Allow user to drag form by title bar
If Button = vbLeftButton Then
    Call Form_MouseMove(vbLeftButton, Shift, X, Y)
End If
End Sub

Private Sub imgClose_Click()
'Close the form
Unload Me

End Sub

Private Sub lblAuthorsEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Stop scrolling when cursor is placed on top of my email
lblAuthorsEmail.ForeColor = &H80&
Timer1.Enabled = False

End Sub
Private Sub lblAuthorsEmail_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Call default mail program to send mail to me!
ShellExecute hWnd, "open", "mailto:danielyh@optushome.com.au", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub p2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Restore things back after cursor've moved away from my
'email
lblAuthorsEmail.ForeColor = &H800000
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    'Make p2 move upwards by changing its Top property
    'consistently
    p2.Top = p2.Top - 10
    'If p2 is off the screen, reset its position to bottom
    'and continuing the scrolling cycle
    If p2.Top <= -3500 Then
        p2.Top = 2270
    End If
    
End Sub

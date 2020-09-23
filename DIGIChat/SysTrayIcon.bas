Attribute VB_Name = "SysTrayIcon"
Option Explicit

Public Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, IpData As NOTIFYICONDATA) As Long

Public Const ICON_MESSAGE = 1
Public Const ICON_ICON = 2
Public Const ICON_TIP = 4
Public Const ADD_ICON = 0
Public Const MODIFY_ICON = 1
Public Const DELETE_ICON = 2
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONDOWN = &H204

Type NOTIFYICONDATA     'Information structure for the system tray
     cbSize As Long     'icon.
     hWnd As Long
     uID As Long
     uFlags As Long
     uCallbackMessage As Long
     hIcon As Long
     szTip As String * 64
End Type

Public IconData As NOTIFYICONDATA     'Stores the information about the TNA Icon

Public Sub CreateSysIcon()
''-------------------------------------------------------''
'' Sets variables for Taskbar Notification Area icon and ''
'' displays it.                                          ''
''-------------------------------------------------------''
Dim result As Long
IconData.cbSize = Len(IconData)
IconData.hWnd = frmMain.hWnd     'The Form You Want To Handle The Clicks
IconData.uID = vbNull
IconData.uFlags = ICON_MESSAGE Or ICON_ICON Or ICON_TIP
IconData.hIcon = frmMain.picSysTrayIcon.Picture   'The Name Of The Picture Control Containing The Icon
IconData.szTip = "Place The Information You Want To Appear In The Tool Tip"
IconData.uCallbackMessage = WM_RBUTTONDOWN
result = Shell_NotifyIcon(DELETE_ICON, IconData)
result = Shell_NotifyIcon(ADD_ICON, IconData)
End Sub



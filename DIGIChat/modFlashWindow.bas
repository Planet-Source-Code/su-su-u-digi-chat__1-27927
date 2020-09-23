Attribute VB_Name = "modFlashWindow"
Option Explicit

'Flash the window's caption
Public Declare Function FlashWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)



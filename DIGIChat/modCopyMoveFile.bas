Attribute VB_Name = "modCopyMoveFile"
Option Explicit

'Used for copying or moving a file to a destination
Public Declare Function CopyFile Lib "kernel32" Alias _
"CopyFileA" (ByVal lpExistingFileName As String, ByVal _
lpNewFileName As String, ByVal bFailIfExists As Long) As Long

Public Declare Function MoveFile Lib "kernel32" Alias _
"MoveFileA" (ByVal lpExistingFileName As String, ByVal _
lpNewFileName As String) As Long




Attribute VB_Name = "modFileTransfer"
Option Explicit

'This boolean indicates if this side if currently sending a file
Global SendingFile As Boolean
'This boolean indicates if the other side denies to recieve the file
Global AbortFile As Boolean
'This variable contains the file that is being sent from the other side
Global RecievedFile As String
'This variable is used in keeping track of the transfer rate.
Global BeginTransfer As Single

Global FileName As String
Global FileSize As Long

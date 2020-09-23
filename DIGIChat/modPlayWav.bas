Attribute VB_Name = "modPlayWav"
Option Explicit

'Easily play a .wav file
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Const SND_ASYNC = &H1



Attribute VB_Name = "modRandomNumber"
Option Explicit

'===============================================
'Rand - Return a random number in a given range.
'
'Parameters:
'  Low  - The lower bounds of the range.
'  High - The upper bounds of the range.
'
'Returns:
'  Returns a random number from Low..High.
'===============================================

Public Function Rand(ByVal Low As Long, _
                     ByVal High As Long) As Long
  Rand = Int((High - Low + 1) * Rnd) + Low
End Function





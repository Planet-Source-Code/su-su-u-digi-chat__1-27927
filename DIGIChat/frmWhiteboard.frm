VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmWhiteboard 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Whiteboard"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   5280
   Begin VB.Timer tmrStickWithfrmMain 
      Interval        =   100
      Left            =   2280
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   1320
      Top             =   1560
   End
   Begin VB.Timer timerUnloader 
      Interval        =   50
      Left            =   1800
      Top             =   1560
   End
   Begin VB.Timer timerLoader 
      Interval        =   50
      Left            =   840
      Top             =   1560
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdBackColor 
      Caption         =   "Back color"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrushSize 
      Caption         =   "Brush size"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdBrushColor 
      Caption         =   "Brush Color"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   5040
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cd2 
      Left            =   840
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox p1 
      BackColor       =   &H8000000E&
      Height          =   4695
      Left            =   0
      ScaleHeight     =   4635
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Draw below :"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   4815
   End
   Begin VB.Image imgClose 
      Height          =   255
      Left            =   5040
      Picture         =   "frmWhiteboard.frx":0000
      Top             =   0
      Width           =   255
   End
   Begin VB.Image imgMin 
      Height          =   255
      Left            =   4800
      Picture         =   "frmWhiteboard.frx":03B6
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "frmWhiteboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*******************WhiteboardVariables********************************
Public lastx As Integer 'x co-ordinate for line
Public lasty As Integer 'y co-ordinate for line
Public Draw As String 'string for the draw feature
Public Color 'color of brush
Public Size As Integer 'size of the brush
'*************************************************************

Public Reverse As Boolean 'Used for starting up and unloading animation

Private Sub cmdBrushColor_Click()
cd2.ShowColor
Color = cd2.Color
End Sub


Private Sub cmdBrushSize_Click()
On Error Resume Next

'Let the user enter the brush size (simplest way)

Size = InputBox("Please enter a size", "Brush Size")

If IsNumeric(Size) = False Then
    MsgBox "Please enter a valid Brush Size.", vbInformation, "Invalid Brush Size"
End If



p1.DrawWidth = Size


End Sub

Private Sub cmdBackColor_Click()
'Set the backcolor
cd2.ShowColor
BackColor = cd2.Color
p1.BackColor = BackColor
'Make whiteboard on other side the same BG color
If frmMain.sckConnect.State = sckConnected Then
    frmMain.sckConnect.SendData "[WHITEBOARD_BACKCOLOR] " & BackColor
End If
End Sub

Private Sub cmdClear_Click()
'clear out the whiteboard
p1.Cls
'Tell other side to clear out whiteboard as well
If frmMain.sckConnect.State = sckConnected Then
    frmMain.sckConnect.SendData "[WHITEBOARD_CLS] "
End If

End Sub

Private Sub Form_Load()

'If connected then start sending what is being drawn on the
'board
If frmMain.sckConnect.State = sckConnected Then
    Timer1.Enabled = True
End If

'Disabled (temporarily) the timer used for closing the form
timerUnloader.Enabled = False

'Cool stuff - Make frmWhiteboard stick to frmMain!
tmrStickWithfrmMain.Enabled = True

'Configurations...
With Me
    .Reverse = False
    .Top = frmMain.Top
    .Left = frmMain.Left + frmMain.Width
    .Width = 50
    .Refresh
    .timerLoader.Enabled = True
End With
    
'Disable the close button to make sure the fascinating
''closing' animation must be played to close the form
EnableCloseButton Me.hWnd, False


End Sub


Private Sub Form_Unload(Cancel As Integer)
'Release the timer
tmrStickWithfrmMain.Enabled = False

End Sub

Private Sub imgClose_Click()
'Start cool animation...
timerUnloader.Enabled = True
'If frmMain.sckConnect.State = sckConnected Then
'    frmMain.sckConnect.SendData "[CLOSE_WHITEBOARD] "
'End If

End Sub

Private Sub imgMin_Click()
'Minimize the form
Me.WindowState = 1
End Sub

Private Sub p1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then 'if the user is holding down the left mouse button...
p1.Line (lastx, lasty)-(X, Y), Color 'this draws the line using the two variables we chose at top
If Color = "" Then Color = vbBlack
Draw = Draw & lastx & "$" & lasty & "$" & X & "$" & Y & "$" & Color & "$" & Size & "," 'this is the string that I have named 'draw'. In this string, winsock sends all the data needed for the construction of a line on the other computer, including color
End If

lastx = X 'assigns the variable a value
lasty = Y 'assigns the variable a value




End Sub

Private Sub Timer1_Timer()
If Len(Draw) <> 0 Then 'if the length of the string 'draw' is no 0 then send the data

If Len(Draw) > 3000 Then 'this is the limit that I have made. It can be altered, but note that much higher will cause winsock to crash
Draw = "" 'clear the string so that it doesnt accumulate
Exit Sub 'if the data is too large, don't send it, and exit this sub
End If

If frmMain.sckConnect.State = sckConnected Then
    frmMain.sckConnect.SendData "[DRAW] " & Draw 'sends a draw protocal followed by the string containing all the info
    Draw = "" 'clear the string so that it doesnt accumulate
End If
End If
End Sub

Private Sub timerLoader_Timer()
On Error Resume Next

'Add to the form's width on every set interval
Me.Width = Me.Width + 100
If Me.Width >= 5360 Then
    timerLoader.Enabled = False
End If

End Sub

Private Sub timerUnloader_Timer()
'Take off the form's width on every set interval
Me.Width = Me.Width - 100
If Me.Width <= 1000 Then
     timerUnloader.Enabled = False
    Unload Me
End If

End Sub

Private Sub tmrStickWithfrmMain_Timer()
'Make frmWhiteboard stick to frmMain
Me.Top = frmMain.Top
Me.Left = frmMain.Left + frmMain.Width

End Sub

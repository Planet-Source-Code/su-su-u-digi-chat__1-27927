VERSION 5.00
Begin VB.Form frmPopup 
   BorderStyle     =   0  'None
   ClientHeight    =   2400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   6000
      Left            =   3960
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3240
      Top             =   1920
   End
   Begin VB.Label lblMsgDisplay 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
   Begin VB.Image picBOX 
      BorderStyle     =   1  'Fixed Single
      Height          =   2475
      Left            =   0
      Picture         =   "frmPopup.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4755
   End
End
Attribute VB_Name = "frmPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Reverse As Boolean

Private Sub Form_Load()
On Error Resume Next

'Make form always on top
Ontop Me

'Configurations...
With Me
    .Reverse = False
    .Top = Screen.Height - GetTaskbarHeight
    .Left = Screen.Width - Me.Width
    .Height = 50
    Me.Refresh
    'Start the animation using timer
    Me.Timer1.Enabled = True
End With

    
    
End Sub



Private Sub Timer1_Timer()
On Error Resume Next
If Me.Reverse = False Then
    'The greater the number it's added on, the faster
    'will the form pop up
    Me.Height = Me.Height + 100
    Me.Top = Screen.Height - GetTaskbarHeight - Me.Height
    Me.picBOX.Refresh
    'Stop extending if the form's height is over 2400 twips
    If Me.Height >= 2400 Then
        Me.Timer1.Enabled = False
        Me.Timer2.Enabled = True
    End If
Else
    'If 'Reverse' is true then start looping backwards to
    'close the form
    Me.Height = Me.Height - 100
    Me.Top = Screen.Height - GetTaskbarHeight - Me.Height
        'If form's height is less than or equal to 50
        'then stop the timer and unload the form
        If Me.Height <= 50 Then
            Me.Timer1.Enabled = False
            Me.Timer2.Enabled = False
            Unload Me
        End If
End If
End Sub

Private Sub Timer2_Timer()
On Error Resume Next

Me.Reverse = True
Me.Timer1.Enabled = True



End Sub

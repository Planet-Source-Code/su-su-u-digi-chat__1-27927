VERSION 5.00
Begin VB.Form frmProfile 
   BorderStyle     =   0  'None
   Caption         =   "ShapedForm"
   ClientHeight    =   3000
   ClientLeft      =   2535
   ClientTop       =   3000
   ClientWidth     =   6930
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6930
   Begin VB.CommandButton Command1 
      Caption         =   "C l o s e"
      Height          =   1575
      Left            =   6600
      TabIndex        =   7
      Top             =   960
      Width           =   255
   End
   Begin VB.Timer tmrStickWithfrmMain 
      Interval        =   100
      Left            =   3120
      Top             =   720
   End
   Begin VB.Timer tmrUnloader 
      Interval        =   50
      Left            =   4200
      Top             =   600
   End
   Begin VB.Timer tmrLoader 
      Interval        =   50
      Left            =   5160
      Top             =   600
   End
   Begin VB.TextBox txtRemoteIP 
      BackColor       =   &H80000003&
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   2280
      Width           =   3735
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H80000003&
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1800
      Width           =   3735
   End
   Begin VB.TextBox txtNick 
      BackColor       =   &H80000003&
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   1320
      Width           =   3735
   End
   Begin VB.Label Label3 
      Caption         =   "IP Address :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Email         :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Nickname   :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   5055
   End
End
Attribute VB_Name = "frmProfile"
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
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Type POINTAPI
   x As Long
   y As Long
End Type
Private Const RGN_COPY = 5
Private Const CreatedBy = "VBSFC 6.2"
Private Const RegisteredTo = "Not Registered"
Private ResultRegion As Long

Public Reverse As Boolean 'Used to determine which way the animation goes

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
'!2,42,87,462,200,32,25,1
    ObjectRegion = CreateRoundRectRgn(87 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 42 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 462 * ScaleX * 15 / Screen.TwipsPerPixelX + OffsetX, 200 * ScaleY * 15 / Screen.TwipsPerPixelY + OffsetY, 64 * ScaleX * 15 / Screen.TwipsPerPixelX, 50 * ScaleY * 15 / Screen.TwipsPerPixelY)
    nRet = CombineRgn(ResultRegion, ObjectRegion, ObjectRegion, RGN_COPY)
    DeleteObject ObjectRegion
    DeleteObject HolderRegion
    CreateFormRegion = ResultRegion
End Function

Private Sub Command1_Click()
Dim GotoVal
Dim GoInto
GotoVal = Me.Height / 2
    For GoInto = 1 To GotoVal
        DoEvents
        Me.Height = Me.Height - 100
        Me.Top = (Screen.Height - Me.Height) \ 2
        If Me.Height <= 500 Then
            Exit For
            Unload Me
        End If
    Next GoInto
End Sub

Private Sub Form_paint()
    Dim nRet As Long
    nRet = SetWindowRgn(Me.hwnd, CreateFormRegion(1, 1, 0, 0), True)
    'If the above two lines are modified or moved a second copy of
    'them may be added again if the form is later Modified by VBSFC.
    
    
    'Randomly choose a gradient color to be the form's BG
    Randomize
    Dim RandomInt
    RandomInt = Rand(1, 7)
    Select Case RandomInt
        Case 1
            FadeFormPurple Me
        Case 2
            FadeFormRed Me
        Case 3
            FadeFormBlue Me
        Case 4
            FadeFormYellow Me
        Case 5
            FadeFormGreen Me
        Case 6
            FadeFormGrey Me
        Case 7
            FadeFormGrey Me
    End Select
    

   
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'Next two lines enable window drag from anywhere on form.  Remove them
'to allow window drag from title bar only.
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
End Sub
Private Sub Form_Unload(Cancel As Integer)
    DeleteObject ResultRegion
    'If the above line is modified or moved a second copy of it
    'may be added again if the form is later Modified by VBSFC.
End Sub





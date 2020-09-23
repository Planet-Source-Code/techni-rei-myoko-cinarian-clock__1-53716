VERSION 5.00
Begin VB.Form frmmain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Cinarian Clock"
   ClientHeight    =   1920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1920
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmmain.frx":0E42
   ScaleHeight     =   128
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   128
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   120
      Top             =   240
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ReleaseCapture Lib "USER32" () As Long
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Declare Function SetWindowPos Lib "User32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Const SWP_NOMOVE = &H2
    Private Const SWP_NOSIZE = &H1
    'Used to set window to always be on top or not
    Private Const HWND_NOTOPMOST = -2
    Private Const HWND_TOPMOST = -1

Public fillmode As Long

Public Sub dragform(hWnd As Long)
On Error Resume Next
  ReleaseCapture
  SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
Private Sub Form_Load()
    Me.ScaleMode = vbPixels

    Me.Move Val(GetSetting("Cinarian Clock", "Main", "Left", Me.left)), Val(GetSetting("Cinarian Clock", "Main", "Top", Me.Top))
    fillmode = Val(GetSetting("Cinarian Clock", "Main", "Fillmode", 1))
    Timer.Tag = GetSetting("Cinarian Clock", "Main", "Alarm", "12:00:?? PM") & ""
    
    setAlwaysOnTop Me.hWnd
    
    Timer_Timer
End Sub
Public Function isinregion(left, Top, Width, Height, X, Y) As Boolean
    On Error Resume Next
    If X >= left And X <= left + Width And Y >= Top And Y <= Top + Height Then isinregion = True Else isinregion = False
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If isinregion(17, 0, 110, 8, X, Y) Then dragform Me.hWnd
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If isinregion(110, 0, 8, 8, X, Y) Then Me.WindowState = vbMinimized

If isinregion(0, 0, 8, 8, X, Y) Then frmcinarian.Show vbModal, Me

If isinregion(119, 0, 8, 8, X, Y) Then
    Unload Me
    End
End If

If isinregion(0, 11, 128, 128, X, Y) Then
    If Button = vbRightButton Then fillmode = fillmode + 1
    If Button = vbLeftButton Then fillmode = fillmode - 1
    If fillmode = -1 Then fillmode = 2
    If fillmode = 3 Then fillmode = 0
    Timer_Timer
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "Cinarian Clock", "Main", "Left", Me.left
    SaveSetting "Cinarian Clock", "Main", "Top", Me.Top
    SaveSetting "Cinarian Clock", "Main", "Fillmode", fillmode
End Sub

Public Sub setAlwaysOnTop(hWnd As Long, Optional ontop As Boolean = True)
    On Error Resume Next
    If ontop = False Then Call SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 1, 1, SWP_NOMOVE Or SWP_NOSIZE)
    If ontop = True Then Call SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 1, 1, SWP_NOMOVE Or SWP_NOSIZE)
End Sub
Public Sub Timer_Timer()
    Dim currtime As String
    Me.Cls
    drawCinarian Me, 64, 68, 50, 0, 65280, Time2Cin(Now), fillmode
    currtime = Time
    currtime = left(currtime, InStrRev(currtime, ":") - 1) & right(currtime, 3)
    Me.CurrentX = (128 - TextWidth(currtime)) / 2
    Me.CurrentY = 1
    Me.Print currtime
End Sub

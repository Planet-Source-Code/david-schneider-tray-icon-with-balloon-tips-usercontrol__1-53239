VERSION 5.00
Begin VB.UserControl TrayIcon 
   CanGetFocus     =   0   'False
   ClientHeight    =   225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   ClipBehavior    =   0  'None
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   225
   ScaleWidth      =   240
   Begin VB.Frame hWndHolder 
      BorderStyle     =   0  'None
      Height          =   15
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "ST"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "TrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_BALLOONLCLK = &H405
Private Const WM_BALLOONRCLK = &H404
Private Const WM_BALLOONXCLK = WM_BALLOONRCLK
Private Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uID As Long
   uFlags As Long
   uCallbackMessage As Long
   hIcon As Long
   szTip As String * 128
   dwState As Long
   dwStateMask As Long
   szInfo As String * 256
   uTimeout As Long
   szInfoTitle As String * 64
   dwInfoFlags As Long
End Type
Private Const NOTIFYICON_VERSION = 3
Private Const NOTIFYICON_OLDVERSION = 0
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10
Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2
Private Const NIIF_NONE = &H0
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3
Private Const NIIF_INFO = &H1
Private Const NIIF_GUID = &H4
Private m_TrayIcon As StdPicture
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private m_IconData As NOTIFYICONDATA

Public Enum BalloonTipStyle
    btsNoIcon = NIIF_NONE
    btsWarning = NIIF_WARNING
    btsError = NIIF_ERROR
    btsInfo = NIIF_INFO
End Enum
Public Enum stMouseEvent
    stLeftButtonDown = WM_LBUTTONDOWN
    stLeftButtonUp = WM_LBUTTONUP
    stLeftButtonDoubleClick = WM_LBUTTONDBLCLK
    stLeftButtonClick = WM_LBUTTONUP
    stRightButtonDown = WM_RBUTTONDOWN
    stRightButtonUp = WM_RBUTTONUP
    stRightButtonDoubleClick = WM_RBUTTONDBLCLK
    stRightButtonClick = WM_RBUTTONUP
End Enum
Public Enum stBalloonClickType
    stbLeftClick = WM_BALLOONLCLK
    stbRightClick = WM_BALLOONRCLK
    stbXClick = WM_BALLOONXCLK
End Enum
Public Event TrayClick(Button As stMouseEvent)
Public Event BalloonClick(ClickType As stBalloonClickType)

Private Sub hWndHolder_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim msg As Long
    msg = X / Screen.TwipsPerPixelX
    If msg >= WM_LBUTTONDOWN And msg <= WM_RBUTTONDBLCLK Then
        RaiseEvent TrayClick(msg)
    ElseIf msg >= WM_BALLOONXCLK And msg <= WM_BALLOONLCLK Then
        RaiseEvent BalloonClick(msg)
    End If
End Sub

Private Sub UserControl_Resize()
    UserControl.Size 240, 225
End Sub

Public Property Let ToolTip(Caption As String)
    With m_IconData
        .szTip = Caption & vbNullChar
        .szInfo = "" & Chr(0)
        .szInfoTitle = "" & Chr(0)
        .dwInfoFlags = NIIF_NONE
        .uTimeout = 0
    End With
    Shell_NotifyIcon NIM_MODIFY, m_IconData
End Property

Public Property Get ToolTip() As String
    ToolTip = m_IconData.szTip
End Property

Public Sub Create(ToolTipText, Optional Icon As StdPicture)
    If Not Icon Is Nothing Then Set m_TrayIcon = Icon
    With m_IconData
        .cbSize = Len(m_IconData)
        .hwnd = hWndHolder.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        If Not m_TrayIcon Is Nothing Then .hIcon = m_TrayIcon
        If Not IsMissing(ToolTipText) Then .szTip = ToolTipText & vbNullChar
        .dwState = 0
        .dwStateMask = 0
        .szInfo = "" & Chr(0)
        .szInfoTitle = "" & Chr(0)
        .dwInfoFlags = NIIF_NONE
    End With
    Shell_NotifyIcon NIM_ADD, m_IconData
End Sub

Public Sub Remove()
    Shell_NotifyIcon NIM_DELETE, m_IconData
End Sub

Public Sub BalloonTip(Prompt As String, Optional Style As BalloonTipStyle = btsNoIcon, Optional Title As String, Optional Timeout As Long = 2000)
    If Title = Empty Then Title = App.Title
    If Prompt = Empty Then Prompt = " "
    With m_IconData
        .szInfo = Prompt & Chr(0)
        .szInfoTitle = Title & Chr(0)
        .dwInfoFlags = Style
        .uTimeout = Timeout
    End With
    Shell_NotifyIcon NIM_MODIFY, m_IconData
End Sub

Public Sub PopupMenu(Menu As Menu, Optional Flags, Optional DefaultMenu)
    SetForegroundWindow Menu.Parent.hwnd
    If IsMissing(Flags) And IsMissing(DefaultMenu) Then
        Menu.Parent.PopupMenu Menu
    ElseIf IsMissing(Flags) Then
        Menu.Parent.PopupMenu Menu, , , , DefaultMenu
    Else
        Menu.Parent.PopupMenu Menu, Flags, , , DefaultMenu
    End If
End Sub

Property Set Icon(Icon As StdPicture)
    Set m_TrayIcon = Icon
    With m_IconData
        .hIcon = m_TrayIcon
        .szInfo = "" & Chr(0)
        .szInfoTitle = "" & Chr(0)
        .dwInfoFlags = NIIF_NONE
        .uTimeout = 0
    End With
    Shell_NotifyIcon NIM_MODIFY, m_IconData
End Property

Property Get Icon() As StdPicture
    Set Icon = m_TrayIcon
End Property

Private Sub UserControl_Terminate()
    Remove
End Sub

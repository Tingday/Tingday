Attribute VB_Name = "ÍÐÅÌÍ¼±ê"
Option Explicit
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Const MAX_TOOLTIP As Integer = 64
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_TIP = &H4
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Private Type NOTIFYICONDATA
   cbSize           As Long
   hWnd             As Long
   uID              As Long
   uFlags           As Long
   uCallbackMessage As Long
   hIcon            As Long
   szTip            As String * MAX_TOOLTIP
End Type
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private dIcon As NOTIFYICONDATA
Private c_hwnd As Long
Private c_hIcon As Long
Public Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long

Public Sub SNIcon_Add(hIcon As Long, hWnd As Long, showText As String)    'Ìí¼ÓÍ¼±ê
    c_hwnd = hWnd
    c_hIcon = hIcon
    With dIcon
        .cbSize = Len(dIcon)
        .hWnd = c_hwnd
        .uID = 1&
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = c_hIcon
        .szTip = showText & vbNullChar
    End With
    Call Shell_NotifyIcon(NIM_ADD, dIcon)
End Sub

Public Sub SNIcon_Refresh(Optional hIcon As Long, Optional hWnd As Long, Optional showText As String) 'Ë¢ÐÂÍ¼±ê
Dim c_showText As String
    If hWnd <> 0 Then c_hwnd = hWnd
    If hIcon <> 0 Then c_hIcon = hIcon
    If showText <> "" Then c_showText = showText
    dIcon.szTip = c_showText & vbNullChar
    dIcon.hWnd = c_hwnd
    dIcon.hIcon = hIcon
    Call Shell_NotifyIcon(NIM_MODIFY, dIcon)
End Sub

Public Sub SNIcon_Del() 'É¾³ýÍ¼±ê
    Call Shell_NotifyIcon(NIM_DELETE, dIcon)
End Sub

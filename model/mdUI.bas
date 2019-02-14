Attribute VB_Name = "ΩÁ√Ê"
Option Explicit
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public FormA  As RECT

Public Sub setTrsp(ByVal Trsp As Long, ByVal hWnd As Long)
    Dim Trs As Long
    Trs = GetWindowLong(hWnd, GWL_EXSTYLE)
    Trs = Trs Or WS_EX_LAYERED
    SetWindowLong hWnd, GWL_EXSTYLE, Trs
    SetLayeredWindowAttributes hWnd, 0, Trsp, LWA_ALPHA
End Sub

Public Sub setTop(ByVal hWnd As Long)
'¥∞ÃÂ÷√∂•
SetWindowPos hWnd, -1, 0, 0, 0, 0, 3
End Sub


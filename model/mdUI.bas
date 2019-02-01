Attribute VB_Name = "界面"
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

Public Sub hit(ByRef mForm As RECT, ByVal mForm2 As Long)
    Dim mCursor As POINTAPI '获取鼠标的位置
    '检查是否应该缩到顶部
    GetCursorPos mCursor '获取位置
    If mCursor.X > mForm.Left And mCursor.X < mForm.Right And mCursor.Y > mForm.Top - 5 And mCursor.Y < mForm.Bottom Then
        '鼠标在窗体里面
        Select Case mForm2
        Case 1
            If frmMain.Top < 0 Then
                '窗体动画
                Do
                DoEvents
                frmMain.Move frmMain.Left, frmMain.Top + 30
                If frmMain.Top > 0 Then frmMain.Top = 0
                Loop Until frmMain.Top >= 0
                setTop frmMain.hWnd
            End If
            
        Case 2
            If frmLRC.Top < 0 Then
                Do
                DoEvents
                frmLRC.Move frmLRC.Left, frmLRC.Top + 30
                If frmLRC.Top > 0 Then frmLRC.Top = 0
                Loop Until frmLRC.Top = 0
                setTop frmLRC.hWnd
            End If
        End Select
    Else
        Select Case mForm2
        Case 1
            If frmMain.Top = 0 Then
                Do
                DoEvents
                
                frmMain.Move frmMain.Left, frmMain.Top - 30
                Loop Until frmMain.Top <= -frmMain.Height + 50
            End If
        Case 2
            If frmLRC.Top = 0 Then
                Do
                DoEvents
                
                frmLRC.Move frmLRC.Left, frmLRC.Top - 30
                Loop Until frmLRC.Top <= -frmLRC.Height + 50
            End If
        End Select
    End If
End Sub

Public Sub setTop(ByVal hWnd As Long)
'窗体置顶
SetWindowPos hWnd, -1, 0, 0, 0, 0, 3
End Sub


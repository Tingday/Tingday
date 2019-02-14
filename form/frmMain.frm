VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "TingDay"
   ClientHeight    =   7620
   ClientLeft      =   5655
   ClientTop       =   1080
   ClientWidth     =   10635
   FillColor       =   &H80000011&
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   10635
   Begin VB.Menu 关于 
      Caption         =   "关于"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cTingDay As LibZPlayCore

Private Sub Form_Load()
    Me.Move 4080, 240
    Set cTingDay = New LibZPlayCore
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture '移动窗体
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Me.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    cTingDay.stopPlay
    Set cTingDay = Nothing
    Unload frmMenu
End Sub


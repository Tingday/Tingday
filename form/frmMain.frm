VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "TingDay"
   ClientHeight    =   2130
   ClientLeft      =   5655
   ClientTop       =   780
   ClientWidth     =   3660
   FillColor       =   &H80000011&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3660
   Begin VB.CommandButton cmdPlay 
      Caption         =   "播放"
      Height          =   1095
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cTingDay As LibZPlayCore

Private Sub cmdPlay_Click()
    If cmdPlay.Caption <> "播放" Then
        cmdPlay.Caption = "播放"
    Else
        cmdPlay.Caption = "暂停"
    End If
End Sub

Private Sub Form_Load()
    Me.Move 4080, 240
    Load frmMenu
    Set cTingDay = New LibZPlayCore
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture '移动窗体
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then Me.Show
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Me.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    cTingDay.stopPlay
    Set cTingDay = Nothing
    Unload frmMenu
End Sub


VERSION 5.00
Begin VB.Form frmBusy 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1545
   ClientLeft      =   7260
   ClientTop       =   4095
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Begin VB.Label lbTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TingDay如咖啡般，与你同在。"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3330
   End
End
Attribute VB_Name = "frmBusy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
lbTitle.Caption = "TingDay如咖啡般，与你同在。" & vbCrLf & _
                "正在为首次运行配置文件。"
setTrsp 210, Me.hWnd
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 '移动窗体
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
End Sub

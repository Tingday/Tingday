VERSION 5.00
Begin VB.Form frmAbout 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "����"
   ClientHeight    =   2670
   ClientLeft      =   5970
   ClientTop       =   2130
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   ScaleHeight     =   2670
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   Begin VB.Label lbshow 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   60
   End
   Begin VB.Label lb_End 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "r"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   330
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
lbshow.Caption = "�汾��20150211" & vbCrLf & _
                "���ߣ�0yufan0 " & vbCrLf & _
                "���䣺woyufan@163.com" & vbCrLf & _
                "TingDay,���Ȱ�����ͬ�ڡ�"
setTrsp 210, Me.hWnd
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 '�ƶ�����
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
End Sub

Private Sub lb_End_Click()
Unload Me
End Sub

Private Sub lbshow_Click()
 '�ƶ�����
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
End Sub

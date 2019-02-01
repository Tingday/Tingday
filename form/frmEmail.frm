VERSION 5.00
Begin VB.Form frmEmail 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "∑¥¿°"
   ClientHeight    =   3195
   ClientLeft      =   7080
   ClientTop       =   2415
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtQQ 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2520
      Width           =   1935
   End
   Begin VB.TextBox txtShow 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Í«≥∆£∫"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   630
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
      Left            =   4920
      TabIndex        =   5
      Top             =   120
      Width           =   330
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "¡™œµ∑Ω Ω£∫"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   360
      TabIndex        =   4
      Top             =   2520
      Width           =   1050
   End
   Begin VB.Label lbClick 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ÃÓ∫√¡À"
      BeginProperty Font 
         Name            =   "Œ¢»Ì—≈∫⁄"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   2400
      Width           =   1125
   End
End
Attribute VB_Name = "frmEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    setTrsp 210, Me.hWnd
End Sub


Private Sub lb_End_Click()
Unload Me
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 '“∆∂Ø¥∞ÃÂ
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
      
End Sub




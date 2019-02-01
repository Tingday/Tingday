VERSION 5.00
Begin VB.Form frmLRC 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "LRC"
   ClientHeight    =   10620
   ClientLeft      =   7110
   ClientTop       =   0
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   ScaleHeight     =   708
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   4560
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   4560
      Top             =   1320
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   10
      Left            =   1440
      TabIndex        =   10
      Top             =   6840
      Width           =   75
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   1
      Left            =   1440
      TabIndex        =   9
      Top             =   2400
      Width           =   75
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   4
      Left            =   1440
      TabIndex        =   8
      Top             =   3840
      Width           =   75
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   1440
      TabIndex        =   7
      Top             =   3360
      Width           =   75
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   1440
      TabIndex        =   6
      Top             =   2880
      Width           =   75
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   9
      Left            =   1440
      TabIndex        =   5
      Top             =   6360
      Width           =   75
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   8
      Left            =   1440
      TabIndex        =   4
      Top             =   5880
      Width           =   75
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   7
      Left            =   1440
      TabIndex        =   3
      Top             =   5400
      Width           =   75
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   6
      Left            =   1920
      TabIndex        =   2
      Top             =   4920
      Width           =   75
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Index           =   5
      Left            =   960
      TabIndex        =   1
      Top             =   4320
      Width           =   90
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
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Width           =   330
   End
End
Attribute VB_Name = "frmLRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iiii As Integer
Private Sub Form_Load()
Dim I As Integer
Me.Height = frmMain.Height
setTrsp 210, Me.hWnd
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'ÒÆ¶¯´°Ìå
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
End Sub



Private Sub lb_End_Click()
frmMain.lbLyric.ForeColor = &H808080
Unload Me
End Sub

Private Sub Timer1_Timer()
Dim mForm As RECT
mForm.Left = ScaleX(Me.Left, vbTwips, vbPixels)
mForm.Right = ScaleX(Me.Width, vbTwips, vbPixels) + ScaleX(Me.Left, vbTwips, vbPixels)
mForm.Top = ScaleY(Me.Top, vbTwips, vbPixels)
mForm.Bottom = ScaleY(Me.Height, vbTwips, vbPixels) + ScaleY(Me.Top, vbTwips, vbPixels)
'´°Ìå¶¥²¿Òþ²Ø
hit mForm, 2
End Sub


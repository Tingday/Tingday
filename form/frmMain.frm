VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "TingDay"
   ClientHeight    =   9930
   ClientLeft      =   5655
   ClientTop       =   780
   ClientWidth     =   9705
   FillColor       =   &H80000011&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9930
   ScaleWidth      =   9705
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      ForeColor       =   &H80000008&
      Height          =   90
      Left            =   3240
      ScaleHeight     =   1
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   16
      Top             =   2640
      Width           =   900
      Begin VB.Line liVolume 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   8
         X1              =   0
         X2              =   90
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.ListBox lstSongs 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   5730
      Left            =   360
      TabIndex        =   4
      Top             =   3360
      Width           =   8775
   End
   Begin VB.Timer tmrPlayer 
      Interval        =   1000
      Left            =   0
      Top             =   3360
   End
   Begin VB.Timer tmrForm 
      Interval        =   100
      Left            =   4920
      Top             =   4560
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   105
      Left            =   360
      ScaleHeight     =   1
      ScaleMode       =   0  'User
      ScaleWidth      =   93.916
      TabIndex        =   0
      Top             =   1080
      Width           =   3735
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   5
         X1              =   0
         X2              =   0.482
         Y1              =   0.4
         Y2              =   0.4
      End
   End
   Begin VB.Label lbLyric 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "´Ê"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   3480
      TabIndex        =   17
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lblSound 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "U"
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
      Height          =   375
      Left            =   2760
      TabIndex        =   15
      ToolTipText     =   "ÒôÁ¿µ÷½Ú/¾²Òô"
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   14
      ToolTipText     =   "ËÑË÷ÍøÂç¸èÇú"
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label lbOpen 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   15.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   13
      ToolTipText     =   "Ìí¼ÓÎÄ¼þ¼Ð"
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label lbPosition 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   24
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2160
      TabIndex        =   12
      ToolTipText     =   "¶¨Î»µ½µ±Ç°²¥·ÅµÄ¸èÇú"
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label lbPlayMode 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Ë³"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   600
      TabIndex        =   11
      ToolTipText     =   "µ±Ç°²¥·ÅµÄÄ£Ê½"
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TingDay£¬¿§·È°ãÓëÄãÍ¬ÔÚ¡£"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
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
      TabIndex        =   10
      Top             =   9480
      Width           =   2730
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   4080
      Y1              =   9240
      Y2              =   9240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   360
      X2              =   4080
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label LblClose 
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
      TabIndex        =   9
      Top             =   120
      Width           =   330
   End
   Begin VB.Label lbNext 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   2640
      TabIndex        =   8
      Top             =   1680
      Width           =   450
   End
   Begin VB.Label lbPrev 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   21.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Left            =   3480
      TabIndex        =   6
      Top             =   120
      Width           =   330
   End
   Begin VB.Label lbPlay 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   42
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   1680
      TabIndex        =   5
      Top             =   1440
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   1320
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   465
   End
   Begin VB.Label lbTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TingDay"
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
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   930
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
    ReleaseCapture 'ÒÆ¶¯´°Ìå
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If LblClose.ForeColor = vbRed Then LblClose.ForeColor = vbWhite 'X
    If Label6.ForeColor = &H2469F6 Then Label6.ForeColor = vbWhite
    If lbPosition.ForeColor = &H2469F6 Then lbPosition.ForeColor = vbWhite
    If lbPlay.ForeColor = &H2469F6 Then lbPlay.ForeColor = vbWhite
    If lbOpen.ForeColor = &H2469F6 Then lbOpen.ForeColor = vbWhite
    If lbPrev.ForeColor = &H2469F6 Then lbPrev.ForeColor = vbWhite
    If lbNext.ForeColor = &H2469F6 Then lbNext.ForeColor = vbWhite
    If Label15.ForeColor = &H2469F6 Then Label15.ForeColor = vbWhite
    If lbPlayMode.ForeColor = &H2469F6 Then lbPlayMode.ForeColor = vbWhite
    If lblSound.ForeColor = &H2469F6 Then lblSound.ForeColor = vbWhite
    If Button = 1 Then Me.Show
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    cTingDay.stopPlay
    Set cTingDay = Nothing
    Unload frmMenu
End Sub

Private Sub LblClose_Click()
    Unload Me
End Sub

Private Sub lbOpen_Click()
    Dim tDialog As New DiaLogBox
    Dim tSeeker As New Seeker
    Dim tFoldPath As String
    Dim tSongs As String
    Dim ta As Variant
    Dim i As Integer
    tFoldPath = tDialog.showFolder("Ñ¡Ôñ¸èÇúÎÄ¼þ¼Ð", Me.hWnd)
    tSongs = tSeeker.songs(tFoldPath, ".mp3")
    ta = Split(tSongs, "|")
    For i = 0 To UBound(ta)
        lstSongs.AddItem ta(i)
    Next
    Set tDialog = Nothing
End Sub

Private Sub lbPlayMode_Click()
    Dim playStyle As Integer
    playStyle = cTingDay.playStyle
    Select Case playStyle
    Case 0 'µ¥Çú
        cTingDay.playStyle = 1
        lbPlayMode = "Ëæ"
    Case 1 '
        cTingDay.playStyle = 2
        lbPlayMode = "Ë³"
    Case 2
        cTingDay.playStyle = 3
        lbPlayMode = "Ñ­"
    Case 3
        cTingDay.playStyle = 0
        lbPlayMode = "µ¥"
    End Select
End Sub

Private Sub lbPlayMode_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lbPlayMode.ForeColor = vbWhite Then lbPlayMode.ForeColor = &H2469F6
End Sub

Private Sub lbPosition_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lbPosition.ForeColor = vbWhite Then lbPosition.ForeColor = &H2469F6
End Sub

Private Sub lbOpen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lbOpen.ForeColor = vbWhite Then lbOpen.ForeColor = &H2469F6
End Sub

Private Sub label15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Label15.ForeColor = vbWhite Then Label15.ForeColor = &H2469F6
End Sub

Private Sub lblSound_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lblSound.ForeColor = vbWhite Then lblSound.ForeColor = &H2469F6
End Sub

Private Sub lblSound_Click() '¾²Òô
    Dim tNoSound As Boolean
    tNoSound = cTingDay.noSound
    If tNoSound = True Then
        cTingDay.volume = cTingDay.volume
        If cTingDay.volume = 0 Then
            cTingDay.volume = 80
        End If
        lblSound.ForeColor = vbWhite
    Else
        cTingDay.volume = 0
        lblSound.ForeColor = vbRed
    End If
End Sub

Private Sub lbPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lbPlay.ForeColor = vbWhite Then lbPlay.ForeColor = &H2469F6
End Sub

Private Sub lbLyric_Click()
If frmLRC.Visible = True Then
    Unload frmLRC
    lbLyric.ForeColor = &H808080
Else
    With frmLRC
    .Left = Me.Left + Me.Width + 30
    .Top = Me.Top
    .Show
    End With
    lbLyric.ForeColor = &HFFFFFF
End If
End Sub

Private Sub Label6_Click()
    Me.Hide
    Me.WindowState = 0
End Sub

Private Sub Label6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Label6.ForeColor = vbWhite Then Label6.ForeColor = &H2469F6
End Sub

Private Sub lbPrev_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lbPrev.ForeColor = vbWhite Then lbPrev.ForeColor = &H2469F6
End Sub

Private Sub lbNext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lbNext.ForeColor = vbWhite Then lbNext.ForeColor = &H2469F6
End Sub

Private Sub LblClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If LblClose.ForeColor = vbWhite Then LblClose.ForeColor = vbRed
End Sub

Private Sub lstSongs_Click()
    cTingDay.loadSingleFile lstSongs.List(lstSongs.ListIndex)
    cTingDay.play
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X > 100 Then Exit Sub
cTingDay.volume = X
liVolume.X2 = X
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X > 100 Then Exit Sub
If Button = 1 Then
    cTingDay.volume = X
    liVolume.X2 = X
End If
End Sub

Private Sub tmrForm_Timer()
    Dim mForm As RECT
    mForm.Left = ScaleX(Me.Left, vbTwips, vbPixels)
    mForm.Right = ScaleX(Me.Width, vbTwips, vbPixels) + ScaleX(frmMain.Left, vbTwips, vbPixels)
    mForm.Top = ScaleY(frmMain.Top, vbTwips, vbPixels)
    mForm.Bottom = ScaleY(frmMain.Height, vbTwips, vbPixels) + ScaleY(frmMain.Top, vbTwips, vbPixels)
      '´°Ìå¶¥²¿Òþ²Ø
    hit mForm, 1
End Sub



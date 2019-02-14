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
End
Attribute VB_Name = "frmLRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 'ÒÆ¶¯´°Ìå
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&
End Sub

Private Sub lb_End_Click()
    Unload Me
End Sub


VERSION 5.00
Begin VB.Form frmMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "TingDay"
   ClientHeight    =   2655
   ClientLeft      =   8160
   ClientTop       =   4080
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   800
      Left            =   720
      Top             =   1080
   End
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   2040
      Top             =   1200
   End
   Begin VB.Menu �赥�˵� 
      Caption         =   "�赥�˵�"
      Begin VB.Menu ����ļ��� 
         Caption         =   "����ļ���"
      End
      Begin VB.Menu �򿪸赥 
         Caption         =   "�򿪸赥"
      End
      Begin VB.Menu �ָ���_�˵�1 
         Caption         =   "-"
      End
      Begin VB.Menu ������ 
         Caption         =   "������"
      End
      Begin VB.Menu �ָ���2_�˵�1 
         Caption         =   "-"
      End
      Begin VB.Menu �½��赥 
         Caption         =   "�½��赥"
      End
      Begin VB.Menu ɾ���赥 
         Caption         =   "ɾ���赥"
      End
      Begin VB.Menu ����赥 
         Caption         =   "����赥"
      End
   End
   Begin VB.Menu ���̲˵� 
      Caption         =   "���̲˵�"
      Begin VB.Menu ��ʾ 
         Caption         =   "TingDay"
      End
      Begin VB.Menu ���̲˵�_�ָ���1 
         Caption         =   "-"
      End
      Begin VB.Menu ������� 
         Caption         =   "�������"
      End
      Begin VB.Menu ���̲˵�_�ָ���4 
         Caption         =   "-"
      End
      Begin VB.Menu ͣ�� 
         Caption         =   "����"
      End
      Begin VB.Menu ��һ�� 
         Caption         =   "��һ��"
      End
      Begin VB.Menu ��һ�� 
         Caption         =   "��һ��"
      End
      Begin VB.Menu ���̲˵�_�ָ���2 
         Caption         =   "-"
      End
      Begin VB.Menu �����ȼ� 
         Caption         =   "�����ȼ�"
      End
      Begin VB.Menu ���̲˵�_�ָ���3 
         Caption         =   "-"
      End
      Begin VB.Menu ���� 
         Caption         =   "����"
      End
      Begin VB.Menu ���� 
         Caption         =   "����"
      End
      Begin VB.Menu �ָ���1 
         Caption         =   "-"
      End
      Begin VB.Menu �˳� 
         Caption         =   "�˳�"
      End
   End
   Begin VB.Menu �����˵� 
      Caption         =   "�����˵�"
      Begin VB.Menu ����ļ� 
         Caption         =   "����ļ�"
      End
      Begin VB.Menu ���ļ�λ�� 
         Caption         =   "���ļ�λ��"
      End
      Begin VB.Menu ɾ������ 
         Caption         =   "ɾ������"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Visible = False
    SNIcon_Add Me.Icon, Me.hWnd, "tingday����"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    frmMain.Show
    frmMain.WindowState = 0
    If frmMain.Top < 0 Then frmMain.Top = 100
ElseIf Button = 2 Then
    Me.PopupMenu ���̲˵�
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SNIcon_Del
    End
End Sub

Private Sub ����_Click()
frmEmail.Show
End Sub

Private Sub ����_Click()
frmAbout.Show
End Sub

Private Sub �˳�_Click()
    Unload Me
End Sub

Private Sub ��ʾ_Click()
    frmMain.WindowState = 0
    frmMain.Show
End Sub

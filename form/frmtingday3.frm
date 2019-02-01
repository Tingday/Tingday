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
   Begin VB.Menu 歌单菜单 
      Caption         =   "歌单菜单"
      Begin VB.Menu 添加文件夹 
         Caption         =   "添加文件夹"
      End
      Begin VB.Menu 打开歌单 
         Caption         =   "打开歌单"
      End
      Begin VB.Menu 分割线_菜单1 
         Caption         =   "-"
      End
      Begin VB.Menu 重命名 
         Caption         =   "重命名"
      End
      Begin VB.Menu 分割线2_菜单1 
         Caption         =   "-"
      End
      Begin VB.Menu 新建歌单 
         Caption         =   "新建歌单"
      End
      Begin VB.Menu 删除歌单 
         Caption         =   "删除歌单"
      End
      Begin VB.Menu 保存歌单 
         Caption         =   "保存歌单"
      End
   End
   Begin VB.Menu 托盘菜单 
      Caption         =   "托盘菜单"
      Begin VB.Menu 显示 
         Caption         =   "TingDay"
      End
      Begin VB.Menu 托盘菜单_分割线1 
         Caption         =   "-"
      End
      Begin VB.Menu 随机播放 
         Caption         =   "随机播放"
      End
      Begin VB.Menu 托盘菜单_分割线4 
         Caption         =   "-"
      End
      Begin VB.Menu 停播 
         Caption         =   "播放"
      End
      Begin VB.Menu 上一曲 
         Caption         =   "上一曲"
      End
      Begin VB.Menu 下一曲 
         Caption         =   "下一曲"
      End
      Begin VB.Menu 托盘菜单_分割线2 
         Caption         =   "-"
      End
      Begin VB.Menu 禁用热键 
         Caption         =   "禁用热键"
      End
      Begin VB.Menu 托盘菜单_分割线3 
         Caption         =   "-"
      End
      Begin VB.Menu 反馈 
         Caption         =   "反馈"
      End
      Begin VB.Menu 关于 
         Caption         =   "关于"
      End
      Begin VB.Menu 分割线1 
         Caption         =   "-"
      End
      Begin VB.Menu 退出 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu 歌曲菜单 
      Caption         =   "歌曲菜单"
      Begin VB.Menu 添加文件 
         Caption         =   "添加文件"
      End
      Begin VB.Menu 打开文件位置 
         Caption         =   "打开文件位置"
      End
      Begin VB.Menu 删除歌曲 
         Caption         =   "删除歌曲"
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
    SNIcon_Add Me.Icon, Me.hWnd, "tingday音乐"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    frmMain.Show
    frmMain.WindowState = 0
    If frmMain.Top < 0 Then frmMain.Top = 100
ElseIf Button = 2 Then
    Me.PopupMenu 托盘菜单
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SNIcon_Del
    End
End Sub

Private Sub 反馈_Click()
frmEmail.Show
End Sub

Private Sub 关于_Click()
frmAbout.Show
End Sub

Private Sub 退出_Click()
    Unload Me
End Sub

Private Sub 显示_Click()
    frmMain.WindowState = 0
    frmMain.Show
End Sub

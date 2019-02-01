VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DiaLogBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetColorAdjustment Lib "gdi32" (ByVal hdc As Long, lpca As COLORADJUSTMENT) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Const BIF_RETURNONLYFSDIRS = &H1
Private pidl As Long
Private Type COLORADJUSTMENT
caSize As Integer
        caFlags As Integer
        caIlluminantIndex As Integer
        caRedGamma As Integer
        caGreenGamma As Integer
        caBlueGamma As Integer
        caReferenceBlack As Integer
        caReferenceWhite As Integer
        caContrast As Integer
        caBrightness As Integer
        caColorfulness As Integer
        caRedGreenTint As Integer
End Type

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

'打开文件对话框
Public Function ShowOpen(ByVal yTitle As String, ByVal Filter As String, ByVal hWnd As Long) As String
    Dim I As Integer
    Dim OpenFolder As OPENFILENAME
    Dim fileName As String
    With OpenFolder
        .lStructSize = Len(OpenFolder)
        .hwndOwner = hWnd
        .hInstance = App.hInstance
        .lpstrFile = Space(254)
        .nMaxFile = 255
        .lpstrFileTitle = Space(254)
        .nMaxFileTitle = 255
        '.lpstrInitialDir = App.path
        .flags = 6148
        '过滤文件类型
        .lpstrFilter = Filter
        '对话框文字
        .lpstrTitle = yTitle
        '显示对话框
        I = GetOpenFileName(OpenFolder)
        If I >= 1 Then
            ShowOpen = .lpstrFile
        End If
    End With
End Function

'保存文件对话框
Public Function ShowSave(ByVal yTitle As String, ByVal Filter As String, ByVal hWnd As Long) As String
    Dim I As Integer
    Dim saveFolder As OPENFILENAME
    Dim fileName As String
    With saveFolder
        .lStructSize = Len(saveFolder)
        .hwndOwner = hWnd
        .hInstance = App.hInstance
        .lpstrFile = Space(254)
        .nMaxFile = 255
        .lpstrFileTitle = Space(254)
        .nMaxFileTitle = 255
        '.lpstrInitialDir = App.path
        .flags = 6148
        '过滤文件类型
        .lpstrFilter = Filter
        '对话框文字
        .lpstrTitle = yTitle
        '显示对话框
        I = GetSaveFileName(saveFolder)
        If I >= 1 Then
            ShowSave = .lpstrFile
        End If
    End With
End Function

'打开文件夹
Public Function showFolder(ByVal yTitle As String, ByVal hWnd As Long) As String     '打开文件夹 并选择一个文件夹
Dim bi As BROWSEINFO
Dim r As Long
Dim pidl As Long
Dim path As String
Dim pos As Integer
'句柄
bi.hOwner = hWnd
'展开根目录
bi.pidlRoot = 0&
'列表框标题
bi.lpszTitle = yTitle
'规定只能选择文件夹，其他无效
bi.ulFlags = BIF_RETURNONLYFSDIRS
'调用API函数显示列表框
pidl = SHBrowseForFolder(bi)
'利用API函数获取返回的路径
path = Space$(512)
r = SHGetPathFromIDList(ByVal pidl&, ByVal path)
If r Then
    pos = InStr(path, Chr$(0))
    showFolder = Left(path, pos - 1)
Else
    showFolder = ""
End If
End Function




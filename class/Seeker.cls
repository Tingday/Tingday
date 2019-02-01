VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Seeker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function songs(ByVal Folder As String, ByVal Filter As String) As String
    Dim fso As New FileSystemObject
    Dim mFile As String
    Dim objFolder As Object
    Dim path As Variant
    Dim Files As Variant
    Dim saveFolder As String
    Dim mPath As Variant
    Dim objPath As Object
    Dim objPath2 As Object
    Dim mPath2 As Variant
    If Folder = "" Then Exit Function
        songs = getFiles(Folder, Filter)
        '搜索本目录下的其它文件夹
        Set objFolder = fso.GetFolder(Folder) '设置要搜索的文件
        For Each path In objFolder.SubFolders
            songs = songs & "|" & getFiles(path, Filter)
            Set objPath = fso.GetFolder(path)
                    For Each mPath In objPath.SubFolders '第二层子目录的文件
                        songs = songs & "|" & getFiles(mPath, Filter)
                        Set objPath2 = fso.GetFolder(mPath) '第三层子目录的文件
                            For Each mPath2 In objPath2.SubFolders
                                 songs = songs & "|" & getFiles(mPath2, Filter)
                            Next
                        Set fso = Nothing
                    Next
            Set fso = Nothing
        Next
        Set fso = Nothing
End Function

Private Function getFiles(ByVal mPath As String, ByVal mFilter As String) As String
    Dim File As Variant
    Dim fso As New FileSystemObject
    For Each File In fso.GetFolder(mPath).Files
        If mFilter <> ".*" Then
            If Right(File, 4) = mFilter Then
                If getFiles = "" Then
                    getFiles = File
                Else
                    getFiles = getFiles & "|" & File
                End If
            End If
        End If
    Next
End Function

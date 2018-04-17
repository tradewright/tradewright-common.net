#Region "License"

' The MIT License (MIT)
'
' Copyright (c) 2017 Richard L King (TradeWright Software Systems)
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

#End Region

Imports System.IO

Public Module DataStorage

    Private mAppDataPathUser As String

    Public Property ApplicationDataPathUser As String
        Get
            If String.IsNullOrEmpty(mAppDataPathUser) Then Throw New InvalidOperationException("LocalUserAppDataPath has not yet been set")
            Return mAppDataPathUser
        End Get
        Set
            If Not String.IsNullOrEmpty(mAppDataPathUser) Then Throw New InvalidOperationException("LocalUserAppDataPath has already been set")
            mAppDataPathUser = Value
        End Set
    End Property

    Public Function CreateWriteableTextFile(
                filePath As String,
                Optional append As Boolean = False,
                Optional createBackup As Boolean = False,
                Optional incrementFilenameIfInUse As Boolean = False) As FileStream
        If String.IsNullOrEmpty(filePath) Then Throw New ArgumentException("No filepath specified")

        Try
            If (File.GetAttributes(filePath) And FileAttributes.Directory) <> 0 Then Throw New ArgumentException("Specified file is a directory")
        Catch e As FileNotFoundException
        Catch e As IOException
        End Try

        Dim directory = Path.GetDirectoryName(filePath)
        Dim filename = Path.GetFileName(filePath)
        Dim extension = Path.GetExtension(filePath)

        If Not createBackup Then
            ' nothing to do here
        ElseIf append Then
            ' ignore createBackup if we're appending
        ElseIf Not File.Exists(filePath) Then
            ' nothing to back up
        Else
            createBackupFile(filePath, directory, filename, extension)
        End If

        Return createNewFile(filePath, directory, filename, extension, append, incrementFilenameIfInUse)
    End Function

    Private Sub createBackupFile(
                    originalfilePath As String,
                    directory As String,
                    filename As String,
                    extension As String)
        Dim newFilename = $"{directory}{Path.DirectorySeparatorChar}{filename}.bak.{extension}"
        If File.Exists(newFilename) Then
            Dim i = 0
            Do
                i = i + 1
                newFilename = $"{directory}{Path.DirectorySeparatorChar}{filename}.bak{i}.{extension}"
            Loop Until Not File.Exists(newFilename)
        End If
        Try
            File.Move(originalfilePath, newFilename)
        Catch e As IOException
        Catch e As UnauthorizedAccessException
        End Try
    End Sub

    Private Function createNewFile(
                    originalfilePath As String,
                    directory As String,
                    filename As String,
                    extension As String,
                    append As Boolean,
                    increment As Boolean) As FileStream
        Dim fileInfo = New FileInfo(originalfilePath)
        Try
            Return fileInfo.Open(If(append, FileMode.Append, FileMode.Create), FileAccess.Write)
        Catch e As System.UnauthorizedAccessException

            If Not increment Then Throw New UnauthorizedAccessException("File already in use or access denied")

            ' now increment the filename until we find one we can use
            Dim i = 1
            Do
                Dim filePath = $"{directory}{Path.DirectorySeparatorChar}{filename}-{i}.{extension}"
                If Not File.Exists(filePath) Then
                    Try
                        IO.Directory.CreateDirectory(Path.GetDirectoryName(filePath))
                        Dim fi = New FileInfo(filePath)
                        Dim fs = fi.Open(If(append, FileMode.Append, FileMode.Create), FileAccess.Write)
                        Return fs
                    Catch ex As IOException
                    End Try
                End If
                i = i + 1
            Loop
        End Try
    End Function
End Module

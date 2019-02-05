#Region "License"

' The MIT License (MIT)
'
' Copyright (c) 2017-2019 Richard L King (TradeWright Software Systems)
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

    ''' <summary>
    '''     Creates a <c>FileStream</c> object for a file that allows write access and 
    '''     concurrent read access.
    ''' </summary>
    ''' <param name="filePath">
    '''     The path to the file for a which a <c>FileStream</c> is required.
    '''     
    '''     Note that the actual file path used may be modified as a result of
    '''     using the <c>createBackup</c> and <c>incrementFilenameIfInUse</c> parameters.
    ''' </param>
    ''' <param name="append">
    '''     The filestream is to be opened for appending. If this is <c>True</c>,
    '''     the <c>createBackup</c> parameter is ignored.
    ''' </param>
    ''' <param name="createBackup">
    '''     If a file with the specified name already exists, that file is renamed with 
    '''     a <c>.bakn</c> string before the extension. n is blank for the first such
    '''     rename, 1 for the second, 2 for the third, and so on.
    ''' </param>
    ''' <param name="incrementFilenameIfInUse">
    '''     If the relevant file is already in use, or access is denied, then
    '''     the filename is modified by appending <c>-n</c> before the extension,
    '''     where n is successively incremented until an available filename is found.
    ''' </param>
    ''' <returns>A <c>FileStream</c> object</returns>
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

        If Not createBackup Then
            ' nothing to do here
        ElseIf append Then
            ' ignore createBackup if we're appending
        ElseIf Not File.Exists(filePath) Then
            ' nothing to back up
        Else
            createBackupFile(filePath)
        End If

        Return createNewFile(filePath, append, incrementFilenameIfInUse)
    End Function

    Private Sub createBackupFile(
                    originalfilePath As String)
        Dim directory = Path.GetDirectoryName(originalfilePath)
        Dim filename = Path.GetFileNameWithoutExtension(originalfilePath)
        Dim extension = Path.GetExtension(originalfilePath)

        Dim newFilename = $"{directory}{Path.DirectorySeparatorChar}{filename}.bak{extension}"
        Dim i = 0
        Do
            If Not File.Exists(newFilename) Then
                Try
                    File.Move(originalfilePath, newFilename)
                    Exit Do
                Catch e As IOException
                Catch e As UnauthorizedAccessException
                End Try
            End If
            i += 1
            newFilename = $"{directory}{Path.DirectorySeparatorChar}{filename}.bak{i}{extension}"
        Loop
    End Sub

    Private Function createNewFile(
                    originalfilePath As String,
                    append As Boolean,
                    increment As Boolean) As FileStream
        Dim directory = Path.GetDirectoryName(originalfilePath)
        Dim filename = Path.GetFileNameWithoutExtension(originalfilePath)
        Dim extension = Path.GetExtension(originalfilePath)

        IO.Directory.CreateDirectory(directory)

        Dim filePath = originalfilePath
        Dim i = 0
        Do
            Try
                Return New FileStream(filePath,
                                    If(append, FileMode.Append, FileMode.Create),
                                    FileAccess.ReadWrite,
                                    FileShare.Read)
            Catch e As UnauthorizedAccessException
            Catch e As IOException
            End Try

            If Not increment Then Throw New UnauthorizedAccessException("File already in use or access denied")

            ' now increment the filename until we find one we can use
            Do
                i += 1
                filePath = $"{directory}{Path.DirectorySeparatorChar}{filename}-{i}{extension}"
            Loop While File.Exists(filePath)
        Loop
    End Function
End Module

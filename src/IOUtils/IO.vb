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

Public NotInheritable Class IO
    Private Sub New()
    End Sub

    Public Shared Function CreateWriteableTextFile(filePath As String, Optional pOverwrite As Boolean = False, Optional pCreateBackup As Boolean = False, Optional pUnicode As Boolean = False, Optional pIncrementFilenameIfInUse As Boolean = False) As StreamWriter
        Dim directory = Path.GetDirectoryName(filePath)
        Dim filename = Path.GetFileName(filePath)
        Dim extension = Path.GetExtension(filePath)

        If pOverwrite And pCreateBackup Then
            If File.Exists(filePath) Then
                Dim newFilename = $"{directory}{Path.DirectorySeparatorChar}{filename}.bak.{extension}"
                If File.Exists(newFilename) Then
                    Dim i = 0
                    Do
                        i = i + 1
                        newFilename = $"{directory}{Path.DirectorySeparatorChar}{filename}.bak{i}.{extension}"
                    Loop Until Not File.Exists(newFilename)
                End If
                Try
                    File.Move(filePath, newFilename)
                Catch e As System.UnauthorizedAccessException
                    ' can't rename the existing file: this may be either because it's in use by another
                    ' process, or we don't have access permission

                    If Not pIncrementFilenameIfInUse Then Throw New UnauthorizedAccessException("File already in use or access denied")

                    ' now increment the pFilename to find one we can use
                    Dim i = 1
                    filePath = $"{directory}{filename}-{i}.{extension}"
                    Do While File.Exists(filePath)
                        i = i + 1
                    Loop
                End Try
            End If
        End If

        Return New StreamWriter(filePath, Not pOverwrite, If(pUnicode, System.Text.Encoding.UTF8, System.Text.Encoding.Default))
    End Function

End Class

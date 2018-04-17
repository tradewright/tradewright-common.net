#Region "License"

' The MIT License (MIT)
'
' Copyright (c) 2017-2018 Richard L King (TradeWright Software Systems)
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

Imports System.Collections.Generic

Friend NotInheritable Class LogManager
    Implements IDisposable

    ''
    ' Maintains information about a set of logging-related objects.
    '
    '@/

    '@================================================================================
    ' Interfaces
    '@================================================================================

    '@================================================================================
    ' Events
    '@================================================================================

    '@================================================================================
    ' Enums
    '@================================================================================

    '@================================================================================
    ' Types
    '@================================================================================

    '@================================================================================
    ' Constants
    '@================================================================================

    '@================================================================================
    ' Member variables
    '@================================================================================

    Private mLoggers As New Dictionary(Of String, Logger)
    Private mRootLogger As New Logger("")

    '@================================================================================
    ' Class Event Handlers
    '@================================================================================

    Friend Sub New()
        mRootLogger.LogToParent = False
    End Sub

    '@================================================================================
    ' XXXX Interface Members
    '@================================================================================

    '@================================================================================
    ' XXXX Event Handlers
    '@================================================================================

    '@================================================================================
    ' Properties
    '@================================================================================

    '@================================================================================
    ' Methods
    '@================================================================================

    Friend Sub Close()
        Dispose(True)
    End Sub

    Friend Function GetLogger(infoType As String) As Logger
        If infoType = "" Then Return mRootLogger

        Debug.Assert(infoType.Substring(0, 1) <> "$", "Infotypes starting with $ are reserved for system use")

        GetLogger = GetLoggerEx(infoType)
    End Function

    Friend Function GetLoggerEx(infoType As String) As Logger
        If infoType = "" Then Return mRootLogger

        SyncLock mLoggers
            If mLoggers.ContainsKey(infoType) Then Return mLoggers.Item(infoType)

            GetLoggerEx = New Logger(infoType)
            mLoggers.Add(infoType, GetLoggerEx)

            Dim templogger As Logger = GetLoggerEx

            infoType = removeElement(infoType)
            Do While infoType <> ""

                Dim parentLogger As Logger = Nothing

                If mLoggers.ContainsKey(infoType) Then
                    templogger.Parent = mLoggers.Item(infoType)
                    Exit Function
                End If

                parentLogger = New Logger(infoType)
                mLoggers.Add(infoType, parentLogger)

                templogger.Parent = parentLogger
                templogger = parentLogger

                infoType = removeElement(infoType)

            Loop

            templogger.Parent = mRootLogger
        End SyncLock
    End Function

    '@================================================================================
    ' Helper Functions
    '@================================================================================

    Private Function removeElement(infoType As String) As String
        Dim l = infoType.LastIndexOf(".")
        If l >= 0 Then
            removeElement = infoType.Substring(0, l)
        Else
            removeElement = ""
        End If
    End Function

    Private isDisposed As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Private Sub Dispose(disposing As Boolean)
        If Not Me.isDisposed Then
            If disposing Then
                SyncLock mLoggers
                    For Each logger As Logger In mLoggers.Values
                        logger.Dispose()
                    Next
                End SyncLock
            End If

            mRootLogger.Dispose()

            ' TODO: free your own state (unmanaged objects).
            ' TODO: set large fields to null.
        End If
        Me.isDisposed = True
    End Sub

#Region " IDisposable Support "
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class

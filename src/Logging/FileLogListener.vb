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

Imports System.IO

Imports TradeWright.Utilities.DataStorage
Imports TradeWright.Utilities.Time

''' <summary>
''' Writes log records to a specified file.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class FileLogListener

    '@================================================================================
    ' Interfaces
    '@================================================================================

    Implements ILogListener
    Implements IDisposable

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

    Private WithEvents mLogger As Logger

    Private mFilename As String

    Private mWriter As StreamWriter

    Private mFormatter As ILogFormatter

    Private mFinished As Boolean

    Private mSynchronized As Boolean

    Private ReadOnly mLock As Object = New Object

    '@================================================================================
    ' Class Event Handlers
    '@================================================================================

    Public Sub New(stream As FileStream,
                   Optional timestampFormat As Time.TimestampFormats = Time.TimestampFormats.DateAndTimeISO8601,
                   Optional formatter As ILogFormatter = Nothing,
                   Optional unicode As Boolean = False)
        mWriter = New StreamWriter(stream, If(unicode, System.Text.Encoding.UTF8, System.Text.Encoding.Default))
        setFormatter(formatter, timestampFormat)
    End Sub

    Public Sub New(fileInfo As FileInfo,
                   Optional timestampFormat As Time.TimestampFormats = Time.TimestampFormats.DateAndTimeISO8601,
                   Optional formatter As ILogFormatter = Nothing,
                   Optional append As Boolean = False,
                   Optional unicode As Boolean = False)
        IO.Directory.CreateDirectory(Path.GetDirectoryName(fileInfo.FullName))
        mWriter = New StreamWriter(fileInfo.Open(If(append, FileMode.Append, FileMode.Create)),
                                   If(unicode, System.Text.Encoding.UTF8, System.Text.Encoding.Default))
        setFormatter(formatter, timestampFormat)
    End Sub

    Public Sub New(writer As StreamWriter,
                   Optional timestampFormat As Time.TimestampFormats = Time.TimestampFormats.DateAndTimeISO8601,
                   Optional formatter As ILogFormatter = Nothing)
        mWriter = writer
        setFormatter(formatter, timestampFormat)
    End Sub

    Public Sub New(filename As String,
                   Optional timestampFormat As Time.TimestampFormats = Time.TimestampFormats.DateAndTimeISO8601,
                   Optional formatter As ILogFormatter = Nothing,
                   Optional append As Boolean = False,
                   Optional unicode As Boolean = False)
        IO.Directory.CreateDirectory(Path.GetDirectoryName(filename))
        mWriter = New StreamWriter(filename, append, If(unicode, System.Text.Encoding.UTF8, System.Text.Encoding.Default))
        setFormatter(formatter, timestampFormat)
    End Sub

    '@================================================================================
    ' ILogListener Interface Members
    '@================================================================================

    Public Sub Initialise(logger As Logger, synchronized As Boolean) Implements ILogListener.Initialise
        mLogger = logger
        mSynchronized = synchronized
    End Sub

    '@================================================================================
    ' mLogger Event Handlers
    '@================================================================================

    Private Sub mLogger_Finished(s As Object, e As EventArgs) Handles mLogger.Finished
        If mSynchronized Then
            SyncLock mLock
                Dispose()
            End SyncLock
        Else
            Dispose()
        End If
    End Sub

    Private Sub mLogger_LogRecord(s As Object, e As LogRecordEventArgs) Handles mLogger.LogRecord
        If mFinished Then Exit Sub
        If mSynchronized Then
            SyncLock mLock
                mWriter.WriteLine(mFormatter.FormatRecord(e.LogRecord))
                mWriter.Flush()
            End SyncLock
        Else
            mWriter.WriteLine(mFormatter.FormatRecord(e.LogRecord))
            mWriter.Flush()
        End If
    End Sub

    '@================================================================================
    ' Properties
    '@================================================================================

    '@================================================================================
    ' Methods
    '@================================================================================

    '@================================================================================
    ' Helper Functions
    '@================================================================================

    Private Sub setFormatter(formatter As ILogFormatter, timestampFormat As Time.TimestampFormats)
        If formatter Is Nothing Then
            mFormatter = New BasicLogFormatter(timestampFormat, False)
        Else
            mFormatter = formatter
        End If

        If mFormatter.Header <> "" Then mWriter.WriteLine(mFormatter.Header)
    End Sub

    Private isDisposed As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Private Sub Dispose(disposing As Boolean)
        If Not Me.isDisposed Then
            If disposing Then
                mFinished = True
                If mFormatter.Trailer <> "" Then mWriter.WriteLine(mFormatter.Trailer)
                mWriter.Flush()
                mWriter.Dispose()
            End If
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

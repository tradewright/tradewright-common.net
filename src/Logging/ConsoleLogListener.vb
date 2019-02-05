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

''' <summary>
''' Writes log records to the console.
''' </summary>
''' <remarks></remarks>
Public NotInheritable Class ConsoleLogListener
    Implements ILogListener

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

    Private WithEvents mLogger As Logger

    Private mFormatter As ILogFormatter

    Private mFinished As Boolean

    '@================================================================================
    ' Class Event Handlers
    '@================================================================================

    Public Sub New(Optional formatter As ILogFormatter = Nothing, Optional timestampFormat As Time.TimestampFormats = Time.TimestampFormats.DateAndTimeISO8601)
        If formatter Is Nothing Then
            mFormatter = New BasicLogFormatter(timestampFormat, False)
        Else
            mFormatter = formatter
        End If

        If mFormatter.Header <> "" Then Console.WriteLine(mFormatter.Header)
    End Sub

    '@================================================================================
    ' ILogListener Interface Members
    '@================================================================================

    Public Sub Initialise(logger As Logger, synchronized As Boolean) Implements ILogListener.Initialise
        ' note that only static members of Console are used, and they are all thread-safe
        mLogger = logger
    End Sub

    Private Sub mLogger_Finished(s As Object, e As System.EventArgs) Handles mLogger.Finished
        mFinished = True

        If mFormatter.Trailer <> "" Then Console.WriteLine(mFormatter.Trailer)
    End Sub

    Private Sub mLogger_LogRecord(s As Object, e As LogRecordEventArgs) Handles mLogger.LogRecord
        If mFinished Then Exit Sub
        Console.WriteLine(mFormatter.FormatRecord(e.LogRecord))
    End Sub

    '@================================================================================
    ' XXXX Event Handlers
    '@================================================================================

    '@================================================================================
    ' Properties
    '@================================================================================

    '@================================================================================
    ' Methods
    '@================================================================================

    '@================================================================================
    ' Helper Functions
    '@================================================================================

    Private IsDisposed As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Private Sub Dispose(disposing As Boolean)
        If Not Me.IsDisposed Then
            If disposing Then
                ' TODO: free other state (managed objects).
            End If

            ' TODO: free your own state (unmanaged objects).
            ' TODO: set large fields to null.
        End If
        Me.IsDisposed = True
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

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

Imports TradeWright.Utilities.Time

Public NotInheritable Class BasicLogFormatter

    ''
    ' A <c>LogFormatter</c> that provides a simple text conversion of a log record.
    '
    '@/

    '@================================================================================
    ' Interfaces
    '@================================================================================

    Implements ILogFormatter

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


    Private Const ModuleName As String = "BasicLogFormatter"

    '@================================================================================
    ' Member variables
    '@================================================================================

    Private mTimestampFormat As Time.TimestampFormats
    Private mIncludeInfoType As Boolean
    Private mIncludeTimestamp As Boolean
    Private mIncludeLogLevel As Boolean

    '@================================================================================
    ' Class Event Handlers
    '@================================================================================

    Public Sub New(Optional timestampFormat As TimestampFormats = TimestampFormats.DateAndTimeISO8601,
                   Optional includeInfoType As Boolean = False,
                   Optional includeTimestamp As Boolean = True,
                   Optional includeLogLevel As Boolean = True)
        mTimestampFormat = timestampFormat
        mIncludeInfoType = includeInfoType
        mIncludeTimestamp = includeTimestamp
        mIncludeLogLevel = includeLogLevel
    End Sub

    '@================================================================================
    ' LogFormatter Interface Members
    '@================================================================================

    Public Function FormatRecord(Logrec As LogRecord) As String Implements ILogFormatter.FormatRecord
        Static sSpacer As String = String.Empty

        Dim timestamp = If(mIncludeTimestamp, Time.Time.FormatTimestamp(Logrec.Timestamp, mTimestampFormat) & " ", "")
        Dim logLevel = If(mIncludeLogLevel, Logging.LogLevelToShortString(Logrec.LogLevel) & " ", "")
        Dim infoType = If(mIncludeInfoType, Logrec.InfoType & ": ", "")

        If String.IsNullOrEmpty(sSpacer) Then sSpacer = New String(" "c, timestamp.Length + logLevel.Length)

        Dim data = Logrec.Data.ToString & formatData(Logrec.Data)

        If data.IndexOf(System.Environment.NewLine, 0) <> -1 Then
            data = data.Replace(System.Environment.NewLine, System.Environment.NewLine & sSpacer)
        End If

        Return timestamp & logLevel & infoType & data

    End Function

    Public ReadOnly Property Header() As String Implements ILogFormatter.Header
        Get
            Return String.Empty
        End Get
    End Property

    Public ReadOnly Property Trailer() As String Implements ILogFormatter.Trailer
        Get
            Return String.Empty
        End Get
    End Property

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

    Private Function formatData(data As Object) As String
        If Not data.GetType.IsArray Then Return String.Empty

        Dim ar = CType(data, Array)

        If ar.Length = 0 Then Return "{}"

        Dim s(Math.Min(ar.Length - 1, 5)) As String

        Dim i = 0
        For i = 0 To ar.Length - 1
            s(i) = ar.GetValue(i).ToString
        Next

        s(0) = "{" & s(0)
        If i > 5 Then
            s(5) = "...}"
        Else
            s(i - 1) &= "}"
        End If

        Return String.Join(", ", s)
    End Function

End Class

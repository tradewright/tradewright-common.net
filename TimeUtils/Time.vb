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


''' <summary>
''' Timestamp format specifiers for use with the <c>FormatTimestamp</c>
''' global method.
''' </summary>
Public Enum TimestampFormats
    ''' <summary>
    ''' The formatted Timestamp is of the form <em>hhmmss.lll</em> where <em>lll</em>
    ''' is milliseconds.
    ''' </summary>
    TimeOnly = 0

    ''' <summary>
    '''  The formatted Timestamp is of the form <em>yyyymmdd</em>.
    ''' </summary>
    DateOnly = 1

    ''' <summary>
    ''' The formatted Timestamp is of the form <em>yyyymmddhhmmss.lll</em> where <em>lll</em>
    ''' is milliseconds.
    ''' </summary>
    DateAndTime = 2

    ''' <summary>
    ''' The formatted Timestamp is of the form <em>hh:mm:ss.lll</em> where <em>lll</em>
    ''' is milliseconds.
    ''' </summary>
    TimeOnlyISO8601 = 4

    ''' <summary>
    ''' The formatted Timestamp is of the form <em>yyyy-mm-dd</em>.
    ''' </summary>
    DateOnlyISO8601 = 5

    ''' <summary>
    ''' The formatted Timestamp is of the form <em>yyyy-mm-dd hh:mm:ss.lll</em> where <em>lll</em>
    ''' is milliseconds.
    ''' </summary>
    DateAndTimeISO8601 = 6

    ''' <summary>
    ''' The formatted Timestamp is of the form <em>&lt;time&gt;.lll</em>, where &lt;time&gt; is
    ''' the time in the format defined in the Control Panel's Regional and Language Options,
    ''' and <em>lll</em> is milliseconds.
    ''' </summary>
    TimeOnlyLocal = 7

    ''' <summary>
    ''' The formatted Timestamp is of the form <em>&lt;date&gt;.lll</em>, where &lt;date&gt; is
    ''' the date in the format defined in the Control Panel's Regional and Language Options,
    ''' and <em>lll</em> is milliseconds.
    ''' </summary>
    DateOnlyLocal = 8

    ''' <summary>
    ''' The formatted Timestamp is of the form <em>&lt;date &amp; time&gt;.lll</em>, where
    ''' &lt;date &amp; time&gt; is the date and time in the format defined in the Control Panel's
    ''' Regional and Language Options, and <em>lll</em> is milliseconds.
    ''' </summary>
    DateAndTimeLocal = 9

    ''' <summary>
    ''' Used additively in conjunction with other values, prevents the milliseconds
    ''' value being included in the Timestamp.
    ''' </summary>
    NoMillisecs = &H40000000

    ''' <summary>
    ''' Used additively in conjunction with other values, results in a timezone specifier
    ''' being appended. The format of the timezone specifier is +|-hh:mm.
    ''' </summary>
    IncludeTimezone = &H80000000
End Enum

Public NotInheritable Class Time
    Private Sub New()
    End Sub

    ''' <summary>
    ''' Formats the DateTime supplied in <paramref name="timestamp"/> according to the value specified in <paramref name="formatOption"/>.
    ''' </summary>
    ''' <param name="timestamp"></param>
    ''' <param name="formatOption"></param>
    ''' <remarks>This is a convenience method that removes the need to remember or lookup the 
    ''' various formatting options in the DateTime.ToString() method.</remarks>
    ''' <returns></returns>
    Public Shared Function FormatTimestamp(
                        timestamp As DateTime,
                        Optional formatOption As TimestampFormats = TimestampFormats.DateAndTime) As String

        Dim includeTimezone = (formatOption And TimestampFormats.IncludeTimezone) <> 0
        formatOption = formatOption And (Not TimestampFormats.IncludeTimezone)

        Dim noMillisecs = CBool(formatOption And TimestampFormats.NoMillisecs)
        formatOption = formatOption And (Not TimestampFormats.NoMillisecs)

        Dim milliseconds = timestamp.Millisecond

        Select Case formatOption
            Case TimestampFormats.TimeOnly
                If noMillisecs Then Return timestamp.ToString("HHmmss")
                Return timestamp.ToString("HHmmss.fff")
            Case TimestampFormats.DateOnly
                Return timestamp.ToString("yyyyMMdd")
            Case TimestampFormats.DateAndTime
                If noMillisecs Then Return timestamp.ToString("yyyyMMddHHmmss")
                Return timestamp.ToString("yyyyMMddHHmmss.fff")
            Case TimestampFormats.TimeOnlyISO8601
                If noMillisecs Then Return timestamp.ToString("HH:mm:ss")
                Return timestamp.ToString("HH:mm:ss.fff")
            Case TimestampFormats.DateOnlyISO8601
                Return timestamp.ToString("yyyy-MM-dd")
            Case TimestampFormats.DateAndTimeISO8601
                If noMillisecs Then Return timestamp.ToString("yyyy-MM-dd HH:mm:ss")
                Return timestamp.ToString("yyyy-MM-dd HH:mm:ss.fff")
            Case TimestampFormats.TimeOnlyLocal
                Return timestamp.ToLongTimeString()
            Case TimestampFormats.DateOnlyLocal
                Return timestamp.ToShortDateString()
            Case TimestampFormats.DateAndTimeLocal
                Return timestamp.ToShortDateString() & " " & timestamp.ToLongTimeString()
            Case Else
                Throw New ArgumentException("Invalid format option")
        End Select

        If includeTimezone Then Throw New NotImplementedException
    End Function

End Class



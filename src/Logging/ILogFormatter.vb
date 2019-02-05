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
''' This interface is implemented by classes that provide services for formatting
''' <c>LogRecord</c>s.
''' </summary>
Public Interface ILogFormatter

    '@/

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
    ' Properties
    '@================================================================================

    ''
    ' Returns any Header text that should precede the formatted log records.
    '
    ' @return
    '   Any Header text that should precede the formatted log records.
    '@/
    ReadOnly Property Header() As String

    ''
    ' Returns any Trailer text that should follow the formatted log records.
    '
    ' @return
    '   Any Trailer text that should follow the formatted log records.
    '@/
    ReadOnly Property Trailer() As String

    '@================================================================================
    ' Methods
    '@================================================================================

    ''
    ' Returns a string resulting from formatting the contents of a
    ' <c>LogRecord</c> object.
    '
    ' @return
    '   The formatted log record.
    ' @param Logrec
    '   The <c>LogRecord</c> object to be formatted.
    ' @see
    '
    '@/
    Function FormatRecord(Logrec As LogRecord) As String

End Interface

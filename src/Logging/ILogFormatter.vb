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

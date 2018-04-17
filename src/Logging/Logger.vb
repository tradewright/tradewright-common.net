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

''
' Used to control the level of detail or importance of information that is
' logged.
'
' @param AsParent
'   The <c>Logger</c> object will obtain its log level from its Parent. This Value
'   must only be used to set the <c>LogLevel</c> property of <c>Logger</c>
'   objects. It must not be used in logging methods.
' @param UseDefault
'   The <c>Logger</c> object will use the Value of the global <c>DefaultLogLevel</c>
'   property as its log level. This Value must only be used to set the <c>LogLevel</c>
'   property of <c>Logger</c> objects. It must not be used in logging methods.
' @param All
'   The <c>Logger</c> object will log all information. This Value
'   must only be used to set the <c>LogLevel</c> property of <c>Logger</c>
'   objects. It must not be used in logging methods.
' @param HighDetail
'   Use this level for very detailed tracing or diagnostic information. Typically this is
'   information that is only of interest to specialist support staff or developers, or is
'   very voluminous.
' @param MediumDetail
'   Use this level for fairly detailed tracing or diagnostic information. Typically this is
'   information that is only of interest to specialist support staff or developers, and is
'   not very voluminous.
' @param Detail
'   Use this level for moderately detailed tracing or diagnostic information. Typically this is
'   information that is only of interest to support staff or developers, and is
'   not very voluminous.
' @param Normal
'   Use this level for general information that may be of interest to end users
'   or administrators, and is not very voluminous, but that does not generally
'   need to be brought to the user's attention.
' @param Info
'   Use this level for information that may be of interest to end users or administrators,
'   and that should be brought to the user's attention.
' @param Warning
'   Use this level for information that may indicate a potential problem or non-critical failure.
' @param Severe
'   Use this level for information regarding a critical failure.
' @param None
'   The <c>Logger</c> object will log no information. This Value
'   must only be used to set the <c>LogLevel</c> property of <c>Logger</c>
'   objects. It must not be used in logging methods.
'@/
Public Enum LogLevel
    All = &H80000000
    HighDetail = -3000
    MediumDetail = -2000
    Detail = -1000
    Normal = 0
    Info = 1000
    Warning = 2000
    Severe = 3000
    UseDefault = &H7FFFFFFD
    None = &H7FFFFFFE
    AsParent = &H7FFFFFFF
End Enum

Public NotInheritable Class Logger
    Implements IDisposable
    Implements ILogger

    ''
    ' Objects of this class are used to log information of a particular type
    ' (commonly known as an &quot;infotype&quot;).
    '
    ' Information type names form a hierarchical namespace, components of a name
    ' being separated by a period Character (ie &quot;.&quot;). The root of the tree is
    ' an empty string. Information type names are not case sensitive.
    '
    ' Infotype names starting with a &quot;$&quot; Character are reserved for use within
    ' the system. Applications are not able to obtain direct access to the <c>Logger</c>
    ' object for such infotypes.
    '
    ' There is a single <c>Logger</c> object for each information type, which
    ' may be obtained by calling the global <c>GetLogger</c> method.
    '
    ' Information logged by a <c>Logger</c> object is, by default, also logged by the
    ' <c>Logger</c> object for the Parent information type, recursively up the
    ' information type namespace tree.
    '
    ' Each <c>Logger</c> object has a <c>LogLevel</c> property which is a
    ' Value within the <c>LogLevels</c> enum. If it is set to the Value
    ' <c>AsParent</c>, then the <c>Logger</c> object obtains its Value
    ' from its Parent <c>Logger</c> object. A Value of <c>AsParent</c>
    ' for the root <c>Logger</c> object has the same effect as <c>None</c>,
    ' ie it logs nothing.
    '
    ' When a <c>Logger</c> object is created, its initial log level is set to
    ' the current default log level, as set by the global <c>DefaultLogLevel</c> property.
    '
    '@/

    '@================================================================================
    ' Interfaces
    '@================================================================================

    '@================================================================================
    ' Events
    '@================================================================================

    Public Event Finished(sender As Object, e As EventArgs) Implements ILogger.Finished
    Public Event LogRecord(sender As Object, e As LogRecordEventArgs) Implements ILogger.LogRecord

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

    Private mLogLevel As LogLevel = LogLevel.UseDefault

    Private mParent As Logger
    Private mLogToParent As Boolean = True

    Private mInfoType As String

    '@================================================================================
    ' Class Event Handlers
    '@================================================================================

    Private Sub New()
    End Sub

    Friend Sub New(infoType As String)
        mInfoType = infoType
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

    ''
    ' Gets or set the logging level for this <c>Logger</c> object.
    '
    ' @param Value
    '   The logging level used by this <c>Logger</c> object in deciding
    '   whether to log.
    '@/
    Public Property LogLevel() As LogLevel Implements ILogger.LogLevel
        Get
            If mLogLevel = LogLevel.AsParent Then
                If mParent Is Nothing Then
                    LogLevel = LogLevel.None
                Else
                    LogLevel = mParent.LogLevel
                End If
            ElseIf mLogLevel = LogLevel.UseDefault Then
                LogLevel = Logging.DefaultLogLevel
            Else
                LogLevel = mLogLevel
            End If
        End Get
        Set(Value As LogLevel)
            mLogLevel = Value
        End Set
    End Property

    ''' <summary>
    ''' Specifies whether this <c>Logger</c> object is to log to its Parent
    ''' <c>Logger</c> object.
    ''' </summary>
    ''' <returns>If <c>True</c>, this <c>Logger</c> object logs to its Parent
    '''   <c>Logger</c> object.
    ''' </returns>
    Public Property LogToParent() As Boolean Implements ILogger.LogToParent
        Get
            LogToParent = mLogToParent
        End Get
        Set(Value As Boolean)
            mLogToParent = Value
        End Set
    End Property

    Friend WriteOnly Property Parent() As Logger
        Set(Value As Logger)
            mParent = Value
        End Set
    End Property

    '@================================================================================
    ' Methods
    '@================================================================================

    ''
    ' Indicates whether a message of a specified log level would be logged.
    '
    ' @return
    '   Returns <c>True</c> if the message would be logged, and <c>False</c>
    '   otherwise.
    ' @param level
    '   The relevant log level.
    ' @see
    '
    '@/
    Public Function IsLoggable(level As LogLevel) As Boolean Implements ILogger.IsLoggable
        If Not Logging.IsLogLevelPermittedForApplication(level) Then Return False

        Return (level >= LogLevel)
    End Function

    ''
    ' Logs Data of a specified log level.
    '
    ' @remarks
    '   The Data is logged only if the specified log level is not less than the
    '   current <c>LogLevel</c> property.
    '
    '   If the Data is logged and the <c>logToParent</c> property is <c>True</c>,
    '   then the Data is also logged by the Parent logger.
    '
    '   If the Data is of a type that has no string representation, then it may be ignored
    '   by some <c>LogListener</c> objects. To ensure that this does not happen:
    '   <ul>
    '   <li>Avoid logging User Defined Types, unless you can be sure that listeners are
    '   prepared to handle them.</li>
    '   <li>Only log objects that implement the <c>Stringable</c> interface.</li>
    '   </ul>
    ' @param level
    '   The log level of the specified Data.
    ' @param Data
    '   The Data to be logged. Note that this is a variant, so any Data type
    '   that can be held in a variant can be logged (but see the Remarks concerning UDTs and
    '   objects).
    ' @param Source
    '   Information that identifies the Source of this Log Record (this information
    '   could for example be a reference to an object, or something that uniquely identifies
    '   an object). It need not be supplied where there is no need to distinguish between
    '   log records from different sources.
    '@/
    Public Sub Log(pLevel As LogLevel, pData As Object, Optional pSource As Object = Nothing) Implements ILogger.Log
        Try ' to avoid problems with errors here being caught by
            ' lower stack frames which may attempt to log them!

            If pData Is Nothing Then Exit Sub
            If Not IsLoggable(pLevel) Then Exit Sub

            Logrec(New LogRecord(pLevel, mInfoType, Logging.GetNextLoggingSequenceNum, pData, pSource))
        Catch e As Exception
        End Try
    End Sub

    Friend Sub Logrec(rec As LogRecord)
        If IsLoggable(rec.LogLevel) Then
            RaiseEvent LogRecord(Me, New LogRecordEventArgs(rec))
            If LogToParent Then mParent.Logrec(rec)
        End If
    End Sub

    '@================================================================================
    ' Helper Functions
    '@================================================================================

    Private isDisposed As Boolean = False        ' To detect redundant calls

    ' IDisposable
    Private Sub Dispose(disposing As Boolean)
        If Not Me.isDisposed Then
            If disposing Then
                RaiseEvent Finished(Me, EventArgs.Empty)
            End If

            ' TODO: free your own state (unmanaged objects).
            ' TODO: set large fields to null.
        End If
        Me.isDisposed = True
    End Sub

#Region " IDisposable Support "
    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose, ILogger.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

End Class


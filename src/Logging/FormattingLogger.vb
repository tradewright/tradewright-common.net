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

Imports System.Diagnostics
Imports System.Reflection

''' <summary>
''' Provides an easy means of logging information with
''' a consistent format.
''' </summary>
''' <remarks>
''' <para>Log record data is formatted as follows:</para>
''' <para>[projectname.modulename:procedurename] message: message-qualifier</para>
''' <para>where projectname is supplied when the <c>FormattingLogger</c> object
''' is created and modulename, procedurename, message and message-qualifier are
''' supplied in the call to the <c>Log</c> method.</para>
'''
''' </remarks>
Public NotInheritable Class FormattingLogger
    Implements IFormattingLogger

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

    Private mLogger As Logger

    Private mProjectName As String

    '@================================================================================
    ' Class Event Handlers
    '@================================================================================

    Private Sub New()
    End Sub

    ''' <summary>
    ''' Creates a new instance of the <see cref="FormattingLogger"></see> class.
    ''' </summary>
    ''' <param name="infoType">The info type to be logged by this <see cref="FormattingLogger"></see>.</param>
    ''' <remarks></remarks>
    Public Sub New(infoType As String)
        Init(infoType, "")
    End Sub

    ''' <summary>
    ''' Creates a new instance of the <see cref="FormattingLogger"></see> class.
    ''' </summary>
    ''' <param name="infoType">The info type to be logged by this <see cref="FormattingLogger"></see>.</param>
    ''' <param name="projectName">A name used in the log to identify events originating within the current project.</param>
    ''' <remarks></remarks>
    Public Sub New(infoType As String, projectName As String)
        Init(infoType, projectName)
    End Sub

    Private Sub Init(infoType As String, projectName As String)
        mLogger = Logging.GetLogger(infoType)
        If String.IsNullOrEmpty(projectName) Then projectName = New StackFrame(1).GetMethod.Module.Name
        mProjectName = projectName
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
    ' Gets or set the logging level for this <c>FormattingLogger</c> object.
    '
    ' @param Value
    '   The logging level used by this <c>FormattingLogger</c> object in deciding
    '   whether to log.
    '@/
    Public Property LogLevel() As LogLevel Implements IFormattingLogger.LogLevel
        Get
            Return mLogger.LogLevel
        End Get
        Set(Value As LogLevel)
            mLogger.LogLevel = Value
        End Set
    End Property

    ''' <summary>
    ''' Specifies whether this <c>Logger</c> object is to log to its Parent
    ''' <c>Logger</c> object.
    ''' </summary>
    ''' <returns>If <c>True</c>, this <c>Logger</c> object logs to its Parent
    '''   <c>Logger</c> object.
    ''' </returns>
    Public Property LogToParent() As Boolean Implements IFormattingLogger.LogToParent
        Get
            LogToParent = mLogger.LogToParent
        End Get
        Set(Value As Boolean)
            mLogger.LogToParent = Value
        End Set
    End Property

    '@================================================================================
    ' Methods
    '@================================================================================

    ''' <summary>
    ''' Indicates whether a message of a specified log level would be logged.
    ''' </summary>
    ''' <param name="level">The relevant log level.</param>
    ''' <returns>Returns <c>True</c> if the message would be logged, and <c>False</c>
    '''  otherwise.</returns>
    ''' <remarks></remarks>
    Public Function IsLoggable(level As LogLevel) As Boolean Implements IFormattingLogger.IsLoggable
        IsLoggable = mLogger.IsLoggable(level)
    End Function

    ''' <summary>
    ''' Logs string data.
    ''' </summary>
    ''' <param name="msg">The data to be logged.</param>
    ''' <param name="logLevel">The log level of the specified Data.</param>
    ''' <remarks> The data is logged only if the specified log level is not less than the
    ''' current <c>LogLevel</c> property of the <see cref="Logger"></see>
    ''' object for the infotype handled by this <see cref="FormattingLogger"></see>.
    '''  
    ''' If the data is logged and the <c>logToParent</c> property of the <see cref="Logger"></see>
    ''' object for the infotype handled by this <see cref="FormattingLogger"></see> is <c>True</c>,
    ''' then the data is also logged by the parent logger.
    ''' </remarks>
    Public Sub Log(msg As String, Optional logLevel As LogLevel = LogLevel.Normal) Implements IFormattingLogger.Log
        LogIt(msg, "", "", logLevel)
    End Sub

    ''' <summary>
    ''' Logs string data.
    ''' </summary>
    ''' <param name="msg">The data to be logged.</param>
    ''' <param name="moduleName">The name of the code module where this event was logged.</param>
    ''' <param name="procName">The name of the procedure where this event was logged.</param>
    ''' <param name="logLevel">The log level of the specified Data.</param>
    ''' <remarks> The data is logged only if the specified log level is not less than the
    ''' current <c>LogLevel</c> property of the <see cref="Logger"></see>
    ''' object for the infotype handled by this <see cref="FormattingLogger"></see>.
    '''  
    ''' If the data is logged and the <c>logToParent</c> property of the <see cref="Logger"></see>
    ''' object for the infotype handled by this <see cref="FormattingLogger"></see> is <c>True</c>,
    ''' then the data is also logged by the parent logger.
    ''' </remarks>
    Public Sub Log(msg As String, moduleName As String, procName As String, Optional logLevel As LogLevel = LogLevel.Normal) Implements IFormattingLogger.Log
        LogIt(msg, moduleName, procName, logLevel)
    End Sub

    '@================================================================================
    ' Helper Functions
    '@================================================================================

    Private Sub LogIt(msg As String, moduleName As String, procName As String, logLevel As LogLevel)
        If Not mLogger.IsLoggable(logLevel) Then Exit Sub

        If String.IsNullOrEmpty(moduleName) Or String.IsNullOrEmpty(procName) Then
            Dim methodBase = New StackFrame(2).GetMethod
            If String.IsNullOrEmpty(moduleName) Then moduleName = methodBase.ReflectedType.Name
            If String.IsNullOrEmpty(procName) Then procName = methodBase.Name
        End If

        mLogger.Log(logLevel, $"[{mProjectName}.{moduleName}:{procName}] {msg}")
    End Sub

End Class


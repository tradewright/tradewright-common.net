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

Imports System.Reflection
Imports System.Windows.Forms

Public NotInheritable Class Logging

    ' command line switch specifying the log filename
    Private Const SwitchLogFilename As String = "log"

    ' command line switch specifying the log level
    Private Const SwitchLogLevel As String = "loglevel"

    Private Shared gLogManager As New LogManager
    Private Shared gDefaultLogLevel As LogLevel = LogLevel.Normal
    Private Shared gSequenceNumber As Integer = 0
    Private Shared gLogFileName As String = String.Empty
    Private Shared gDefaultLogger As New FormattingLogger("")

    Private Shared gLogListeners As New HashSet(Of ILogListener)

    Private Sub New()
    End Sub

    ''' <summary>
    ''' Returns the path and filename of the log file created by a call to the
    ''' <c>SetupDefaultLogging</c> method.
    ''' </summary>
    ''' <returns>The path and filename of the log file.</returns>
    Public Shared ReadOnly Property DefaultLogFileName() As String
        Get
            If gLogFileName = "" Then
                Dim clp = New CommandLine.CommandLineParser(Environment.CommandLine, " ")
                If clp.IsSwitchSet(SwitchLogFilename) Then gLogFileName = clp.SwitchValue(SwitchLogFilename)

                If gLogFileName = "" Then
                    gLogFileName = Application.LocalUserAppDataPath & "\" & "Logfile.log"
                End If
            End If
            Return gLogFileName
        End Get
    End Property

    Public Shared ReadOnly Property DefaultLogger() As FormattingLogger
        Get
            Return gDefaultLogger
        End Get
    End Property

    ''' <summary>
    ''' Sets or gets the default log level for <c>Logger</c> objects whose <c>LogLevel</c>
    ''' is not explicitly set.
    ''' </summary>
    ''' <returns>The default <c>LogLevel</c></returns>
    ''' <remarks> If this property is not set by the application, a Value of <c>Normal</c>
    ''' is assumed.</remarks>
    Public Shared Property DefaultLogLevel() As LogLevel
        Get
            Return gDefaultLogLevel
        End Get
        Set(value As LogLevel)
            If value = LogLevel.All Then
                ' this is allowed
            Else
                Debug.Assert(IsLogLevelPermittedForApplication(value), "This Value is not permitted in this context")
            End If

            gDefaultLogLevel = value
        End Set
    End Property

    Public Shared Sub AddLogListener(infoType As String, listener As ILogListener, Optional synchronized As Boolean = False)
        SyncLock gLogListeners
            If gLogListeners.Contains(listener) Then Throw New InvalidOperationException("This listener has already been added")
            gLogListeners.Add(listener)
            listener.Initialise(GetLogger(infoType), synchronized)
        End SyncLock
    End Sub

    Public Shared Sub Close()
        gLogManager.Close()
        SyncLock gLogListeners
            For Each logListener In gLogListeners
                logListener.Dispose()
            Next
            gLogListeners.Clear()
        End SyncLock
    End Sub

    Public Shared Function GetLogger(infoType As String) As Logger
        Return gLogManager.GetLogger(infoType)
    End Function

    Friend Shared Function GetNextLoggingSequenceNum() As Integer
        Return System.Threading.Interlocked.Increment(gSequenceNumber)
    End Function

    Public Shared Function IsLogLevelPermittedForApplication(value As LogLevel) As Boolean
        Select Case value
            Case LogLevel.All,
                    LogLevel.UseDefault,
                    LogLevel.None,
                    LogLevel.AsParent
                Return False
            Case Else
                Return True
        End Select
    End Function

    Public Shared Function LogLevelFromString(Value As String) As LogLevel
        Select Case Value.ToUpper
            Case "A", "ALL"
                Return LogLevel.All
            Case "D", "DETAIL"
                Return LogLevel.Detail
            Case "H", "HIGH", "HIGH DETAIL", "HIGHDETAIL"
                Return LogLevel.HighDetail
            Case "I", "INFO"
                Return LogLevel.Info
            Case "M", "MEDIUM", "MEDIUMDETAIL", "MEDIUM DETAIL"
                Return LogLevel.MediumDetail
            Case "0", "NONE"
                Return LogLevel.None
            Case "N", "NORMAL"
                Return LogLevel.Normal
            Case "-", "NULL"
                Return LogLevel.AsParent
            Case "S", "SEVERE"
                Return LogLevel.Severe
            Case "W", "WARNING"
                Return LogLevel.Warning
            Case "U", "DEFAULT"
                Return LogLevel.UseDefault
            Case Else
                Throw New ArgumentException("Invalid log level name")
        End Select
    End Function

    Public Shared Function LogLevelToShortString(Value As LogLevel) As String
        Select Case Value
            Case LogLevel.All
                Return "A "
            Case LogLevel.Detail
                Return "D "
            Case LogLevel.HighDetail
                Return "H "
            Case LogLevel.Info
                Return "I "
            Case LogLevel.MediumDetail
                Return "M "
            Case LogLevel.None
                Return "0 "
            Case LogLevel.Normal
                Return "N "
            Case LogLevel.AsParent
                Return "- "
            Case LogLevel.Severe
                Return "S "
            Case LogLevel.Warning
                Return "W "
            Case LogLevel.UseDefault
                Return "U "
            Case Else
                Return "? "
        End Select
    End Function

    Public Shared Function LogLevelToString(Value As LogLevel) As String
        Select Case Value
            Case LogLevel.All
                Return "All"
            Case LogLevel.Detail
                Return "Detail"
            Case LogLevel.HighDetail
                Return "High detail"
            Case LogLevel.Info
                Return "Info"
            Case LogLevel.MediumDetail
                Return "Medium detail"
            Case LogLevel.None
                Return "None"
            Case LogLevel.Normal
                Return "Normal"
            Case LogLevel.AsParent
                Return "Null"
            Case LogLevel.Severe
                Return "Severe"
            Case LogLevel.Warning
                Return "Warning"
            Case LogLevel.UseDefault
                Return "Default"
            Case Else
                Return CStr(Value)
        End Select
    End Function

    Public Shared Sub RemoveLogListener(listener As ILogListener)
        SyncLock gLogListeners
            gLogListeners.Remove(listener)
        End SyncLock
        listener.Dispose()
    End Sub

    ''' <summary>
    '''  Sets up logging to write all logged events for all infotypes to a default log file.
    ''' </summary>
    ''' <param name="synchronized">Set to true if multiple threads may use logging simultaneously.</param>
    ''' <param name="overwriteExisting">
    ''' If <c>True</c>, an existing log file of the same name is overwritten. Otherwise
    ''' new log entries are appended to an existing log file of the same name (if non exists,
    ''' a new log file is created.
    ''' </param>
    ''' <param name="createBackup">
    ''' If <c>True</c>, a backup copy of an existing log file of the same name is created.
    ''' </param>
    ''' <returns>The path and filename of the log file.</returns>
    ''' <remarks>
    ''' The logging level is governed by the global <see cref="DefaultLogLevel"/> property,
    ''' but can be overridden if the application's command line specifies the <c>/loglevel</c>
    ''' switch, which can take any of the following values:
    ''' <ul>
    '''     <li>None    or 0</li>
    '''     <li>Severe  or S</li>
    '''     <li>Warning or W</li>
    '''     <li>Info    or I</li>
    '''     <li>Normal  or N</li>
    '''     <li>Detail  or D</li>
    '''     <li>Medium  or M</li>
    '''     <li>High    or H</li>
    '''     <li>All     or A</li>
    ''' </ul>
    '''
    ''' If the application's command line includes the &quot;/log&quot; switch, then
    ''' the switch's value specifies the log file path and filename.
    '''
    ''' Otherwise, the log file is called <c>.log</c>, and is created in the application's 
    ''' settings folder as determined by the <see cref="Application.LocalUserAppDataPath"/> property).
    ''' </remarks>
    ''' 
    Public Shared Function SetupDefaultLogging(Optional synchronized As Boolean = False, Optional overwriteExisting As Boolean = True, Optional createBackup As Boolean = False) As String
        Static called As Boolean
        If called Then Throw New InvalidOperationException("SetupDefaultLogging has already been called")
        called = True

        Dim clp = New CommandLine.CommandLineParser(Environment.CommandLine, " ")
        If clp.IsSwitchSet(SwitchLogLevel) Then
            DefaultLogLevel = LogLevelFromString(clp.SwitchValue(SwitchLogLevel))
        ElseIf DefaultLogLevel <> LogLevel.Normal Then
        Else
            DefaultLogLevel = LogLevel.Normal
        End If

        AddLogListener("", New FileLogListener(DefaultLogFileName,
                                                  New BasicLogFormatter(),
                                                  overwriteExisting,
                                                  createBackup,
                                                  True), synchronized)

        Dim ass = Assembly.GetEntryAssembly
        If ass IsNot Nothing Then DefaultLogger.Log(ass.FullName)
        DefaultLogger.Log(Assembly.GetExecutingAssembly.FullName)
        DefaultLogger.Log("Log level: " & LogLevelToString(DefaultLogLevel))
        DefaultLogger.Log("Log file:  " & DefaultLogFileName)
        Return DefaultLogFileName
    End Function

End Class

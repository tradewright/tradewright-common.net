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
Imports System.Reflection

Imports TradeWright.Utilities.DataStorage

Public Module Logging

    ' command line switch specifying the log filename
    Private Const SwitchLogFilename As String = "log"

    ' command line switch specifying the log level
    Private Const SwitchLogLevel As String = "loglevel"

    Private ReadOnly mLogManager As New LogManager
    Private mDefaultLogLevel As LogLevel = LogLevel.Normal
    Private mSequenceNumber As Integer = 0
    Private mLogFileName As String = String.Empty

    Private mLogListeners As New HashSet(Of ILogListener)

    ''' <summary>
    ''' Returns the path and filename of the log file created by a call to the
    ''' <c>SetupDefaultLogging</c> method.
    ''' </summary>
    ''' <returns>The path and filename of the log file.</returns>
    Public ReadOnly Property DefaultLogFileName() As String
        Get
            If mLogFileName = "" Then
                Dim clp = New CommandLine.CommandLineParser(Environment.CommandLine, " ")
                If clp.IsSwitchSet(SwitchLogFilename) Then mLogFileName = clp.SwitchValue(SwitchLogFilename)

                If mLogFileName = "" Then
                    mLogFileName = $"{ApplicationDataPathUser}{Path.DirectorySeparatorChar}Logfile.log"
                End If
            End If
            Return mLogFileName
        End Get
    End Property

    Public ReadOnly Property DefaultLogger() As FormattingLogger = New FormattingLogger("")

    ''' <summary>
    ''' Sets or gets the default log level for <c>Logger</c> objects whose <c>LogLevel</c>
    ''' is not explicitly set.
    ''' </summary>
    ''' <returns>The default <c>LogLevel</c></returns>
    ''' <remarks> If this property is not set by the application, a Value of <c>Normal</c>
    ''' is assumed.</remarks>
    Public Property DefaultLogLevel() As LogLevel
        Get
            Return mDefaultLogLevel
        End Get
        Set(value As LogLevel)
            If value = LogLevel.All Then
                ' this is allowed
            Else
                Debug.Assert(IsLogLevelPermittedForApplication(value), "This value is not permitted in this context")
            End If

            mDefaultLogLevel = value
        End Set
    End Property

    Public Sub AddLogListener(infoType As String, listener As ILogListener, Optional synchronized As Boolean = False)
        SyncLock mLogListeners
            If mLogListeners.Contains(listener) Then Throw New InvalidOperationException("This listener has already been added")
            mLogListeners.Add(listener)
            listener.Initialise(GetLogger(infoType), synchronized)
        End SyncLock
    End Sub

    Public Sub Close()
        mLogManager.Close()
        SyncLock mLogListeners
            For Each logListener In mLogListeners
                logListener.Dispose()
            Next
            mLogListeners.Clear()
        End SyncLock
    End Sub

    Public Function GetLogger(infoType As String) As Logger
        Return mLogManager.GetLogger(infoType)
    End Function

    Friend Function GetNextLoggingSequenceNum() As Integer
        Return System.Threading.Interlocked.Increment(mSequenceNumber)
    End Function

    Public Function IsLogLevelPermittedForApplication(value As LogLevel) As Boolean
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

    Public Function LogLevelFromString(Value As String) As LogLevel
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

    Public Function LogLevelToShortString(Value As LogLevel) As String
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

    Public Function LogLevelToString(Value As LogLevel) As String
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

    Public Sub RemoveLogListener(listener As ILogListener)
        SyncLock mLogListeners
            mLogListeners.Remove(listener)
        End SyncLock
        listener.Dispose()
    End Sub

    ''' <summary>
    '''  Sets up logging to write all logged events for all infotypes to a 
    '''  default log file.
    ''' </summary>
    ''' <param name="synchronized">Set to true if multiple threads may use 
    ''' logging simultaneously.</param>
    ''' <param name="overwriteExisting">
    ''' If <c>True</c>, an existing log file of the same name is overwritten. 
    ''' Otherwise new log entries are appended to an existing log file of the 
    ''' same name: if none exists, a new log file is created.
    ''' </param>
    ''' <param name="createBackup">
    ''' If <c>True</c>, a backup copy of an existing log file of the same name 
    ''' is created.
    ''' </param>
    ''' <returns>The path and filename of the log file. Note that this may 
    ''' differ from the value returned by <c>DefaultLogFileName</c> before 
    ''' this method was called, as the filename may have been modified (for 
    ''' example if the specified logfile is already in use).</returns>
    ''' <remarks>
    ''' The logging level is governed by the global <see cref="DefaultLogLevel"/> 
    ''' property, but can be overridden if the application's command line 
    ''' specifies the <c>/loglevel</c> switch, which can take any of the 
    ''' following values:
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
    ''' If the application's command line includes the /log switch, then
    ''' the switch's value specifies the log file path and filename.
    '''
    ''' Otherwise, the log file is called <c>logfile.log</c>, and is created 
    ''' in the application's settings folder as determined by the 
    ''' <see cref="DataStorage.ApplicationDataPathUser"/> property).
    ''' </remarks>
    ''' 
    Public Function SetupDefaultLogging(
                    Optional synchronized As Boolean = False,
                    Optional overwriteExisting As Boolean = True,
                    Optional createBackup As Boolean = False) As String
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

        Dim logFile = CreateWriteableTextFile(DefaultLogFileName,
                                              Not overwriteExisting,
                                              createBackup,
                                              True)

        AddLogListener("", New FileLogListener(logFile), synchronized)

        Dim ass = Assembly.GetEntryAssembly
        If ass IsNot Nothing Then DefaultLogger.Log(ass.FullName)
        DefaultLogger.Log(Assembly.GetExecutingAssembly.FullName)
        DefaultLogger.Log("Log level: " & LogLevelToString(DefaultLogLevel))
        DefaultLogger.Log("Log file:  " & DefaultLogFileName)
        Return logFile.Name
    End Function

End Module

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

Public NotInheritable Class LogRecord


    ''
    ' Objects of this class are the unit of logged information.
    '
    '@/

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


    Private Const ModuleName As String = NameOf(LogRecord)

    '@================================================================================
    ' Member variables
    '@================================================================================

    Private mLogLevel As LogLevel
    Private mInfoType As String
    Private mTimestamp As Date
    Private mSequenceNumber As Integer
    Private mData As Object
    Private mSource As Object

    '@================================================================================
    ' Class Event Handlers
    '@================================================================================

    Private Sub New()
    End Sub

    Friend Sub New(pLevel As LogLevel, infoType As String, sequenceNumber As Integer, pData As Object, Optional pSource As Object = Nothing)
        mTimestamp = Date.Now
        mLogLevel = pLevel
        mInfoType = infoType
        mSequenceNumber = sequenceNumber
        mData = pData
        mSource = pSource
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
    ' Returns the Data for this Log Record.
    '
    ' @return
    '   The Data for this Log Record.
    '@/
    Public ReadOnly Property Data() As Object
        Get
            Return mData
        End Get
    End Property


    ''
    ' Returns the information type for the <c>data</c> property.
    '
    ' @return
    '   The information type for the <c>data</c> property.
    '@/
    Public ReadOnly Property InfoType() As String
        Get
            InfoType = mInfoType
        End Get
    End Property


    ''
    ' Returns the log level for this Log Record.
    '
    ' @return
    '   The log level for this object.
    '@/
    Public ReadOnly Property LogLevel() As LogLevel
        Get
            LogLevel = mLogLevel
        End Get
    End Property


    ''
    ' Returns the sequence number for this Log Record.
    '
    ' @remarks
    '   Sequence numbers are allocated consecutively to log records as they are created.
    '
    ' @return
    '   The sequence number for this object.
    '@/
    Public ReadOnly Property SequenceNumber() As Integer
        Get
            SequenceNumber = mSequenceNumber
        End Get
    End Property


    ''
    ' Returns information that identifies the Source of this Log Record (this information
    ' could for example be a reference to an object, or something that uniquely identifies
    ' an object).
    '
    ' @remarks
    '   This Value may be <c>Empty</c> where there is no need to distinguish between
    '   log records from different sources.
    '
    ' @return
    '   Information identifying the Source of this Log Record.
    '@/
    Public ReadOnly Property Source() As Object
        Get
            Return mSource
        End Get
    End Property


    ''
    ' Returns the Timestamp for this Log Record.
    '
    ' @return
    '   The Timestamp for this object.
    '@/
    Public ReadOnly Property Timestamp() As Date
        Get
            Timestamp = mTimestamp
        End Get
    End Property

    '@================================================================================
    ' Methods
    '@================================================================================

    '@================================================================================
    ' Helper Functions
    '@================================================================================

End Class


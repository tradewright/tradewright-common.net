Public Class LogRecordEventArgs
    Inherits EventArgs

    Public Sub New(logRecord As LogRecord)
        Me.LogRecord = logRecord
    End Sub

    Public ReadOnly Property LogRecord() As LogRecord
End Class


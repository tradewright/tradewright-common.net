Imports TradeWright.Utilities.Logging

Public Interface ILogger
    Property LogLevel As LogLevel
    Property LogToParent As Boolean
    Event Finished(sender As Object, e As EventArgs)
    Event LogRecord(sender As Object, e As LogRecordEventArgs)
    Sub Dispose()
    Sub Log(pLevel As LogLevel, pData As Object, Optional pSource As Object = Nothing)
    Function IsLoggable(level As LogLevel) As Boolean
End Interface

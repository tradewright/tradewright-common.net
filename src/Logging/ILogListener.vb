Public Interface ILogListener
    Inherits IDisposable
    Sub Initialise(logger As Logger, synchronized As Boolean)
End Interface

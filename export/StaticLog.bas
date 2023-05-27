Attribute VB_Name = "StaticLog"
'@Folder("Logging")
Option Explicit

Public Function Log() As IDebugEx
    Static mLog As IDebugEx
    If mLog Is Nothing Then
        Set mLog = New DebugEx
        With mLog
            .AddProvider FileLoggingProvider.Create
            .AddProvider ImmediateLoggingProvider.Create
            .StartLogging
            .LogClear
        End With
    End If
    
    Set Log = mLog
End Function


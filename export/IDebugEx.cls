VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDebugEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder "Logging.Abstract"
Option Explicit

Public Enum LogLevel
    NoOutput_Level = 1
    UI_Level
    Error_level
    Warning_Level
    Info_level
    Debug_Level
    Verbose_Level
End Enum

Public Sub Many(ByVal ArrayToLog As Variant, Optional ByVal Topic As Variant, Optional ByVal LogLevel As LogLevel = -1)
End Sub

Public Sub Variable(ByVal VariableToLog As Variant, Optional ByVal Topic As Variant, Optional ByVal LogLevel As LogLevel = -1)
End Sub

Public Sub Message(ByVal Message As String, Optional ByVal Topic As Variant, Optional ByVal LogLevel As LogLevel = -1)
End Sub

Public Sub LogStop(ByVal Message As String, Optional ByVal Topic As Variant, Optional ByVal LogLevel As LogLevel = -1)
End Sub

Public Sub LogHR()
End Sub

Public Sub LogClear()
End Sub

Public Sub SetDefaultLevel(ByVal LogLevel As LogLevel)
End Sub

Public Sub SetFilterLevel(ByVal LogLevel As LogLevel)
End Sub

Public Sub StartLogging()
End Sub

Public Sub StopLogging()
End Sub

Public Sub AddProvider(ByVal Provider As ILoggingProvider)
End Sub

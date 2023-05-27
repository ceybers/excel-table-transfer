VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImmediateLoggingProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "Logging.Providers"
Option Explicit
Implements ILoggingProvider

Private Sub ILoggingProvider_WriteDebugMessage(ByVal DMessage As DebugMessage)
    Debug.Assert Not Me Is ImmediateLoggingProvider
    
    Dim OutputText As String
    OutputText = OutputText & "[" & DMessage.Timestamp & "] "
    OutputText = OutputText & "[" & DMessage.Topic & "]" & vbTab
    OutputText = OutputText & DMessage.Message
    
    Debug.Print OutputText
End Sub

Public Function Create() As ILoggingProvider
    Dim Result As ImmediateLoggingProvider
    Set Result = New ImmediateLoggingProvider
    Set Create = Result
End Function

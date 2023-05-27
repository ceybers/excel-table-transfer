VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileLoggingProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "Logging.Providers"
Option Explicit
Implements ILoggingProvider

Private Const DEFAULT_OUTPUT_FILE As String = "C:\Users\User\desktop\debugex.log"

Private Type TState
    OutputFile As String
End Type
Private This As TState

Public Property Get OutputFile() As String
    Debug.Assert Not Me Is FileLoggingProvider
    OutputFile = This.OutputFile
End Property

Public Property Let OutputFile(ByVal vNewValue As String)
    Debug.Assert Not Me Is FileLoggingProvider
    This.OutputFile = vNewValue
End Property

Private Sub ILoggingProvider_WriteDebugMessage(ByVal DMessage As DebugMessage)
    Debug.Assert Not Me Is FileLoggingProvider
    
    Dim OutputText As String
    
    OutputText = OutputText & "[" & DMessage.Timestamp & "] "
    OutputText = OutputText & "[" & DMessage.Topic & "]" & vbTab
    OutputText = OutputText & DMessage.Message
    OutputText = DMessage.Color & "," & OutputText
    
    Dim FileNumber As Long
    FileNumber = FreeFile
    Open This.OutputFile For Append Access Write Shared As #FileNumber
    Write #FileNumber, OutputText
    Close #FileNumber
End Sub

Public Function Create(Optional ByVal OutputFile As String) As ILoggingProvider
    Dim Result As FileLoggingProvider
    Set Result = New FileLoggingProvider
    
    If OutputFile = vbNullString Then
        Result.OutputFile = DEFAULT_OUTPUT_FILE
    Else
        Result.OutputFile = OutputFile
    End If
    
    Set Create = Result
End Function

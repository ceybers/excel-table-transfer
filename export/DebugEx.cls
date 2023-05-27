VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "Logging"
Option Explicit
Implements IDebugEx

Private Const DEFAULT_TOPIC As String = "default"
Private Const SYSTEM_TOPIC As String = "System"
Private Const MAX_MANY_ITEMS_PRINT As Long = 10

Private Type TState
    StartTime As Double
    ItemCount As Long
    CurrentLevel As LogLevel
    FilterLevel As LogLevel
    ManyItemsPrint As Long
    Providers As Collection
End Type
Private This As TState

Private Sub Class_Initialize()
    Set This.Providers = New Collection
End Sub

Private Sub IDebugEx_SetFilterLevel(ByVal LogLevel As LogLevel)
    This.FilterLevel = LogLevel
End Sub

Private Sub IDebugEx_SetDefaultLevel(ByVal LogLevel As LogLevel)
    This.CurrentLevel = LogLevel
End Sub

Private Sub IDebugEx_StartLogging()
    This.StartTime = Timer()
    This.ItemCount = 0
    This.CurrentLevel = NoOutput_Level
    This.FilterLevel = Verbose_Level
    This.ManyItemsPrint = MAX_MANY_ITEMS_PRINT
    IDebugEx_Message "Logging started...", SYSTEM_TOPIC, UI_Level
End Sub

Private Sub IDebugEx_StopLogging()
    IDebugEx_Message "Logging stopped.", SYSTEM_TOPIC, UI_Level
    IDebugEx_Message This.ItemCount & " event(s) logged.", SYSTEM_TOPIC, UI_Level
    IDebugEx_Message "Duration: " & GetDuration, SYSTEM_TOPIC, UI_Level
    IDebugEx_Message String$(80, "-"), SYSTEM_TOPIC, UI_Level
End Sub

Private Sub IDebugEx_Message(ByVal Message As String, Optional ByVal Topic As Variant, _
    Optional ByVal LogLevel As LogLevel = -1)
    Dim DMessage As DebugMessage
    Set DMessage = New DebugMessage
    With DMessage
        .Message = Message
        .Topic = ConvertOptionalTopic(Topic, IsMissing(Topic))
        .Timestamp = GetDuration
        .LogLevel = IIf(LogLevel > 0, LogLevel, This.CurrentLevel)
        .Color = LogLevelToColor(.LogLevel)
    End With
    
    WriteDebugMessage DMessage
End Sub

Private Sub IDebugEx_LogStop(ByVal Message As String, Optional ByVal Topic As Variant, _
    Optional ByVal LogLevel As LogLevel = -1)
    IDebugEx_Message Message, Topic, LogLevel
    '@Ignore StopKeyword
    Stop
End Sub

Private Sub IDebugEx_LogHR()
    IDebugEx_Message String$(80, "-"), "Rule", UI_Level
End Sub

Private Sub IDebugEx_LogClear()
    IDebugEx_Message Space$(80), "Clear", UI_Level
    IDebugEx_Message Space$(80), "Clear", UI_Level
    IDebugEx_Message Space$(80), "Clear", UI_Level
End Sub

Private Sub IDebugEx_Many(ByVal ArrayToLog As Variant, Optional ByVal Topic As Variant, _
    Optional ByVal LogLevel As LogLevel = -1)
    If Not IsArray(ArrayToLog) Then Exit Sub
    
    Dim Rank As Long
    Rank = 1
    On Error Resume Next
    Dim vTest As Variant
    vTest = ArrayToLog(LBound(ArrayToLog), 1)
    If Not IsEmpty(vTest) Then Rank = 2
    On Error GoTo 0
    
    Dim FirstIndex As Long
    Dim LastIndex As Long
    FirstIndex = LBound(ArrayToLog, 1)
    LastIndex = UBound(ArrayToLog, 1)
    
    Dim Rank2Size As String
    If Rank = 2 Then
        Rank2Size = ", " & LBound(ArrayToLog, 2) & " to " & UBound(ArrayToLog, 2)
    End If
    IDebugEx_Message "Array size (" & FirstIndex & " to " & LastIndex & Rank2Size & ")", Topic, LogLevel
    
    If (LastIndex - FirstIndex) > This.ManyItemsPrint Then
        LastIndex = FirstIndex + This.ManyItemsPrint
    End If
    
    Dim i As Long
    For i = FirstIndex To LastIndex
        If Rank = 1 Then
            IDebugEx_Variable "• " & Format(i, "00") & ": " & (ArrayToLog(i)), Topic, LogLevel
        Else
            If IsError(ArrayToLog(i, 1)) Then
                IDebugEx_Variable "• " & Format(i, "00") & ": " & CStr(ArrayToLog(i, 1)), Topic, LogLevel
            Else
                ' Need to change this to separate prefix (bullet and index) from variant
                IDebugEx_Variable (ArrayToLog(i, 1)), Topic, LogLevel
            End If
        End If
    Next i
    
    Dim Count As Long
    Count = LastIndex - FirstIndex
    
    IDebugEx_Message "----- ", Topic, LogLevel 'Done printing (first) " & Count & " item(s)", Topic, LogLevel
End Sub

Private Sub IDebugEx_Variable(ByVal VariableToLog As Variant, Optional ByVal Topic As Variant, _
    Optional ByVal LogLevel As LogLevel = -1)
    Dim Message As String
    If IsObject(VariableToLog) Then
        Message = "Object of Type " & TypeName(VariableToLog) & " Is Nothing = " & CStr(VariableToLog Is Nothing)
    Else
        If IsArray(VariableToLog) Then ' Only works if it is a 1-dimensional array
            Message = "Array of " & UBound(VariableToLog) & " item(s) (VarType = " & VarType(VariableToLog) & ", TypeName = " & TypeName(VariableToLog) & ")"
        Else
            Message = CStr(VariableToLog) & " (VarType = " & VarType(VariableToLog) & ", TypeName = " & TypeName(VariableToLog) & ")"
        End If
    End If
   
    If IsMissing(Topic) Then Topic = TypeName(VariableToLog)
    IDebugEx_Message Message, ConvertOptionalTopic(Topic, IsMissing(Topic)), LogLevel
End Sub

Public Function Create() As IDebugEx
    Set Create = New DebugEx
End Function

Private Function LogLevelToColor(ByVal LogLevel As Variant) As String
    ' PowerShell Write-Host -Foreground Colors:
    '    Black, DarkBlue, DarkGreen, DarkCyan, DarkRed, DarkMagenta, DarkYellow, Gray,
    '    DarkGray, Blue, Green, Cyan, Red, Magenta , Yellow, White
    Select Case LogLevel
        Case UI_Level: LogLevelToColor = "Blue"
        Case NoOutput_Level: LogLevelToColor = "White" 'default
        Case Error_level: LogLevelToColor = "DarkRed"
        Case Warning_Level: LogLevelToColor = "Magenta"
        Case Info_level: LogLevelToColor = "White"
        Case Debug_Level: LogLevelToColor = "DarkYellow"
        Case Verbose_Level: LogLevelToColor = "Green"
    End Select
End Function

Private Sub WriteDebugMessage(ByVal DMessage As DebugMessage)
    If DMessage.LogLevel > This.FilterLevel Then
        Exit Sub
    End If
    
    Dim Provider As ILoggingProvider
    For Each Provider In This.Providers
        Provider.WriteDebugMessage DMessage
    Next Provider
    
    This.ItemCount = This.ItemCount + 1
End Sub

Private Function ConvertOptionalTopic(ByVal Topic As Variant, ByVal WasMissing As Boolean) As String
    If Not WasMissing Then
        If IsObject(Topic) Then
            ConvertOptionalTopic = TypeName(Topic)
        Else
            ConvertOptionalTopic = CStr(Topic)
        End If
    Else
        ConvertOptionalTopic = DEFAULT_TOPIC
    End If
End Function

Private Function GetDuration() As String
    GetDuration = Format$(Timer() - This.StartTime, "00:00.00")
End Function

Private Sub IDebugEx_AddProvider(ByVal Provider As ILoggingProvider)
    This.Providers.Add Item:=Provider
End Sub


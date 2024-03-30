Attribute VB_Name = "DebugTime"
'@Folder "Helpers.Common"
Option Explicit

Private Const DEBUG_TIME_ENABLED As Boolean = False
Private Const TIME_FORMAT_STRING As String = "000.000"

Public Sub PrintTime(ByVal Message As String, Optional ByVal Reset As Boolean)
    If DEBUG_TIME_ENABLED = False Then Exit Sub
    
    Static StartTime As Double
    If StartTime = 0 Then
        StartTime = Timer()
        Debug.Print String(40, "-")
    End If
    If Reset Then StartTime = Timer()
    
    Debug.Print "["; Format$((Timer() - StartTime), TIME_FORMAT_STRING); "] "; Message
End Sub

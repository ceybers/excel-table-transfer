Attribute VB_Name = "TimerHelper"
'@Folder("Common.Helpers")
Option Explicit

Public Function PrintTime(ByVal Message As String, Optional ByVal Reset As Boolean)
    Static StartTime As Double
    If Reset Or (StartTime = 0) Then
        StartTime = Timer()
    End If
    
    Debug.Print Message & " " & (Timer() - StartTime)
End Function

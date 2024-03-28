Attribute VB_Name = "WorkbookHelpers"
'@Folder "Helpers.Objects"
Option Explicit

Public Function GetPathFromRangeText(ByVal payload As String) As String
    Dim a As Integer
    Dim b As Integer
    a = InStr(payload, "'")
    b = InStr(payload, "[")
    If a = 0 Or b = 0 Then Exit Function
    GetPathFromRangeText = Mid$(payload, a + 1, b - a - 1)
End Function

Public Function GetFilenameFromRangeText(ByVal payload As String) As String
    Dim b As Integer
    Dim c As Integer
    b = InStr(payload, "[")
    c = InStr(payload, "]")
    If b = 0 Or c = 0 Then Exit Function
    GetFilenameFromRangeText = Mid$(payload, b + 1, c - b - 1)
End Function

Public Function IsWorkbookOpen(ByVal filename As String) As Boolean
    Dim wb As Workbook
    If filename = vbNullString Then Exit Function
    
    For Each wb In Application.Workbooks
        If wb.Name = filename Then
            IsWorkbookOpen = True
            Exit Function
        End If
    Next wb
End Function


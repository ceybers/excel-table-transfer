Attribute VB_Name = "WorkbookHelpers"
'@Folder "Helpers.Objects"
Option Explicit

Public Function GetPathFromRangeText(ByVal payload As String) As String
    Dim SingleQuote As Long
    SingleQuote = InStr(payload, "'")
    
    Dim SquareBracketLeft As Long
    SquareBracketLeft = InStr(payload, "[")
    
    If SingleQuote = 0 Or SquareBracketLeft = 0 Then Exit Function
    
    GetPathFromRangeText = Mid$(payload, SingleQuote + 1, SquareBracketLeft - SingleQuote - 1)
End Function

Public Function GetFilenameFromRangeText(ByVal payload As String) As String
    Dim LeftBracket As Long
    LeftBracket = InStr(payload, "[")
    
    Dim RightBracket As Long
    RightBracket = InStr(payload, "]")
    
    If LeftBracket = 0 Or RightBracket = 0 Then Exit Function
    
    GetFilenameFromRangeText = Mid$(payload, LeftBracket + 1, RightBracket - LeftBracket - 1)
End Function

Public Function IsWorkbookOpen(ByVal Filename As String) As Boolean
    Dim Workbook As Workbook
    If Filename = vbNullString Then Exit Function
    
    For Each Workbook In Application.Workbooks
        If Workbook.Name = Filename Then
            IsWorkbookOpen = True
            Exit Function
        End If
    Next Workbook
End Function


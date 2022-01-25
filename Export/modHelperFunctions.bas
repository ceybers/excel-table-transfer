Attribute VB_Name = "modHelperFunctions"
'@Folder "HelperFunctions"
Option Explicit

Public Function TableFromString(ByVal s As String) As ListObject
    Dim wb As Workbook, ws As Worksheet, lo As ListObject
    Dim n() As String
    
    n = Split(s, "\")
    Debug.Assert LBound(n, 1) = 0
    Debug.Assert UBound(n, 1) = 2
    
    n(2) = Replace(n(2), " (active)", vbNullString)
    
    Set wb = Workbooks(n(0))
    Set ws = wb.Worksheets(n(1))
    Set lo = ws.ListObjects(n(2))
    
    Set TableFromString = lo
End Function

Public Function TableToString(ByVal lo As ListObject) As String
    Debug.Assert Not lo Is Nothing
    TableToString = lo.parent.parent.Name & "\" & lo.parent.Name & "\" & lo.Name
End Function

Public Function ToKey(ByVal i As Integer) As String
    Debug.Assert (i >= 0 And i <= 999)
    ToKey = "K" & Trim$(Format$(i, "000"))
End Function

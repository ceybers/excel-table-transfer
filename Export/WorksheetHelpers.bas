Attribute VB_Name = "WorksheetHelpers"
'@Folder "Helpers.Objects"
Option Explicit

Public Function TryGetWorksheetByName(ByVal worksheetName As String, ByRef Worksheet As Worksheet) As Boolean
    Set Worksheet = GetWorksheetByName(worksheetName)
    TryGetWorksheetByName = Not Worksheet Is Nothing
End Function

Public Function DoesWorksheetExist(ByVal worksheetName As String) As Boolean
    DoesWorksheetExist = Not GetWorksheetByName(worksheetName) Is Nothing
End Function

Public Function GetWorksheetByName(ByVal worksheetName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name = worksheetName Then
            Set GetWorksheetByName = ws
        End If
    Next ws
End Function


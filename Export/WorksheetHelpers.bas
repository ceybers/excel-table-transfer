Attribute VB_Name = "WorksheetHelpers"
'@IgnoreModule UseMeaningfulName, ProcedureNotUsed
'@Folder "Helpers.Objects"
Option Explicit

'@Description "Tests if a given string is a valid Worksheet name."
Public Function IsValidWorksheetName(ByVal SheetName As String) As Boolean
Attribute IsValidWorksheetName.VB_Description = "Tests if a given string is a valid Worksheet name."
    ' Reference: https://support.microsoft.com/en-us/office/rename-a-worksheet-3f1f7148-ee83-404d-8ef0-9ff99fbad1f9
    If SheetName = vbNullString Then Exit Function
    If Len(SheetName) > 31 Then Exit Function
    If Left$(SheetName, 1) = "'" Then Exit Function

    Dim InvalidChars As Variant
    InvalidChars = Array("\", "/", "?", "*", "[", "]", ":")

    Dim i As Long
    For i = 1 To UBound(InvalidChars)
        If InStr(SheetName, InvalidChars(i)) > 0 Then Exit Function
    Next i

    IsValidWorksheetName = True
End Function

'@Description "Tries to remove a Worksheet with a given name from a Workbook."
Public Function TryRemoveWorksheet(ByVal Workbook As Workbook, ByVal WorksheetName As String) As Boolean
Attribute TryRemoveWorksheet.VB_Description = "Tries to remove a Worksheet with a given name from a Workbook."
    If Workbook.Worksheets.Count = 1 Then Exit Function
    
    Dim Worksheet As Worksheet
    If Not TryGetWorksheet(Workbook, WorksheetName, Worksheet) Then Exit Function
    
    Application.DisplayAlerts = False
    Worksheet.Delete
    Application.DisplayAlerts = True
    TryRemoveWorksheet = True
End Function

'@Description "Returns True if a Worksheet with the given name exists in a Workbook."
Public Function WorksheetExists(ByVal Workbook As Workbook, ByVal WorksheetName As String) As Boolean
Attribute WorksheetExists.VB_Description = "Returns True if a Worksheet with the given name exists in a Workbook."
    Dim Worksheet As Worksheet
    WorksheetExists = TryGetWorksheet(Workbook, WorksheetName, Worksheet)
End Function

'@Description "Tries to return the Worksheet with the given name from a Workbook."
Public Function TryGetWorksheet(ByVal Workbook As Workbook, ByVal WorksheetName As String, ByRef OutWorksheet As Worksheet) As Boolean
Attribute TryGetWorksheet.VB_Description = "Tries to return the Worksheet with the given name from a Workbook."
    Dim Worksheet As Worksheet
    For Each Worksheet In Workbook.Worksheets
        If Worksheet.Name Like WorksheetName Then
            Set OutWorksheet = Worksheet
            TryGetWorksheet = True
        End If
    Next Worksheet
End Function

'@Description "Returns the Worksheet with a given name. If it does not exist, create and return it."
Public Function AddOrGetWorksheet(ByVal Workbook As Workbook, ByVal WorksheetName As String) As Worksheet
Attribute AddOrGetWorksheet.VB_Description = "Returns the Worksheet with a given name. If it does not exist, create and return it."
    Dim Worksheet As Worksheet
    
    If Not TryGetWorksheet(Workbook, WorksheetName, Worksheet) Then
        With Workbook.Worksheets
            Set Worksheet = .Add(After:=.Item(.Count))
        End With
        Worksheet.Name = WorksheetName
    End If
    
    Set AddOrGetWorksheet = Worksheet
End Function

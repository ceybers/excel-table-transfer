Attribute VB_Name = "ListObjectHelpers"
'@Folder "Common.Helpers.ListObject"
Option Explicit

Private Sub Test()
    Dim lo As ListObject
    Dim Result As String
    
    Set lo = ThisWorkbook.Worksheets(1).ListObjects(1)
    'Debug.Print ListObjectToStringName(lo)
    
    'result = "'C:\Users\User\Documents\excel-related-table-tool\[Production.xlsm]Sheet1'!$A$1:$C$5"
    'Set lo = StringNameToListObject(result)
    If lo Is Nothing Then
        Debug.Print "Path Filename Worksheet Range [FAIL]"
    Else
        Debug.Print "Path Filename Worksheet Range [OK]"
    End If
    
    Result = "[Production.xlsm]Sheet1!$A$1:$C$5" ' Closed file
    Result = "[Development.xlsm]Sheet1!$A$1:$C$5"
    Set lo = StringNameToListObject(Result)
    If lo Is Nothing Then
        Debug.Print "     Filename Worksheet Range [FAIL]"
    Else
        Debug.Print "     Filename Worksheet Range [OK]"
    End If
        
    Result = "Sheet1!A1:C10"
    Set lo = StringNameToListObject(Result)
    If lo Is Nothing Then
        Debug.Print "              Worksheet Range [FAIL]"
    Else
        Debug.Print "              Worksheet Range [OK]"
    End If
    
    'Debug.Print lo.Name
End Sub

Public Function GetListColumnFromRange(ByVal rng As Range) As ListColumn
    If rng Is Nothing Then Exit Function
    If rng.ListObject Is Nothing Then Exit Function
    If rng.Columns.Count <> 1 Then Exit Function
    
    Dim lo As ListObject
    Dim lc As ListColumn
    
    Set lo = rng.ListObject
    For Each lc In lo.ListColumns
        If lc.Range.column = rng.Cells(1, 1).column Then
            Set GetListColumnFromRange = lc
            Exit Function
        End If
    Next lc
End Function

Public Function StringNameToListObject(ByVal stringName As String) As ListObject
    Dim path As String
    Dim filename As String
    Dim worksheetName As String
    Dim rangetext As String
    
    Dim splitByRange As Variant
    Dim splitByQuotes As Variant
    splitByRange = Split(stringName, "!")
    splitByQuotes = Split(splitByRange(0), "'")
    
    rangetext = splitByRange(1)
    
    If UBound(splitByQuotes, 1) = 2 Then
        path = Split(Split(splitByQuotes(1), "]")(0), "[")(0)
        filename = Split(Split(splitByQuotes(1), "]")(0), "[")(1)
        worksheetName = Split(splitByQuotes(1), "]")(1)
    Else
        If UBound(Split(splitByRange(0), "]"), 1) = 0 Then
            worksheetName = splitByRange(0)
        Else
            filename = Split(splitByRange(0), "]")(0)
            filename = Right$(filename, Len(filename) - 1)
            worksheetName = Split(splitByRange(0), "]")(1)
        End If
    End If
    
    If False Then
        Debug.Print "Path="; path
        Debug.Print "Filename="; filename
        Debug.Print "Worksheetname="; worksheetName
        Debug.Print "Range="; rangetext
        Stop
    End If
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    
    If filename = vbNullString Then
        Set wb = ThisWorkbook
    Else
        If Not TryGetWorkbook(filename, wb, path) Then Exit Function
    End If
    
    If TryGetWorkSheet(wb, worksheetName, ws) Then
        Set rng = ws.Range(rangetext)
        'Debug.Print rng.Address
        If rng.ListObject Is Nothing Then
            Debug.Print "Table doesn't exist more"
        Else
            Set StringNameToListObject = rng.ListObject
        End If
    Else
        Debug.Print "Worksheet doesn't exist more"
    End If
    
    'Debug.Print vbNullString
End Function

Public Function TryGetWorkSheet(ByVal wb As Workbook, ByVal worksheetName As String, ByRef ws As Worksheet) As Boolean
    Dim curWS As Worksheet
    For Each curWS In wb.Worksheets
        If curWS.Name = worksheetName Then
            Set ws = curWS
            TryGetWorkSheet = True
        End If
    Next curWS
End Function

Public Function TryGetWorkbook(ByVal filename As String, ByRef wb As Workbook, Optional path As String = vbNullString) As Boolean
    Dim curWB As Workbook
    For Each curWB In Application.Workbooks
        If path = vbNullString Then
            If curWB.Name = filename Then
                Set wb = curWB
                TryGetWorkbook = True
                Exit Function
            End If
        Else
            If curWB.fullname = path & filename Then
                Set wb = curWB
                TryGetWorkbook = True
                Exit Function
            End If
        End If
    Next curWB
End Function

Public Function ListObjectToStringName(ByVal lo As ListObject, Optional ByVal ShowFilename As Boolean = False, Optional ByVal ShowPath As Boolean = False) As String
    Debug.Assert Not lo Is Nothing
    
    Dim ws As Worksheet
    Dim wb As Workbook
    
    Set ws = lo.parent
    Set wb = ws.parent
    
    If ShowFilename Or ShowPath Then
        ListObjectToStringName = lo.Range.Address(external:=True)
    Else
        ListObjectToStringName = lo.Range.Address(external:=False)
    End If
End Function

Public Function TryGetListColumnFromText(ByVal columnName As String, ByVal lo As ListObject, ByRef outLC As ListColumn) As Boolean
    If lo Is Nothing Then Exit Function
    
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If lc.Name = columnName Then
            Set outLC = lc
            TryGetListColumnFromText = True
            Exit Function
        End If
    Next lc
End Function

'@Description "test"
Public Function TryGetTableFromText(ByVal rangetext As String, ByRef outListObject As ListObject, Optional ByVal openIfClosed As Boolean = False) As Boolean
Attribute TryGetTableFromText.VB_Description = "test"
    ' Debug.Print "RR"; rangeText
    
    ' Example:
    ' rangeText = [Development.xlsm]Sheet1!$A$1:$D$11
    
    ' Is the sheet open? Does it have the same path, or is just the filename the same? (but different folder)
    ' If it's not open, ask the user. Try to open: did it succeed or fail?
    ' Does the worksheet still exist?
    ' Does the listobject still exist in that range?
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim sheetname As String
    Dim rangeaddress As String
    Dim filename As String
    
    filename = WorkbookHelpers.GetFilenameFromRangeText(rangetext)
    If IsWorkbookOpen(filename) Then
        Set wb = Application.Workbooks(filename)
        sheetname = Mid$(rangetext, InStr(rangetext, "]") + 1, InStr(rangetext, "!") - InStr(rangetext, "]") - 1)
        Set ws = wb.Worksheets(sheetname) ' TODO Error trap this
        rangeaddress = Mid$(rangetext, InStr(rangetext, "!") + 1, Len(rangetext) - InStr(rangetext, "!") - 1)
        'Debug.Print rangeaddress
        Set rng = ws.Range(rangeaddress)
        'Debug.Print rng.Address
        'If rng.ListObject Is Nothing Then Err.Raise 5, , "TryGetTableFromText failed"
        
        If Not rng.ListObject Is Nothing Then
            Set outListObject = rng.ListObject
            TryGetTableFromText = True
        End If
    Else
        MsgBox "TryGetTableFromText DoOpen NYI"
        Dim path As String
        path = WorkbookHelpers.GetPathFromRangeText(rangetext)
        'Debug.Print path
        ' Need a flag for Transfer class if we cannot open the workbook from an old serialized instructionf
    End If
End Function

Public Function GetColumnHeaderFromListColumn(ByVal lc As ListColumn) As String
    Debug.Assert Not lc Is Nothing
    Dim s As String
    s = lc.Range.EntireColumn.Address
    s = Mid$(s, 2, ((Len(s) - 1) / 2) - 1)
    GetColumnHeaderFromListColumn = s
End Function

Public Function ListColumnHasArray(ByVal lc As ListColumn) As Boolean
    Debug.Assert Not lc Is Nothing
    
    If lc.DataBodyRange Is Nothing Then Exit Function
    
    If IsNull(lc.DataBodyRange.FormulaArray) Then
        'Debug.Print "cells are different"
        ListColumnHasArray = False
    ElseIf Left$(lc.DataBodyRange.FormulaArray, 1) = "=" Then
        'Debug.Print "same formula"
        ListColumnHasArray = True
    Else
        'Debug.Print "same non-formula"
        ListColumnHasArray = False
    End If
End Function

Public Function TryGetListObjectFromWorkbook(ByVal wb As Workbook, ByVal loName As String, ByRef outLO As ListObject) As Boolean
    If wb Is Nothing Then Exit Function
    
    Dim ws As Worksheet
    Dim lo As ListObject
    
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = loName Then
                Set outLO = lo
                TryGetListObjectFromWorkbook = True
                Exit Function
            End If
        Next lo
    Next ws
End Function

Attribute VB_Name = "ListObjectHelpers"
'@Folder "Helpers.Objects"
Option Explicit

' Ref by KeyMapper?
Public Function GetListColumnFromRange(ByVal rng As Range) As ListColumn
    If rng Is Nothing Then Exit Function
    If rng.ListObject Is Nothing Then Exit Function
    If rng.Columns.Count <> 1 Then Exit Function
    
    Dim lo As ListObject
    Dim lc As ListColumn
    
    Set lo = rng.ListObject
    For Each lc In lo.ListColumns
        If lc.Range.Column = rng.Cells.Item(1, 1).Column Then
            Set GetListColumnFromRange = lc
            Exit Function
        End If
    Next lc
End Function

' Ref by TransferInstruction
Public Function TryGetWorkbook(ByVal Filename As String, ByRef wb As Workbook, Optional ByVal path As String = vbNullString) As Boolean
    Dim curWB As Workbook
    For Each curWB In Application.Workbooks
        If path = vbNullString Then
            If curWB.Name = Filename Then
                Set wb = curWB
                TryGetWorkbook = True
                Exit Function
            End If
        Else
            If curWB.fullname = path & Filename Then
                Set wb = curWB
                TryGetWorkbook = True
                Exit Function
            End If
        End If
    Next curWB
End Function

' Ref by ValueMapperViewModel
Public Function GetColumnHeaderFromListColumn(ByVal lc As ListColumn) As String
    Debug.Assert Not lc Is Nothing
    Dim s As String
    s = lc.Range.EntireColumn.Address
    s = Mid$(s, 2, ((Len(s) - 1) / 2) - 1)
    GetColumnHeaderFromListColumn = s
End Function

' Ref by ValueMapperViewModel
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

' Ref by TransferInstructionUnref
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

' Ref by KeyMapperViewModel
Public Function ListColumnExists(ByVal ListObject As ListObject, ByVal ListColumnName As String) As Boolean
    Dim ListColumn As ListColumn
    For Each ListColumn In ListObject.ListColumns
        If ListColumn.Name = ListColumnName Then
            ListColumnExists = True
            Exit Function
        End If
    Next ListColumn
End Function


Public Function GetAllTablesInApplication() As Collection
    Dim Result As Collection
    Set Result = New Collection
    
    Dim Workbook As Workbook
    Dim Worksheet As Worksheet
    Dim ListObject As ListObject
    
    For Each Workbook In Application.Workbooks
        For Each Worksheet In Workbook.Worksheets
            For Each ListObject In Worksheet.ListObjects
                Result.Add ListObject, ListObject.Range.Address(External:=True)
            Next ListObject
        Next Worksheet
    Next Workbook
    
    Set GetAllTablesInApplication = Result
End Function

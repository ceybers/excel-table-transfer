Attribute VB_Name = "ListColumnHelpers"
'@Folder "Helpers"
Option Explicit

Private Const NONE_CAPTION As String = "None"
Private Const SOME_CAPTION As String = "Some"
Private Const ALL_CAPTION As String = "All"

Public Function GetR1C1(ByVal ListColumn As ListColumn) As String
    GetR1C1 = ListColumn.Range.EntireColumn.Address(ColumnAbsolute:=False)
    GetR1C1 = Left$(GetR1C1, InStr(GetR1C1, ":") - 1)
End Function

Public Function TotalRows(ByVal ListColumn As ListColumn) As Long
    TotalRows = ListColumn.DataBodyRange.Rows.Count
End Function

Public Function ColumnHasBlanks(ByVal ListColumn As ListColumn) As String
    ColumnHasBlanks = ColumnHasSpecialCells(ListColumn, xlCellTypeBlanks)
End Function

Public Function ColumnHasErrors(ByVal ListColumn As ListColumn) As String
    Dim ConstantError As String
    Dim FormulaeError As String
    ConstantError = ColumnHasSpecialCells(ListColumn, xlCellTypeConstants, xlErrors)
    FormulaeError = ColumnHasSpecialCells(ListColumn, xlCellTypeFormulas, xlErrors)
    
    ' Doesn't work. A column that is all errors but mixed constants and formulaes will return SOME + SOME -> SOME
    ' Need to count the numbers for both and add them up
    If ConstantError = ALL_CAPTION Or FormulaeError = ALL_CAPTION Then
        ColumnHasErrors = ALL_CAPTION
    ElseIf ConstantError = NONE_CAPTION And FormulaeError = NONE_CAPTION Then
        ColumnHasErrors = NONE_CAPTION
    Else
        ColumnHasErrors = SOME_CAPTION
    End If
End Function

Public Function ColumnHasFormulae(ByVal ListColumn As ListColumn) As String
    ColumnHasFormulae = ColumnHasSpecialCells(ListColumn, xlCellTypeFormulas)
End Function

Public Function ColumnHasValidation(ByVal ListColumn As ListColumn) As String
    ColumnHasValidation = ColumnHasSpecialCells(ListColumn, xlCellTypeAllValidation)
End Function

Public Function ColumnIsLocked(ByVal ListColumn As ListColumn) As String
    If ListColumn.DataBodyRange.Locked = True Then
        ColumnIsLocked = ALL_CAPTION
    ElseIf ListColumn.DataBodyRange.Locked = False Then
            ColumnIsLocked = NONE_CAPTION
    ElseIf IsNull(ListColumn.DataBodyRange.Locked) Then
        ColumnIsLocked = SOME_CAPTION
    Else
        Debug.Assert False
    End If
End Function

Private Function ColumnHasSpecialCells(ByVal ListColumn As ListColumn, ByVal XlCellType As Long, Optional XlSpecialCellsValue As Long = -1) As String
    Dim BlankRowRange As Range
    On Error Resume Next
    If XlSpecialCellsValue = -1 Then
        Set BlankRowRange = ListColumn.DataBodyRange.SpecialCells(Type:=XlCellType)
    Else
        Set BlankRowRange = ListColumn.DataBodyRange.SpecialCells(Type:=XlCellType, Value:=XlSpecialCellsValue)
    End If
    On Error GoTo 0
    
    If BlankRowRange Is Nothing Then
        ColumnHasSpecialCells = NONE_CAPTION
        Exit Function
    End If
    
    If BlankRowRange.Cells.Count = TotalRows(ListColumn) Then
        ColumnHasSpecialCells = ALL_CAPTION
        Exit Function
    End If
    
    ColumnHasSpecialCells = SOME_CAPTION
End Function

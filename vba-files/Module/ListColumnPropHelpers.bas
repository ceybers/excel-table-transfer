Attribute VB_Name = "ListColumnPropHelpers"
'@Folder "Helpers.ListColumn"
Option Explicit

Private Const NONE_CAPTION As String = "None"
Private Const SOME_CAPTION As String = "Some"
Private Const ALL_CAPTION As String = "All"

Public Enum Result
    None
    Some
    All
End Enum

Public Function GetR1C1(ByVal ListColumn As ListColumn) As String
    GetR1C1 = ListColumn.Range.EntireColumn.Address(ColumnAbsolute:=False)
    GetR1C1 = Left$(GetR1C1, InStr(GetR1C1, ":") - 1)
End Function

Private Function TotalRows(ByVal ListColumn As ListColumn) As Long
    TotalRows = ListColumn.DataBodyRange.Rows.Count
End Function

Public Function ColumnHasBlanks(ByVal ListColumn As ListColumn) As Result
    ColumnHasBlanks = ColumnHasSpecialCells(ListColumn, xlCellTypeBlanks)
End Function

Public Function ColumnHasErrors(ByVal ListColumn As ListColumn) As Result
    ColumnHasErrors = Result.None
    Exit Function
    
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

Public Function ColumnHasFormulae(ByVal ListColumn As ListColumn) As Result
    ColumnHasFormulae = ColumnHasSpecialCells(ListColumn, xlCellTypeFormulas)
End Function

Public Function ColumnHasValidation(ByVal ListColumn As ListColumn) As Result
    ColumnHasValidation = ColumnHasSpecialCells(ListColumn, xlCellTypeAllValidation)
End Function

Public Function ColumnIsLocked(ByVal ListColumn As ListColumn) As Result
    If ListColumn.DataBodyRange.Locked = True Then
        ColumnIsLocked = Result.All
    ElseIf ListColumn.DataBodyRange.Locked = False Then
            ColumnIsLocked = Result.None
    ElseIf IsNull(ListColumn.DataBodyRange.Locked) Then
        ColumnIsLocked = Result.Some
    Else
        Debug.Assert False
    End If
End Function

Private Function ColumnHasSpecialCells(ByVal ListColumn As ListColumn, ByVal XlCellType As Long, Optional ByVal XlSpecialCellsValue As Long = -1) As Result
    Dim BlankRowRange As Range
    On Error Resume Next
    If XlSpecialCellsValue = -1 Then
        Set BlankRowRange = ListColumn.DataBodyRange.SpecialCells(Type:=XlCellType)
    Else
        Set BlankRowRange = ListColumn.DataBodyRange.SpecialCells(Type:=XlCellType, Value:=XlSpecialCellsValue)
    End If
    On Error GoTo 0
    
    If BlankRowRange Is Nothing Then
        ColumnHasSpecialCells = Result.None
        Exit Function
    End If
    
    If BlankRowRange.Cells.Count = TotalRows(ListColumn) Then
        ColumnHasSpecialCells = Result.All
        Exit Function
    End If
    
    ColumnHasSpecialCells = Result.Some
End Function

Public Function EnumToString(ByVal EnumValue As Result) As String
    Select Case EnumValue
        Case Result.None:
            EnumToString = NONE_CAPTION
        Case Result.Some:
            EnumToString = SOME_CAPTION
        Case Result.All:
            EnumToString = ALL_CAPTION
    End Select
End Function

Public Function ColumnIsUnique(ByVal ListColumn As ListColumn) As Result
    Dim Dict As Scripting.Dictionary
    Set Dict = New Scripting.Dictionary
    
    Dim Value2 As Variant
    Value2 = ListColumn.DataBodyRange.Value2
    
    Dim CellValue2 As Variant
    On Error Resume Next
    For Each CellValue2 In Value2
        Dict.Item(CellValue2) = CellValue2
    Next CellValue2
    On Error GoTo 0
    
    If Dict.Count = UBound(Value2) Then
        ColumnIsUnique = Result.All
    Else
        ColumnIsUnique = Result.None
    End If
End Function

Public Function GetVarTypeOfColumnRange(ByVal Range As Range) As Long
    Debug.Assert Not Range Is Nothing
    Debug.Assert Range.Columns.Count = 1
    
    Dim Result As Long
    
    Dim ValueVariant As Variant
    ValueVariant = Range.Value
    
    If Range.Rows.Count = 1 Then
        GetVarTypeOfColumnRange = VarType(ValueVariant)
        Exit Function
    End If
    
    Dim i As Long
    For i = 1 To UBound(ValueVariant, 1)
        Select Case VarType(ValueVariant(i, 1))
            Case vbString:
                Result = IIf(Result = vbEmpty Or Result = vbString, vbString, vbVariant)
            Case vbDouble:
                Result = IIf(Result = vbEmpty Or Result = vbDouble, vbDouble, vbVariant)
            Case vbCurrency:
                Result = IIf(Result = vbEmpty Or Result = vbCurrency, vbCurrency, vbVariant)
            Case vbDate:
                Result = IIf(Result = vbEmpty Or Result = vbDate, vbDate, vbVariant)
        End Select
    Next i
    
    GetVarTypeOfColumnRange = Result
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule UnassignedVariableUsage
'@PredeclaredId
'@Folder "MVVM.Models.KeyColumn"
Option Explicit

Private Const EMPTY_TABLE_MESSAGE As String = "Cannot create clsKeyColumn with a column in an empty table"
Private Const DISJOINT_RANGE_MESSAGE As String = "Cannot create clsKeyColumn with a range that spans multiple columns"
Private Const NO_CELLS_AFTER_FILTER_MESSAGE As String = "No cells were visible after filtering"
Private Const LOAD_RANGE_FAILED As String = "KeyColumn.LoadRange failed"
Private Const MAX_LIMIT As Long = 2147483647

Private Type TClsKeyColumn
    Dict As Object
    Range As Range
    Array As Variant
    
    Distinct As Collection
    Unique As Collection
    
    IncludeNumbersAsText As Boolean
    RespectFilter As Boolean
    
    ErrorCount As Long
    BlankCount As Long
    NonTextCount As Long
    
    Limit As Long
End Type
Private This As TClsKeyColumn

Public Property Get Limit() As Double
    Limit = This.Limit
End Property

Public Property Let Limit(ByVal vNewValue As Double)
    This.Limit = vNewValue
End Property

Public Property Get DistinctPercentage() As Double
    If This.Range.Cells.Count = 0 Then Exit Property
    DistinctPercentage = This.Distinct.Count / This.Range.Cells.Count
End Property

Public Property Get GetArray() As Variant
    GetArray = This.Array
End Property

Public Property Get Range() As Range
    Set Range = This.Range
End Property

Public Property Set Range(ByVal vNewValue As Range)
    If vNewValue Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Description:=EMPTY_TABLE_MESSAGE
        Exit Property
    End If
    If Not IsRangeASingleColumn(vNewValue) Then
        Err.Raise Number:=vbObjectError + 1, Description:=DISJOINT_RANGE_MESSAGE
        Exit Property
    End If
    Set This.Range = vNewValue
    LoadRange vNewValue
End Property

Public Property Get RespectFilter() As Boolean
    RespectFilter = This.RespectFilter
End Property

Public Property Let RespectFilter(ByVal vNewValue As Boolean)
    This.RespectFilter = vNewValue
End Property

Public Property Get IncludeNumbersAsText() As Boolean
    IncludeNumbersAsText = This.IncludeNumbersAsText
End Property

Public Property Let IncludeNumbersAsText(ByVal vNewValue As Boolean)
    This.IncludeNumbersAsText = vNewValue
End Property

Public Property Get NonTextCount() As Long
    NonTextCount = This.NonTextCount
End Property

Public Property Get ErrorCount() As Long
    ErrorCount = This.ErrorCount
End Property

Public Property Get BlankCount() As Long
    BlankCount = This.BlankCount
End Property

Public Property Get Keys() As Collection
    Set Keys = This.Distinct
End Property

Public Property Get Count() As Long
    Count = This.Distinct.Count
End Property

Public Property Get UniqueKeys() As Collection
    Set UniqueKeys = This.Unique
End Property

Public Property Get IsDistinct() As Boolean
    IsDistinct = (This.Unique.Count = This.Distinct.Count)
End Property

' ---
Private Sub LoadRange(ByVal rng As Range)
    Debug.Assert rng.Columns.Count = 1
    Set This.Dict = CreateObject("Scripting.Dictionary")
    Dim a As Range
    Dim V As Variant
    Dim vv As Variant
    
    If This.RespectFilter Then
        On Error Resume Next
        Set rng = rng.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        If rng Is Nothing Then
            MsgBox NO_CELLS_AFTER_FILTER_MESSAGE
            Err.Raise 5, , LOAD_RANGE_FAILED
        End If
    End If
    
    Dim n As Long
    For Each a In rng.Areas
        V = a.Value2
        If VarType(V) = vbArray + vbVariant Then
            For Each vv In V
                ProcessVariants vv
                n = n + 1
                If n > This.Limit Then GoTo LimitReached
            Next vv
        Else
            ProcessVariants V
            n = n + 1
            If n > This.Limit Then GoTo LimitReached
        End If
    Next a
    
LimitReached:
    SetUnique
    SetDistinct
    
    This.Array = DisjointRangeToArray(rng)
End Sub

Private Sub ProcessVariants(ByVal Value As Variant)
    Select Case VarType(Value)
        Case vbString
            If Value = vbNullString Then
                This.BlankCount = This.BlankCount + 1
            Else
                AppendOrIncrement Value
            End If
        Case vbDouble
            If IncludeNumbersAsText = True Then
                AppendOrIncrement CStr(Value)
            Else
                This.NonTextCount = This.NonTextCount + 1
            End If
        Case vbError
            This.ErrorCount = This.ErrorCount + 1
        Case vbEmpty
            This.BlankCount = This.BlankCount + 1
        Case Else
            This.NonTextCount = This.NonTextCount + 1
    End Select
End Sub

Private Sub AppendOrIncrement(ByVal Value As String)
    If This.Dict.Exists(Value) Then
        This.Dict.Item(Value) = This.Dict.Item(Value) + 1
    Else
        This.Dict.Item(Value) = 1
    End If
End Sub

Private Sub SetUnique()
    Set This.Unique = New Collection
    Dim Key As Variant
    For Each Key In This.Dict.Keys
        If This.Dict.Item(Key) = 1 Then
            This.Unique.Add Key
        End If
    Next Key
End Sub

Private Sub SetDistinct()
    Set This.Distinct = New Collection
    Dim Key As Variant
    For Each Key In This.Dict.Keys
        This.Distinct.Add Key
    Next Key
End Sub

Public Function Trim(ByVal length As Long) As Collection
    Dim i As Integer
    
    If length > This.Distinct.Count Then
        Set Trim = This.Distinct
        Exit Function
    End If
    
    Set Trim = New Collection
    For i = 1 To length
        Trim.Add This.Distinct.Item(i)
    Next i
End Function

Public Function Find(ByVal Value As Variant) As Long
    Find = -1
    
    If VarType(Value) = vbError Then Exit Function
    If VarType(Value) = vbEmpty Then Exit Function
    If Value = vbNullString Then Exit Function
        
    Dim i As Long
    For i = LBound(This.Array) To UBound(This.Array)
        If VarType(This.Array(i)) <> vbError Then
            If This.Array(i) = Value Then
                Find = i
                Exit Function
            End If
        End If
    Next i
End Function

Public Function Exists(ByVal Value As Variant) As Boolean
    Exists = This.Dict.Exists(Value)
End Function

Public Function FromRange(ByVal rng As Range, Optional ByVal IncludeNumbers As Boolean = False, _
    Optional ByVal Filtered As Boolean = False, Optional ByVal Limit As Long) As KeyColumn
    If rng Is Nothing Then
        'Err.Raise vbObjectError + 3, Err.Source, "Cannot create KeyColumn from empty range (rng is nothing)"
        Exit Function
    End If
    Set FromRange = New KeyColumn
    FromRange.RespectFilter = Filtered
    FromRange.IncludeNumbersAsText = IncludeNumbers
    If Limit > 0 Then
        FromRange.Limit = Limit
    End If
    Set FromRange.Range = rng
End Function

Public Function FromColumn(ByVal lc As ListColumn, Optional ByVal IncludeNumbers As Boolean = False, _
    Optional ByVal Limit As Long) As KeyColumn
    Set FromColumn = FromRange(lc.DataBodyRange, IncludeNumbers, False, Limit)
End Function

Private Function IsRangeASingleColumn(ByVal Range As Range) As Boolean
    If Range Is Nothing Then Exit Function
    Dim Area As Range
    For Each Area In Range.Areas
        If Area.column <> Range.column Then
            Exit Function
        End If
    Next Area
    IsRangeASingleColumn = True
End Function

Private Function DisjointRangeToArray(ByVal rng As Range) As Variant
    Dim arr As Variant
    Dim Area As Range
    Dim vv As Variant
    Dim firstRow As Long
    Dim lastRow As Long
    Dim thisLastRow As Long
    
    firstRow = rng.row
    For Each Area In rng.Areas
        If firstRow > rng.row Then
            firstRow = rng.row
        End If
        thisLastRow = Area.row + Area.rows.Count - 1
        If thisLastRow > lastRow Then
            lastRow = thisLastRow
        End If
    Next Area
    
    ReDim arr(firstRow To thisLastRow)
    
    For Each Area In rng.Areas
        If VarType(Area) = vbArray + vbVariant Then
            For Each vv In Area
                arr(vv.row) = (vv) ' Removed CStr - this was converting errors to text
            Next vv
        Else
            arr(Area.row) = (Area)
        End If
    Next Area
    
    DisjointRangeToArray = arr
End Function

Private Sub Class_Initialize()
    This.Limit = MAX_LIMIT
End Sub

Public Sub DebugPrint()
    Dim Key As Variant
    For Each Key In This.Dict.Keys
        Debug.Print Key; " = "; This.Dict.Item(Key)
    Next Key
End Sub


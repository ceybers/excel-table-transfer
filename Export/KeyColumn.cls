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
'@Folder "Model.KeyColumn"
Option Explicit

Private Type TClsKeyColumn
    Dict As Object
    Distinct As Collection
    Unique As Collection
    rng As Range
    IncludeNumbersAsText As Boolean
    ErrorCount As Long
    BlankCount As Long
    NonTextCount As Long
    Array As Variant
    RespectFilter As Boolean
    Limit As Long
End Type

Private This As TClsKeyColumn

Public Property Get Limit() As Double
    Limit = This.Limit
End Property

Public Property Let Limit(ByVal rhs As Double)
    This.Limit = rhs
End Property

Public Property Get DistinctPercentage() As Double
    If This.rng.Cells.Count = 0 Then Exit Property
    DistinctPercentage = This.Distinct.Count / This.rng.Cells.Count
End Property

Public Property Get GetArray() As Variant
    GetArray = This.Array
End Property

Public Property Get Range() As Range
    Set Range = This.rng
End Property

Public Property Set Range(ByVal rhs As Range)
    If rhs Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Description:="Cannot create clsKeyColumn with a column in an empty table"
        Exit Property
    End If
    If Not IsRangeASingleColumn(rhs) Then
        Err.Raise Number:=vbObjectError + 1, Description:="Cannot create clsKeyColumn with a range that spans multiple columns"
        Exit Property
    End If
    Set This.rng = rhs
    LoadRange rhs
End Property

Public Property Get RespectFilter() As Boolean
    RespectFilter = This.RespectFilter
End Property

Public Property Let RespectFilter(ByVal rhs As Boolean)
    This.RespectFilter = rhs
End Property

Public Property Get IncludeNumbersAsText() As Boolean
    IncludeNumbersAsText = This.IncludeNumbersAsText
End Property

Public Property Let IncludeNumbersAsText(ByVal rhs As Boolean)
    This.IncludeNumbersAsText = rhs
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

Private Sub LoadRange(ByVal rng As Range)
    Debug.Assert rng.Columns.Count = 1
    Set This.Dict = CreateObject("Scripting.Dictionary")
    Dim a As Range
    Dim v As Variant
    Dim vv As Variant
    
    If This.RespectFilter Then
        On Error Resume Next
        Set rng = rng.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        If rng Is Nothing Then
            MsgBox "No cells were visible after filtering"
            Err.Raise 5, , "KeyColumn.LoadRange failed"
        End If
    End If
    
    Dim n As Long
    For Each a In rng.Areas
        v = a.Value2
        If VarType(v) = vbArray + vbVariant Then
            For Each vv In v
                ProcessVariants vv
                n = n + 1
                If n > This.Limit Then GoTo LimitReached
            Next vv
        Else
            ProcessVariants v
            n = n + 1
            If n > This.Limit Then GoTo LimitReached
        End If
    Next a
    
LimitReached:
    
    SetUnique
    SetDistinct
    
    This.Array = DisjointRangeToArray(rng)
End Sub

Private Sub ProcessVariants(ByVal v As Variant)
    Select Case VarType(v)
    Case vbString
        If v = vbNullString Then
            This.BlankCount = This.BlankCount + 1
        Else
            AppendOrIncrement v
        End If
    Case vbDouble
        If IncludeNumbersAsText = True Then
            AppendOrIncrement CStr(v)
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
    ' Debug.Print "AppendOrIncrement"; this.Dict.Count
    If This.Dict.Exists(Value) Then
        This.Dict.Item(Value) = This.Dict.Item(Value) + 1
    Else
        This.Dict.Item(Value) = 1
    End If
End Sub

Public Sub PrintKeys()
    Dim k As Variant
    For Each k In This.Dict.Keys
        Debug.Print k
    Next k
End Sub

Private Sub SetUnique()
    Set This.Unique = New Collection
    Dim key As Variant
    For Each key In This.Dict.Keys
        If This.Dict.Item(key) = 1 Then
            This.Unique.Add key
        End If
    Next key
End Sub

Private Sub SetDistinct()
    Set This.Distinct = New Collection
    Dim key As Variant
    For Each key In This.Dict.Keys
        This.Distinct.Add key
    Next key
End Sub

Public Property Get IsDistinct() As Boolean
    IsDistinct = (This.Unique.Count = This.Distinct.Count)
End Property

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
    'Dim result As Range
    'Set result = this.rng.Find(value)
    'If Not result Is Nothing Then
    '    Find = result.Row
    'End If
    'Debug.Print VarType(value)
    If VarType(Value) = vbError Then Exit Function
    If VarType(Value) = vbEmpty Then Exit Function
    If Value = vbNullString Then Exit Function
    
    'Debug.Print "Find('" & value & "')"
        
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

Public Function FromRange(ByVal rng As Range, Optional ByVal IncludeNumbers As Boolean = False, Optional ByVal Filtered As Boolean = False, Optional ByVal Limit As Long) As KeyColumn
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

Public Function FromColumn(ByVal lc As ListColumn, Optional ByVal IncludeNumbers As Boolean = False, Optional ByVal Limit As Long) As KeyColumn
    Set FromColumn = FromRange(lc.DataBodyRange, IncludeNumbers, False, Limit)
End Function

Private Function IsRangeASingleColumn(ByVal rng As Range) As Boolean
    If rng Is Nothing Then Exit Function
    Dim area As Range
    For Each area In rng.Areas
        If area.column <> rng.column Then
            Exit Function
        End If
    Next area
    IsRangeASingleColumn = True
End Function

Private Function DisjointRangeToArray(ByVal rng As Range) As Variant
    Dim arr As Variant
    Dim area As Range
    Dim vv As Variant
    Dim firstRow As Long
    Dim lastRow As Long
    Dim thisLastRow As Long
    
    firstRow = rng.row
    For Each area In rng.Areas
        If firstRow > rng.row Then
            firstRow = rng.row
        End If
        thisLastRow = area.row + area.rows.Count - 1
        If thisLastRow > lastRow Then
            lastRow = thisLastRow
        End If
    Next area
    
    ReDim arr(firstRow To thisLastRow)
    
    For Each area In rng.Areas
        If VarType(area) = vbArray + vbVariant Then
            For Each vv In area
                arr(vv.row) = (vv)               ' Removed CStr - this was converting errors to text
            Next vv
        Else
            arr(area.row) = (area)
        End If
    Next area
    
    DisjointRangeToArray = arr
End Function

Private Sub class_initialize()
    This.Limit = 2147483647
End Sub


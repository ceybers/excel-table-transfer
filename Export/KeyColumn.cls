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
'@Folder("KeyColumn")
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
End Type

Private This As TClsKeyColumn

Public Property Get GetArray() As Variant
    GetArray = This.Array
End Property

Public Property Get Range() As Range
    Set Range = This.rng
End Property

Public Property Set Range(ByVal RHS As Range)
    If Not IsRangeASingleColumn(RHS) Then
        Err.Raise Number:=vbObjectError + 1, Description:="Cannot create clsKeyColumn with a range that spans multiple columns"
        Exit Property
    End If
    Set This.rng = RHS
    LoadRange RHS
End Property

Public Property Get IncludeNumbersAsText() As Boolean
    IncludeNumbersAsText = This.IncludeNumbersAsText
End Property

Public Property Let IncludeNumbersAsText(ByVal RHS As Boolean)
    This.IncludeNumbersAsText = RHS
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
    
    For Each a In rng.Areas
        v = a.Value2
        If VarType(v) = vbArray + vbVariant Then
            For Each vv In v
                ProcessVariants vv
            Next vv
        Else
            ProcessVariants v
        End If
    Next a
    
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

Private Sub AppendOrIncrement(ByVal value As String)
    If This.Dict.Exists(value) Then
        This.Dict.Item(value) = This.Dict.Item(value) + 1
    Else
        This.Dict.Item(value) = 1
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

Public Function Find(ByVal value As Variant) As Long
    Find = -1
    'Dim result As Range
    'Set result = this.rng.Find(value)
    'If Not result Is Nothing Then
    '    Find = result.Row
    'End If

    If VarType(value) = vbError Then Exit Function
    If VarType(value) = vbEmpty Then Exit Function
    If value = vbNullString Then Exit Function
    
    'Debug.Print "Find('" & value & "')"
        
    Dim i As Long
    For i = LBound(This.Array) To UBound(This.Array)
        If VarType(This.Array(i)) <> vbError Then
            If This.Array(i) = value Then
                Find = i
                Exit Function
            End If
        End If
    Next i
End Function

Public Function Exists(ByVal value As Variant) As Boolean
    Exists = This.Dict.Exists(value)
End Function

Public Function FromRange(ByVal rng As Range, Optional ByVal IncludeNumbers As Boolean = False) As KeyColumn
    Set FromRange = New KeyColumn
    FromRange.IncludeNumbersAsText = IncludeNumbers
    Set FromRange.Range = rng
End Function

Public Function FromColumn(ByVal lc As ListColumn, Optional ByVal IncludeNumbers As Boolean = False) As KeyColumn
    Set FromColumn = FromRange(lc.DataBodyRange, IncludeNumbers)
End Function

Private Function IsRangeASingleColumn(ByVal rng As Range) As Boolean
    Dim area As Range
    For Each area In rng.Areas
        If area.Column <> rng.Column Then
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
    
    firstRow = rng.Row
    For Each area In rng.Areas
        If firstRow > rng.Row Then
            firstRow = rng.Row
        End If
        thisLastRow = area.Row + area.Rows.Count - 1
        If thisLastRow > lastRow Then
            lastRow = thisLastRow
        End If
    Next area
    
    ReDim arr(firstRow To thisLastRow)
    
    For Each area In rng.Areas
        If VarType(area) = vbArray + vbVariant Then
            For Each vv In area
                arr(vv.Row) = vv
            Next vv
        Else
            arr(area.Row) = area
        End If
    Next area
    
    DisjointRangeToArray = arr
End Function


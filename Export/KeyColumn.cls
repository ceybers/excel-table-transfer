VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "MVVM.Model.KeyColumn"
Option Explicit

Private Const DEFAULT_LIMIT As Long = 2147483647 ' TODO Rename to MAX_VAL_LONG

Private Type TClsKeyColumn
    Name As String
    
    Dictionary As Object
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

Public Property Get Name() As String
    Name = This.Name
End Property

Public Property Let Name(ByVal vNewValue As String)
    This.Name = vNewValue
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

Public Property Set Range(ByVal RHS As Range)
    If RHS Is Nothing Then
        Err.Raise Number:=vbObjectError + 1, Description:=StringConstants.ERR_MSG_KEYCOL_EMPTY_TABLE
        Exit Property
    End If
    If Not IsRangeASingleColumn(RHS) Then
        Err.Raise Number:=vbObjectError + 1, Description:=ERR_MSG_MULTIPLE_COLUMNS
        Exit Property
    End If
    Set This.Range = RHS
    LoadRange RHS
End Property

Public Property Get Keys() As Collection
    Set Keys = This.Distinct
End Property

Public Property Get UniqueKeys() As Collection
    Set UniqueKeys = This.Unique
End Property

Public Property Get Count() As Long
    Count = This.Distinct.Count
End Property

Public Property Get IsDistinct() As Boolean
    IsDistinct = (This.Unique.Count = This.Distinct.Count)
End Property

Public Property Get IncludeNumbersAsText() As Boolean
    IncludeNumbersAsText = This.IncludeNumbersAsText
End Property

Public Property Let IncludeNumbersAsText(ByVal RHS As Boolean)
    This.IncludeNumbersAsText = RHS
End Property

Public Property Get RespectFilter() As Boolean
    RespectFilter = This.RespectFilter
End Property

Public Property Let RespectFilter(ByVal RHS As Boolean)
    This.RespectFilter = RHS
End Property

Public Property Get ErrorCount() As Long
    ErrorCount = This.ErrorCount
End Property

Public Property Get BlankCount() As Long
    BlankCount = This.BlankCount
End Property

Public Property Get NonTextCount() As Long
    NonTextCount = This.NonTextCount
End Property

Public Property Get Limit() As Long ' TODO CHK was Double
    Limit = This.Limit
End Property

Public Property Let Limit(ByVal RHS As Long) ' TODO CHK was Double
    This.Limit = RHS
End Property

Private Sub Class_Initialize()
    This.Limit = DEFAULT_LIMIT
End Sub

Public Function FromColumn(ByVal ListColumn As ListColumn, Optional ByVal IncludeNumbers As Boolean = False, Optional ByVal Limit As Long) As KeyColumn
    Set FromColumn = FromRange(ListColumn.DataBodyRange, IncludeNumbers, False, Limit)
    FromColumn.Name = ListColumn.Name
End Function

Public Function FromRange(ByVal Range As Range, Optional ByVal IncludeNumbers As Boolean = False, _
    Optional ByVal Filtered As Boolean = False, Optional ByVal Limit As Long) As KeyColumn
    If Range Is Nothing Then
        'Err.Raise vbObjectError + 3, Err.Source, ERR_MSG_EMPTY_RANGE
        Exit Function
    End If
    
    Dim Result As KeyColumn
    Set Result = New KeyColumn
    
    With Result
        .RespectFilter = Filtered
        .IncludeNumbersAsText = IncludeNumbers
        If Limit > 0 Then .Limit = Limit
        Set .Range = Range
    End With
    
    Set FromRange = Result
End Function

Public Function Exists(ByVal Value As Variant) As Boolean
    Exists = This.Dictionary.Exists(Value)
End Function

Public Function Find(ByVal Criteria As Variant) As Long
    Find = -1
    
    If VarType(Criteria) = vbError Then Exit Function
    If VarType(Criteria) = vbEmpty Then Exit Function
    If Criteria = vbNullString Then Exit Function
    
    Dim i As Long
    For i = LBound(This.Array) To UBound(This.Array)
        If VarType(This.Array(i)) <> vbError Then
            If This.Array(i) = Criteria Then
                Find = i
                Exit Function
            End If
        End If
    Next i
End Function

Public Function Trim(ByVal Length As Long) As Collection
    If Length >= This.Distinct.Count Then
        Set Trim = This.Distinct
        Exit Function
    End If
    
    Dim Result As Collection
    Set Result = New Collection
    
    Dim i As Long
    For i = 1 To Length
        Result.Add This.Distinct.Item(i)
    Next i
    
    Set Trim = Result
End Function

Public Sub PrintKeys()
    Dim Key As Variant
    For Each Key In This.Dictionary.Keys
        Debug.Print Key
    Next Key
End Sub

Private Sub LoadRange(ByVal Range As Range)
    Debug.Assert Range.Columns.Count = 1
    
    Set This.Dictionary = CreateObject("Scripting.Dictionary")
    
    Dim FilteredRange As Range
    Set FilteredRange = Range
    
    If This.RespectFilter Then
        On Error Resume Next
        Set FilteredRange = Range.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        If FilteredRange Is Nothing Then
            'MsgBox "No cells were visible after filtering", vbCritical + vbOKOnly, "No visible cells!"
            Err.Raise 5, , ERR_MSG_NO_VISIBLE_CELLS
        End If
    End If
    
    Dim Count As Long
    
    Dim Area As Range
    For Each Area In FilteredRange.Areas
        Dim ValueVariant As Variant
        ValueVariant = Area.Value2
        
        If VarType(ValueVariant) = vbArray + vbVariant Then
            Dim ValueItem As Variant
            For Each ValueItem In ValueVariant
                ProcessVariants ValueItem
                Count = Count + 1
                If Count > This.Limit Then GoTo LimitReached
            Next ValueItem
        Else
            ProcessVariants ValueVariant
            Count = Count + 1
            If Count > This.Limit Then GoTo LimitReached
        End If
    Next Area
    
LimitReached:
    SetUnique
    SetDistinct
    
    This.Array = DisjointRangeToArray(FilteredRange)
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
    If This.Dictionary.Exists(Value) Then
        This.Dictionary.Item(Value) = This.Dictionary.Item(Value) + 1
    Else
        This.Dictionary.Item(Value) = 1
    End If
End Sub

Private Function IsRangeASingleColumn(ByVal Range As Range) As Boolean
    If Range Is Nothing Then Exit Function
    
    Dim Area As Range
    For Each Area In Range.Areas
        If Area.Column <> Range.Column Then
            Exit Function
        End If
    Next Area
    
    IsRangeASingleColumn = True
End Function

Private Function DisjointRangeToArray(ByVal Range As Range) As Variant
    Dim FirstRow As Long
    FirstRow = Range.Row
    
    Dim LastRow As Long
    'LastRow = FirstRow ' TODO CHK What is the initial value?
    
    Dim Area As Range
    For Each Area In Range.Areas
        If FirstRow > Range.Row Then
            FirstRow = Range.Row
        End If
        
        Dim ThisLastRow As Long
        ThisLastRow = Area.Row + Area.rows.Count - 1
        
        If ThisLastRow > LastRow Then
            LastRow = ThisLastRow
        End If
    Next Area
    
    Dim ResultArray As Variant
    ReDim ResultArray(FirstRow To ThisLastRow)
    
    For Each Area In Range.Areas
        If VarType(Area) = (vbArray + vbVariant) Then
            Dim ValueVariant As Variant
            For Each ValueVariant In Area
                ResultArray(ValueVariant.Row) = (ValueVariant)
            Next ValueVariant
        Else
            ResultArray(Area.Row) = (Area.Value2)
        End If
    Next Area
    
    DisjointRangeToArray = ResultArray
End Function

Private Sub SetUnique()
    Set This.Unique = New Collection
    
    Dim Key As Variant
    For Each Key In This.Dictionary.Keys
        If This.Dictionary.Item(Key) = 1 Then
            This.Unique.Add Key
        End If
    Next Key
End Sub

Private Sub SetDistinct()
    Set This.Distinct = New Collection
    
    Dim Key As Variant
    For Each Key In This.Dictionary.Keys
        This.Distinct.Add Key
    Next Key
End Sub



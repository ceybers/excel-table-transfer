VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyColumnComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule UnassignedVariableUsage
'@PredeclaredId
'@Folder "Model.KeyColumn"
Option Explicit

Private Type TClsCompareKeyColumns
    lhs As KeyColumn
    rhs As KeyColumn
    LeftOnly As Collection
    Intersection As Collection
    RightOnly As Collection
    Dict As Object
End Type

Private This As TClsCompareKeyColumns

Public Property Get Score() As Double
    Dim total As Long
    total = This.LeftOnly.Count + This.Intersection.Count + This.RightOnly.Count
    
    Score = 0
    
    If total < 5 Then Exit Property
    
    If total > 0 Then
        Score = This.Intersection.Count / total
    End If
    
    'Debug.Print "Score: "; score
End Property

Public Property Get lhs() As KeyColumn
    Set lhs = This.lhs
End Property

Public Property Set lhs(ByVal rhs As KeyColumn)
    Set This.lhs = rhs
    TryDoWork
End Property

Public Property Get rhs() As KeyColumn
    Set rhs = This.rhs
End Property

Public Property Set rhs(ByVal rhs As KeyColumn)
    Set This.rhs = rhs
    TryDoWork
End Property

Public Property Get LeftOnly() As Collection
    Set LeftOnly = This.LeftOnly
End Property

Public Property Get Intersection() As Collection
    Set Intersection = This.Intersection
End Property

Public Property Get RightOnly() As Collection
    Set RightOnly = This.RightOnly
End Property

Public Property Get IsSubsetLHS() As Boolean
    IsSubsetLHS = This.LeftOnly.Count = 0
End Property

Public Property Get IsSubsetRHS() As Boolean
    IsSubsetRHS = This.RightOnly.Count = 0
End Property

Public Property Get IsMatch() As Boolean
    IsMatch = IsSubsetLHS And IsSubsetRHS
End Property

' Methods
Private Sub TryDoWork()
    If This.lhs Is Nothing Then Exit Sub
    If This.rhs Is Nothing Then Exit Sub
    DoWork
End Sub

Private Sub DoWork()
    Set This.Dict = CreateObject("Scripting.Dictionary")
    Dim key As Variant
    For Each key In lhs.Keys
        Upsert key, 1
    Next
    For Each key In rhs.Keys
        Upsert key, 2
    Next
    
    Set This.Intersection = New Collection
    Set This.LeftOnly = New Collection
    Set This.RightOnly = New Collection
    
    Dim val As Integer
    
    For Each key In This.Dict.Keys
        Select Case This.Dict.Item(key)
        Case 1
            This.LeftOnly.Add key
        Case 2
            This.RightOnly.Add key
        Case 3
            This.Intersection.Add key
        End Select
    Next key
End Sub

Private Sub Upsert(ByVal key As String, ByVal Value As Integer)
    If This.Dict.Exists(key) Then
        This.Dict.Item(key) = This.Dict.Item(key) + Value
    Else
        This.Dict.Item(key) = Value
    End If
End Sub

Public Function Create(ByVal lhs As KeyColumn, ByVal rhs As KeyColumn) As KeyColumnComparer
    Set Create = New KeyColumnComparer
    Set Create.lhs = lhs
    Set Create.rhs = rhs
End Function

Public Function Map(Optional ByVal FilterSource As Boolean = False, Optional ByVal FilterDestination As Boolean = False) As Variant
    'Debug.Assert Not this.LHS Is Nothing
    'Debug.Assert Not this.rhs Is Nothing
    If This.lhs Is Nothing Or This.rhs Is Nothing Then
        'Err.Raise vbObjectError + 4, Err.Source, "Cannot Map KeyColumns with an empty table"
        Exit Function
    End If
    
    Dim Result As Variant
    Dim arr As Variant
    Dim i As Long
    
    Dim lhsFlt As KeyColumn
    Set lhsFlt = This.lhs
    If FilterSource Then
        Set lhsFlt = KeyColumn.FromRange(This.lhs.Range, This.lhs.IncludeNumbersAsText, True)
    End If
    
    Dim rhsFlt As KeyColumn
    Set rhsFlt = This.rhs
    If FilterDestination Then
        Set rhsFlt = KeyColumn.FromRange(This.rhs.Range, This.rhs.IncludeNumbersAsText, True)
    End If
    
    arr = rhsFlt.GetArray
    ReDim Result(LBound(arr) To UBound(arr))
    
    For i = LBound(arr) To UBound(arr)
        Result(i) = lhsFlt.Find(arr(i))
    Next i
    
    Map = Result
End Function


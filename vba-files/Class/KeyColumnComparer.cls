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
'@Folder "MVVM.Models.KeyColumn"
Option Explicit

Private Type TClsCompareKeyColumns
    lhs As KeyColumn
    RHS As KeyColumn
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

Public Property Set lhs(ByVal RHS As KeyColumn)
    Set This.lhs = RHS
    TryDoWork
End Property

Public Property Get RHS() As KeyColumn
    Set RHS = This.RHS
End Property

Public Property Set RHS(ByVal RHS As KeyColumn)
    Set This.RHS = RHS
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
    If This.RHS Is Nothing Then Exit Sub
    DoWork
End Sub

Private Sub DoWork()
    Set This.Dict = CreateObject("Scripting.Dictionary")
    Dim Key As Variant
    For Each Key In lhs.Keys
        Upsert Key, 1
    Next
    For Each Key In RHS.Keys
        Upsert Key, 2
    Next
    
    Set This.Intersection = New Collection
    Set This.LeftOnly = New Collection
    Set This.RightOnly = New Collection
    
    Dim val As Integer
    
    For Each Key In This.Dict.Keys
        Select Case This.Dict.Item(Key)
            Case 1
                This.LeftOnly.Add Key
            Case 2
                This.RightOnly.Add Key
            Case 3
                This.Intersection.Add Key
        End Select
    Next Key
End Sub

Private Sub Upsert(ByVal Key As String, ByVal Value As Integer)
    If This.Dict.Exists(Key) Then
        This.Dict.Item(Key) = This.Dict.Item(Key) + Value
    Else
        This.Dict.Item(Key) = Value
    End If
End Sub

Public Function Create(ByVal Src As KeyColumn, ByVal Dst As KeyColumn) As KeyColumnComparer
    Set Create = New KeyColumnComparer
    Set Create.lhs = Src
    Set Create.RHS = Dst
End Function

Public Function Map(Optional ByVal FilterSource As Boolean = False, Optional ByVal FilterDestination As Boolean = False) As Variant
    'Debug.Assert Not this.LHS Is Nothing
    'Debug.Assert Not this.rhs Is Nothing
    If This.lhs Is Nothing Or This.RHS Is Nothing Then
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
    Set rhsFlt = This.RHS
    If FilterDestination Then
        Set rhsFlt = KeyColumn.FromRange(This.RHS.Range, This.RHS.IncludeNumbersAsText, True)
    End If
    
    arr = rhsFlt.GetArray
    ReDim Result(LBound(arr) To UBound(arr))
    
    For i = LBound(arr) To UBound(arr)
        Result(i) = lhsFlt.Find(arr(i))
    Next i
    
    Map = Result
End Function

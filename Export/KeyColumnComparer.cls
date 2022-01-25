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
'@Folder("KeyColumn")
Option Explicit

Private Type TClsCompareKeyColumns
    LHS As KeyColumn
    RHS As KeyColumn
    LeftOnly As Collection
    Intersection As Collection
    RightOnly As Collection
    Dict As Object
End Type

Private This As TClsCompareKeyColumns

Public Property Get LHS() As KeyColumn
    Set LHS = This.LHS
End Property

Public Property Set LHS(ByVal RHS As KeyColumn)
    Set This.LHS = RHS
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
    If This.LHS Is Nothing Then Exit Sub
    If This.RHS Is Nothing Then Exit Sub
    DoWork
End Sub

Private Sub DoWork()
    Set This.Dict = CreateObject("Scripting.Dictionary")
    Dim key As Variant
    For Each key In LHS.Keys
        Upsert key, 1
    Next
    For Each key In RHS.Keys
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

Private Sub Upsert(ByVal key As String, ByVal value As Integer)
    If This.Dict.Exists(key) Then
        This.Dict.Item(key) = This.Dict.Item(key) + value
    Else
        This.Dict.Item(key) = value
    End If
End Sub

Public Function Create(ByVal LHS As KeyColumn, ByVal RHS As KeyColumn) As KeyColumnComparer
    Set Create = New KeyColumnComparer
    Set Create.LHS = LHS
    Set Create.RHS = RHS
End Function

Public Function Map() As Variant
    Dim Result As Variant
    Dim arr As Variant
    Dim i As Long
    
    arr = This.RHS.GetArray
    ReDim Result(LBound(arr) To UBound(arr))
    
    For i = LBound(arr) To UBound(arr)
        Result(i) = This.LHS.Find(arr(i))
    Next i
    
    Map = Result
End Function

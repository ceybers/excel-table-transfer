VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnPairs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Model.ColumnPairs"
Option Explicit

Private Type TClsColumnPairs
    Item As Collection
    ConstrainRHS As Boolean
End Type

Private This As TClsColumnPairs

'@Enumerator
Public Property Get Item() As Collection
Attribute Item.VB_UserMemId = -4
    Set Item = This.Item
End Property

Public Property Get Count() As Long
    Count = This.Item.Count
End Property

Public Property Get ConstrainRHS() As Boolean
    ConstrainRHS = This.ConstrainRHS
End Property

Public Property Let ConstrainRHS(ByVal RHS As Boolean)
    This.ConstrainRHS = RHS
End Property

Private Sub Class_Initialize()
    Set This.Item = New Collection
    This.ConstrainRHS = True
End Sub

Public Sub Add(ByVal ColumnPair As ColumnPair)
    If Exists(ColumnPair) Then Exit Sub
    
    This.Item.Add ColumnPair
End Sub

Public Sub AddOrReplace(ByVal ColumnPair As ColumnPair)
    If Not Exists(ColumnPair) Then
        This.Item.Add ColumnPair
        Exit Sub
    End If
    
    Dim FoundPair As ColumnPair
    If ConstrainRHS Then
        Set FoundPair = GetPair(RHS:=ColumnPair.RHS)
    Else
        Set FoundPair = GetPair(ColumnPair.LHS, ColumnPair.RHS)
    End If
    
    Set FoundPair.LHS = ColumnPair.LHS
    Set FoundPair.RHS = ColumnPair.RHS
End Sub

Public Sub Remove(ByVal ColumnPair As ColumnPair)
    Debug.Assert Not ColumnPair Is Nothing
    
    Dim FoundPair As ColumnPair
    Set FoundPair = GetPair(ColumnPair.LHS, ColumnPair.RHS)
    
    Dim idx As Long
    idx = GetIndexOfColumnPair(FoundPair)
    
    If idx >= 0 Then
        This.Item.Remove idx
    End If
End Sub

Private Function GetIndexOfColumnPair(ByVal ColumnPair As ColumnPair) As Long
    Debug.Assert Not ColumnPair Is Nothing
    
    Dim i As Long
    For i = 1 To This.Item.Count
        If This.Item.Item(i).Equals(ColumnPair) Or (This.ConstrainRHS And (This.Item.Item(i).RHS Is ColumnPair.RHS)) Then
            GetIndexOfColumnPair = i
            Exit Function
        End If
    Next i
    
    GetIndexOfColumnPair = -1
End Function

Public Sub Clear()
    CollectionHelpers.CollectionClear This.Item
End Sub

Public Function Exists(ByVal ColumnPair As ColumnPair) As Boolean
    Dim RHS As ColumnPair
    For Each RHS In This.Item
        If ColumnPair.Equals(RHS) Then
            Exists = True
            Exit Function
        ElseIf ConstrainRHS And (ColumnPair.RHS Is RHS.RHS) Then
            Exists = True
            Exit Function
        End If
    Next RHS
End Function

Public Function GetPair(Optional ByVal LHS As ListColumn, Optional ByVal RHS As ListColumn) As ColumnPair
    If (LHS Is Nothing) And (RHS Is Nothing) Then Exit Function

    Dim ColumnPair As ColumnPair
    For Each ColumnPair In This.Item
        If (RHS Is Nothing) And (ColumnPair.LHS Is LHS) Then
            Set GetPair = ColumnPair
            Exit Function
        ElseIf (LHS Is Nothing) And (ColumnPair.RHS Is RHS) Then
            Set GetPair = ColumnPair
            Exit Function
        ElseIf (ColumnPair.LHS Is LHS) And (ColumnPair.RHS Is RHS) Then
            Set GetPair = ColumnPair
            Exit Function
        End If
    Next ColumnPair
End Function

Public Function TryGetPair(ByRef ColumnPair As ColumnPair, Optional ByVal LHS As ListColumn, Optional ByVal RHS As ListColumn) As Boolean
    Dim Result As ColumnPair
    Set Result = GetPair(LHS, RHS)
    
    Set ColumnPair = Result
    
    If Not Result Is Nothing Then
        TryGetPair = True
    End If
End Function

Public Function IsMappedRHS(ByVal ListColumn As ListColumn) As Boolean
    Dim ColumnPair As ColumnPair
    For Each ColumnPair In This.Item
        If ColumnPair.RHS Is ListColumn Then
            IsMappedRHS = True
            Exit Function
        End If
    Next ColumnPair
End Function

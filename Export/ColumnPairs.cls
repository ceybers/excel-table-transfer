VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnPairs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ColumnPairs"
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

Public Sub AddOrReplace(ByVal colPair As ColumnPair)
    Dim foundPair As ColumnPair
    If Exists(colPair) Then
        If ConstrainRHS Then
            Set foundPair = GetPair(RHS:=colPair.RHS)
        Else
            Set foundPair = GetPair(colPair.lhs, colPair.RHS)
        End If
        Set foundPair.lhs = colPair.lhs
        Set foundPair.RHS = colPair.RHS
    Else
        This.Item.Add colPair
    End If
End Sub

Public Sub Add(ByVal colPair As ColumnPair)
    If Exists(colPair) Then Exit Sub
    This.Item.Add colPair
End Sub

Public Sub Remove(ByVal colPair As ColumnPair)
    Debug.Assert Not colPair Is Nothing
    Dim idx As Long
    Dim match As ColumnPair
    Set match = GetPair(colPair.lhs, colPair.RHS)
    idx = GetIndexOfColumnPair(match)
    If idx > -1 Then
        This.Item.Remove idx
    End If
End Sub

Private Function GetIndexOfColumnPair(ByVal colPair As ColumnPair) As Long
    Debug.Assert Not colPair Is Nothing
    Dim i As Long
    For i = 1 To This.Item.Count
        If This.Item(i).Equals(colPair) Or (This.ConstrainRHS And (This.Item(i).RHS Is colPair.RHS)) Then
            GetIndexOfColumnPair = i
            Exit Function
        End If
    Next i
    GetIndexOfColumnPair = -1
End Function

Public Sub Clear()
    Dim i As Long
    For i = This.Item.Count To 1 Step -1
        This.Item.Remove i
    Next i
End Sub

Public Function Exists(ByVal colPair As ColumnPair) As Boolean
    Dim RHS As ColumnPair
    For Each RHS In This.Item
        If colPair.Equals(RHS) Then
            Exists = True
            Exit Function
        ElseIf ConstrainRHS And colPair.RHS Is RHS.RHS Then
            Exists = True
            Exit Function
        End If
    Next RHS
End Function

Public Function GetPair(Optional ByVal lhs As ListColumn, Optional ByVal RHS As ListColumn) As ColumnPair
    Dim colPair As ColumnPair
    
    If (lhs Is Nothing) And (RHS Is Nothing) Then Exit Function

    For Each colPair In This.Item
        If (RHS Is Nothing) And colPair.lhs Is lhs Then
            Set GetPair = colPair
            Exit Function
        ElseIf (lhs Is Nothing) And colPair.RHS Is RHS Then
            Set GetPair = colPair
            Exit Function
        ElseIf colPair.lhs Is lhs And colPair.RHS Is RHS Then
            Set GetPair = colPair
            Exit Function
        End If
    Next colPair
End Function

Public Function TryGetPair(ByRef colPair As ColumnPair, Optional ByVal lhs As ListColumn, Optional ByVal RHS As ListColumn) As Boolean
    Dim Result As ColumnPair
    Set Result = GetPair(lhs, RHS)
    Set colPair = Result
    If Not Result Is Nothing Then TryGetPair = True
End Function

Public Function IsMappedRHS(ByVal lc As ListColumn) As Boolean
    Dim colPair As ColumnPair
    For Each colPair In This.Item
        If colPair.RHS Is lc Then
            IsMappedRHS = True
            Exit Function
        End If
    Next colPair
End Function

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

Public Property Let ConstrainRHS(ByVal rhs As Boolean)
    This.ConstrainRHS = rhs
End Property

Private Sub Class_Initialize()
    Set This.Item = New Collection
    This.ConstrainRHS = True
End Sub

Public Sub AddOrReplace(ByVal colPair As ColumnPair)
    Dim foundPair As ColumnPair
    If Exists(colPair) Then
        If ConstrainRHS Then
            Set foundPair = GetPair(rhs:=colPair.rhs)
        Else
            Set foundPair = GetPair(colPair.lhs, colPair.rhs)
        End If
        Set foundPair.lhs = colPair.lhs
        Set foundPair.rhs = colPair.rhs
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
    Set match = GetPair(colPair.lhs, colPair.rhs)
    idx = GetIndexOfColumnPair(match)
    If idx > -1 Then
        This.Item.Remove idx
    End If
End Sub

Private Function GetIndexOfColumnPair(ByVal colPair As ColumnPair) As Long
    Debug.Assert Not colPair Is Nothing
    Dim i As Long
    For i = 1 To This.Item.Count
        If This.Item(i).Equals(colPair) Or (This.ConstrainRHS And (This.Item(i).rhs Is colPair.rhs)) Then
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
    Dim rhs As ColumnPair
    For Each rhs In This.Item
        If colPair.Equals(rhs) Then
            Exists = True
            Exit Function
        ElseIf ConstrainRHS And colPair.rhs Is rhs.rhs Then
            Exists = True
            Exit Function
        End If
    Next rhs
End Function

Public Function GetPair(Optional ByVal lhs As ListColumn, Optional ByVal rhs As ListColumn) As ColumnPair
    Dim colPair As ColumnPair
    
    If (lhs Is Nothing) And (rhs Is Nothing) Then Exit Function

    For Each colPair In This.Item
        If (rhs Is Nothing) And colPair.lhs Is lhs Then
            Set GetPair = colPair
            Exit Function
        ElseIf (lhs Is Nothing) And colPair.rhs Is rhs Then
            Set GetPair = colPair
            Exit Function
        ElseIf colPair.lhs Is lhs And colPair.rhs Is rhs Then
            Set GetPair = colPair
            Exit Function
        End If
    Next colPair
End Function

Public Function TryGetPair(ByRef colPair As ColumnPair, Optional ByVal lhs As ListColumn, Optional ByVal rhs As ListColumn) As Boolean
    Dim result As ColumnPair
    Set result = GetPair(lhs, rhs)
    Set colPair = result
    If Not result Is Nothing Then TryGetPair = True
End Function

Public Function IsMappedRHS(ByVal lc As ListColumn) As Boolean
    Dim colPair As ColumnPair
    For Each colPair In This.Item
        If colPair.rhs Is lc Then
            IsMappedRHS = True
            Exit Function
        End If
    Next colPair
End Function

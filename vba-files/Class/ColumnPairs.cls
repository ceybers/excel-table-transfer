VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColumnPairs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.Models.ColumnPairs"
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

Public Property Let ConstrainRHS(ByVal vNewValue As Boolean)
    This.ConstrainRHS = vNewValue
End Property

Private Sub Class_Initialize()
    Set This.Item = New Collection
    This.ConstrainRHS = True
End Sub

Public Sub AddOrReplace(ByVal ColumnPair As ColumnPair)
    Dim FoundPair As ColumnPair
    If Exists(ColumnPair) Then
        If ConstrainRHS Then
            Set FoundPair = GetPair(Dst:=ColumnPair.DestinationListColumn)
        Else
            Set FoundPair = GetPair(ColumnPair.SourceListColumn, ColumnPair.DestinationListColumn)
        End If
        Set FoundPair.SourceListColumn = ColumnPair.SourceListColumn
        Set FoundPair.DestinationListColumn = ColumnPair.DestinationListColumn
    Else
        This.Item.Add ColumnPair
    End If
End Sub

Public Sub Add(ByVal ColumnPair As ColumnPair)
    If Exists(ColumnPair) Then Exit Sub
    This.Item.Add ColumnPair
End Sub

Public Sub Remove(ByVal ColumnPair As ColumnPair)
    Debug.Assert Not ColumnPair Is Nothing
    Dim idx As Long
    Dim MatchPair As ColumnPair
    Set MatchPair = GetPair(ColumnPair.SourceListColumn, ColumnPair.DestinationListColumn)
    idx = GetIndexOfColumnPair(MatchPair)
    If idx > -1 Then
        This.Item.Remove idx
    End If
End Sub

Private Function GetIndexOfColumnPair(ByVal ColumnPair As ColumnPair) As Long
    Debug.Assert Not ColumnPair Is Nothing
    
    Dim i As Long
    For i = 1 To This.Item.Count
        If This.Item(i).Equals(ColumnPair) Or _
            (This.ConstrainRHS And (This.Item(i).RHS Is ColumnPair.DestinationListColumn)) Then
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

Public Function Exists(ByVal ColumnPair As ColumnPair) As Boolean
    Dim ThisColumnPair As ColumnPair
    For Each ThisColumnPair In This.Item
        If ColumnPair.Equals(ThisColumnPair) Then
            Exists = True
            Exit Function
        ElseIf ConstrainRHS And (ColumnPair.DestinationListColumn Is ThisColumnPair.DestinationListColumn) Then
            Exists = True
            Exit Function
        End If
    Next ThisColumnPair
End Function

Public Function GetPair(Optional ByVal Src As ListColumn, Optional ByVal Dst As ListColumn) As ColumnPair
    Dim ColumnPair As ColumnPair
    
    If (Src Is Nothing) And (Dst Is Nothing) Then Exit Function

    For Each ColumnPair In This.Item
        If (Dst Is Nothing) And (ColumnPair.SourceListColumn Is Src) Then
            Set GetPair = ColumnPair
            Exit Function
        ElseIf (Src Is Nothing) And (ColumnPair.DestinationListColumn Is Dst) Then
            Set GetPair = ColumnPair
            Exit Function
        ElseIf (ColumnPair.SourceListColumn Is Src) And (ColumnPair.DestinationListColumn Is Dst) Then
            Set GetPair = ColumnPair
            Exit Function
        End If
    Next ColumnPair
End Function

Public Function TryGetPair(ByRef ColumnPair As ColumnPair, Optional ByVal Src As ListColumn, Optional ByVal Dst As ListColumn) As Boolean
    Dim Result As ColumnPair
    Set Result = GetPair(Src, Dst)
    Set ColumnPair = Result
    If Not Result Is Nothing Then TryGetPair = True
End Function

Public Function IsDestinationMapped(ByVal Src As ListColumn) As Boolean
    Dim ColumnPair As ColumnPair
    For Each ColumnPair In This.Item
        If ColumnPair.DestinationListColumn Is Src Then
            IsDestinationMapped = True
            Exit Function
        End If
    Next ColumnPair
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsColumnPairs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("HelperFunctions")
Option Explicit

Private Type TClsColumnPairs
    Item As Collection
    ConstrainRHS As Boolean
End Type

Private this As TClsColumnPairs

'@DefaultMember
'@Enumerator
Public Property Get Item() As Collection
Attribute Item.VB_UserMemId = -4
    Set Item = this.Item
End Property

Public Property Get Count() As Long
    Count = this.Item.Count
End Property

Public Property Get ConstrainRHS() As Boolean
    ConstrainRHS = this.ConstrainRHS
End Property

Public Property Let ConstrainRHS(ByVal RHS As Boolean)
    this.ConstrainRHS = RHS
End Property

Private Sub Class_Initialize()
    Set this.Item = New Collection
    this.ConstrainRHS = True
End Sub

Public Sub AddOrReplace(ByRef colPair As clsColumnPair)
    Dim foundPair As clsColumnPair
    If Exists(colPair) Then
        If ConstrainRHS Then
            Set foundPair = GetPair(RHS:=colPair.RHS)
        Else
            Set foundPair = GetPair(colPair.lhs, colPair.RHS)
        End If
        Set foundPair.lhs = colPair.lhs
        Set foundPair.RHS = colPair.RHS
    Else
        this.Item.Add colPair
    End If
End Sub

Public Sub Add(ByRef colPair As clsColumnPair)
    If Exists(colPair) Then Exit Sub
    this.Item.Add colPair
End Sub

Public Sub Remove(ByRef colPair As clsColumnPair)
    Debug.Assert Not colPair Is Nothing
    Dim idx As Long
    Dim match As clsColumnPair
    Set match = GetPair(colPair.lhs, colPair.RHS)
    idx = GetIndexOfColumnPair(match)
    If idx > -1 Then
        this.Item.Remove idx
    End If
End Sub

Private Function GetIndexOfColumnPair(ByRef colPair) As Long
    Debug.Assert Not colPair Is Nothing
    Dim i As Long
    For i = 1 To this.Item.Count
        If this.Item(i).Equals(colPair) Or (this.ConstrainRHS And (this.Item(i).RHS Is colPair.RHS)) Then
            GetIndexOfColumnPair = i
            Exit Function
        End If
    Next i
    GetIndexOfColumnPair = -1
End Function

Public Sub Clear()
    Dim i As Long
    For i = this.Item.Count To 1 Step -1
        this.Item.Remove i
    Next i
End Sub

Public Function Exists(ByRef colPair As clsColumnPair) As Boolean
    Dim RHS As clsColumnPair
    For Each RHS In this.Item
        If colPair.Equals(RHS) Then
            Exists = True
            Exit Function
        ElseIf ConstrainRHS And colPair.RHS Is RHS.RHS Then
            Exists = True
            Exit Function
        End If
    Next RHS
End Function

Public Function GetPair(Optional ByRef lhs As ListColumn, Optional ByRef RHS As ListColumn) As clsColumnPair
    Dim colPair As clsColumnPair
    
    If IsMissing(lhs) And IsMissing(RHS) Then Exit Function

    For Each colPair In this.Item
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

Public Function TryGetPair(ByRef colPair As clsColumnPair, Optional ByRef lhs As ListColumn, Optional ByRef RHS As ListColumn) As Boolean
    Dim result As clsColumnPair
    Set result = GetPair(lhs, RHS)
    Set colPair = result
    If Not result Is Nothing Then TryGetPair = True
End Function

Public Function IsMappedRHS(ByRef lc As ListColumn) As Boolean
    Dim colPair As clsColumnPair
    For Each colPair In this.Item
        If colPair.RHS Is lc Then
            IsMappedRHS = True
            Exit Function
        End If
    Next colPair
End Function

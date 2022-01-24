VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCompareKeyColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("KeyColumn")
Option Explicit

Private Type TClsCompareKeyColumns
    LHS As clsKeyColumn
    RHS As clsKeyColumn
    LeftOnly As Collection
    Intersection As Collection
    RightOnly As Collection
    Dict As Object
End Type

Private this As TClsCompareKeyColumns

Public Property Get LHS() As clsKeyColumn
    Set LHS = this.LHS
End Property

Public Property Set LHS(ByVal RHS As clsKeyColumn)
    Set this.LHS = RHS
    TryDoWork
End Property

Public Property Get RHS() As clsKeyColumn
    Set RHS = this.RHS
End Property

Public Property Set RHS(ByVal RHS As clsKeyColumn)
    Set this.RHS = RHS
    TryDoWork
End Property

Public Property Get LeftOnly() As Collection
    Set LeftOnly = this.LeftOnly
End Property

Public Property Get Intersection() As Collection
    Set Intersection = this.Intersection
End Property

Public Property Get RightOnly() As Collection
    Set RightOnly = this.RightOnly
End Property

Public Property Get IsSubsetLHS() As Boolean
    IsSubsetLHS = this.LeftOnly.Count = 0
End Property

Public Property Get IsSubsetRHS() As Boolean
    IsSubsetRHS = this.RightOnly.Count = 0
End Property


Public Property Get IsMatch() As Boolean
    IsMatch = IsSubsetLHS And IsSubsetRHS
End Property

' Methods

Private Sub TryDoWork()
    If this.LHS Is Nothing Then Exit Sub
    If this.RHS Is Nothing Then Exit Sub
    DoWork
End Sub

Private Sub DoWork()
    Set this.Dict = CreateObject("Scripting.Dictionary")
    Dim key As Variant
    For Each key In LHS.Keys
        Upsert key, 1
    Next
    For Each key In RHS.Keys
        Upsert key, 2
    Next
    
    Set this.Intersection = New Collection
    Set this.LeftOnly = New Collection
    Set this.RightOnly = New Collection
    
    Dim val As Integer
    
    For Each key In this.Dict.Keys
        Select Case this.Dict(key)
            Case 1
                this.LeftOnly.Add key
            Case 2
                this.RightOnly.Add key
            Case 3
                this.Intersection.Add key
        End Select
    Next key
End Sub

Private Sub Upsert(ByVal key As String, ByVal value As Integer)
    If this.Dict.Exists(key) Then
        this.Dict(key) = this.Dict(key) + value
    Else
        this.Dict(key) = value
    End If
End Sub

Public Function Create(ByRef LHS As clsKeyColumn, ByRef RHS As clsKeyColumn)
    Set Create = New clsCompareKeyColumns
    Set Create.LHS = LHS
    Set Create.RHS = RHS
End Function

Public Function Map() As Variant
    Dim result As Variant
    Dim arr As Variant
    Dim i As Long
    
    arr = this.RHS.GetArray
    ReDim result(LBound(arr) To UBound(arr))
    
    For i = LBound(arr) To UBound(arr)
        result(i) = this.LHS.Find(arr(i))
    Next i
    
    Map = result
End Function

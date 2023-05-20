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

Private Const MINIMUM_TOTAL_SET_SIZE As Long = 5

Private Type TClsCompareKeyColumns
    SrcKeyColumn As KeyColumn
    DstKeyColumn As KeyColumn
    
    SrcSetOnly As Collection
    IntersectionSet As Collection
    DstSetOnly As Collection
    
    Dictionary As Scripting.Dictionary
End Type
Private This As TClsCompareKeyColumns

'@Description "Percentage of the Intersection set versus the total (Union set)."
Public Property Get Score() As Double
Attribute Score.VB_Description = "Percentage of the Intersection set versus the total (Union set)."
    Dim Total As Long
    Total = This.SrcSetOnly.Count + This.IntersectionSet.Count + This.DstSetOnly.Count
    
    Score = 0
    
    If Total < MINIMUM_TOTAL_SET_SIZE Then Exit Property
    
    Score = This.IntersectionSet.Count / Total
End Property

'@Description "ListColumn containing the Keys for the Source column."
Public Property Get SrcKeyColumn() As KeyColumn
Attribute SrcKeyColumn.VB_Description = "ListColumn containing the Keys for the Source column."
    Set SrcKeyColumn = This.SrcKeyColumn
End Property

Public Property Set SrcKeyColumn(ByVal RHS As KeyColumn)
    Debug.Assert Not Me Is KeyColumnComparer
    If Not This.SrcKeyColumn Is Nothing Then
        If This.SrcKeyColumn = RHS Then
            Debug.Assert False
        End If
    End If
    Set This.SrcKeyColumn = RHS
    TryDoWork
End Property

'@Description "ListColumn containing the Keys for the Destination column."
Public Property Get DstKeyColumn() As KeyColumn
Attribute DstKeyColumn.VB_Description = "ListColumn containing the Keys for the Destination column."
    Set DstKeyColumn = This.DstKeyColumn
End Property

Public Property Set DstKeyColumn(ByVal RHS As KeyColumn)
    Debug.Assert Not Me Is KeyColumnComparer
    If Not This.DstKeyColumn Is Nothing Then
        If This.DstKeyColumn = RHS Then
            Debug.Assert False
        End If
    End If
    Set This.DstKeyColumn = RHS
    TryDoWork
End Property

'@Description "The set of Keys that are only present in the Source column."
Public Property Get SrcSetOnly() As Collection
Attribute SrcSetOnly.VB_Description = "The set of Keys that are only present in the Source column."
    Set SrcSetOnly = This.SrcSetOnly
End Property

'@Description "The set of Keys that are present in both columns."
Public Property Get IntersectionSet() As Collection
Attribute IntersectionSet.VB_Description = "The set of Keys that are present in both columns."
    Set IntersectionSet = This.IntersectionSet
End Property

'@Description "The set of Keys that are only present in the Destination column."
Public Property Get DstSetOnly() As Collection
Attribute DstSetOnly.VB_Description = "The set of Keys that are only present in the Destination column."
    Set DstSetOnly = This.DstSetOnly
End Property

'@Description "Returns True if every Key in the Source column is present in the Destination column."
Public Property Get IsSubsetLHS() As Boolean
Attribute IsSubsetLHS.VB_Description = "Returns True if every Key in the Source column is present in the Destination column."
    IsSubsetLHS = This.SrcSetOnly.Count = 0
End Property

'@Description "Returns True if every Key in the Destination column is present in the Source column."
Public Property Get IsSubsetRHS() As Boolean
Attribute IsSubsetRHS.VB_Description = "Returns True if every Key in the Destination column is present in the Source column."
    IsSubsetRHS = This.DstSetOnly.Count = 0
End Property

'@Description "Returns True if each Key is present in both columns."
Public Property Get IsMatch() As Boolean
Attribute IsMatch.VB_Description = "Returns True if each Key is present in both columns."
    IsMatch = IsSubsetLHS And IsSubsetRHS
End Property

' Ctor
Private Sub Class_Initialize()
    Set This.Dictionary = CreateObject("Scripting.Dictionary")
    ClearSets
End Sub

' Methods
Private Sub TryDoWork()
    If This.SrcKeyColumn Is Nothing Then Exit Sub
    If This.DstKeyColumn Is Nothing Then Exit Sub
    CalculateSets
End Sub

'@Description "Calculates which sets each Key is present in."
Private Sub CalculateSets()
Attribute CalculateSets.VB_Description = "Calculates which sets each Key is present in."
    Dim Key As Variant
    
    For Each Key In SrcKeyColumn.Keys
        Upsert Key, 1
    Next Key
    
    For Each Key In DstKeyColumn.Keys
        Upsert Key, 2
    Next Key
    
    ClearSets
    
    For Each Key In This.Dictionary.Keys
        UpdateSet
    Next Key
End Sub

Private Sub ClearSets()
    Set This.IntersectionSet = New Collection
    Set This.SrcSetOnly = New Collection
    Set This.DstSetOnly = New Collection
End Sub

'@Description "Adds the Key to appropriate Collection objects representing the sets."
Private Sub UpdateSet(ByVal Key As Variant)
Attribute UpdateSet.VB_Description = "Adds the Key to appropriate Collection objects representing the sets."
    Select Case This.Dictionary.Item(Key)
        Case 1
            This.SrcSetOnly.Add Key
        Case 2
            This.DstSetOnly.Add Key
        Case 3
            This.IntersectionSet.Add Key
    End Select
End Sub

'@Description "Updates the item for the given Key by incrementing it by the given value. If it doesn't exist, set it to the given value."
Private Sub Upsert(ByVal Key As String, ByVal Value As Long)
Attribute Upsert.VB_Description = "Updates the item for the given Key by incrementing it by the given value. If it doesn't exist, set it to the given value."
    If This.Dictionary.Exists(Key) Then
        This.Dictionary.Item(Key) = This.Dictionary.Item(Key) + Value
    Else
        This.Dictionary.Item(Key) = Value
    End If
End Sub

'@Description "Creates a new instance of KeyColumnComparer."
Public Function Create(ByVal SrcKeyColumn As KeyColumn, ByVal DstKeyColumn As KeyColumn) As KeyColumnComparer
Attribute Create.VB_Description = "Creates a new instance of KeyColumnComparer."
    Dim Result As KeyColumnComparer
    Set Result = New KeyColumnComparer
    With Result
        .SrcKeyColumn = SrcKeyColumn
        .DstKeyColumn = DstKeyColumn
    End With
    Set Create = Result
End Function

'@Description "Returns an array of the same size as the Destination column that has the index of each matching Key in the Source column."
Public Function Map() As Variant
Attribute Map.VB_Description = "Returns an array of the same size as the Destination Key that has the index of each matching Key in the Source column."
    Debug.Assert Not Me Is KeyColumnComparer
    If This.SrcKeyColumn Is Nothing Or This.DstKeyColumn Is Nothing Then
        Err.Raise vbObjectError + 4, Err.Source, "Cannot Map KeyColumns with an empty table"
        Exit Function
    End If
    
    Dim DstKeyArray As Variant
    DstKeyArray = This.DstKeyColumn.GetArray
    
    Dim Result As Variant
    ReDim Result(LBound(DstKeyArray) To UBound(DstKeyArray))
    
    Dim i As Long
    For i = LBound(DstKeyArray) To UBound(DstKeyArray)
        Result(i) = This.SrcKeyColumn.Find(DstKeyArray(i))
    Next i
    
    Map = Result
End Function

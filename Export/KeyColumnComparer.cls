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
'@Folder "MVVM.Model.KeyColumn"
Option Explicit
Implements INotifyPropertyChanged

Private Const MIN_TOTAL As Long = 4

Private Type TClsCompareKeyColumns
    Notifier As INotifyPropertyChanged
    
    LHS As KeyColumn
    RHS As KeyColumn
    
    LeftOnly As Collection
    Intersection As Collection
    RightOnly As Collection
    
    Dictionary As Object
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

Public Property Get Score() As Double
    Dim Total As Long
    Total = This.LeftOnly.Count + This.Intersection.Count + This.RightOnly.Count
    
    Score = 0
    
    If Total <= MIN_TOTAL Then Exit Property
    
    If Total > 0 Then
        Score = This.Intersection.Count / Total
    End If
End Property

Private Sub Class_Initialize()
    Set This.Notifier = New PropertyChangeNotifier
End Sub

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    This.Notifier.RegisterHandler Handler
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    This.Notifier.OnPropertyChanged Source, PropertyName
End Sub

Private Sub OnPropertyChanged(ByVal PropertyName As String)
    INotifyPropertyChanged_OnPropertyChanged Me, PropertyName
End Sub

Public Function Create(ByVal LHS As KeyColumn, ByVal RHS As KeyColumn) As KeyColumnComparer
    Dim Result As KeyColumnComparer
    Set Result = New KeyColumnComparer
    
    With Result
        Set .LHS = LHS
        Set .RHS = RHS
    End With
    
    Set Create = Result
End Function

Public Function Map(Optional ByVal FilterSource As Boolean = False, Optional ByVal FilterDestination As Boolean = False) As Variant
    If This.LHS Is Nothing Or This.RHS Is Nothing Then
        Err.Raise vbObjectError + 4, Err.Source, StringConstants.ERR_MSG_KEYCOLCOMP_EMPTY_TABLE
        Exit Function
    End If
    
    Dim FilterLHS As KeyColumn
    Set FilterLHS = This.LHS
    If FilterSource Then
        Set FilterLHS = KeyColumn.FromRange(This.LHS.Range, This.LHS.IncludeNumbersAsText, True)
    End If
    
    Dim FilterRHS As KeyColumn
    Set FilterRHS = This.RHS
    If FilterDestination Then
        Set FilterRHS = KeyColumn.FromRange(This.RHS.Range, This.RHS.IncludeNumbersAsText, True)
    End If
    
    Dim ArrayRHS As Variant
    ArrayRHS = FilterRHS.GetArray
    
    Dim Result As Variant
    ReDim Result(LBound(ArrayRHS) To UBound(ArrayRHS))
    
    Dim i As Long
    For i = LBound(ArrayRHS) To UBound(ArrayRHS)
        Result(i) = FilterLHS.Find(ArrayRHS(i))
    Next i
    
    Map = Result
End Function

Private Sub TryDoWork()
    If This.LHS Is Nothing Then Exit Sub
    If This.RHS Is Nothing Then Exit Sub
    
    DoWork
End Sub

Private Sub DoWork()
    Set This.Dictionary = CreateObject("Scripting.Dictionary")
    
    Dim Key As Variant
    For Each Key In LHS.Keys
        Upsert Key, 1
    Next
    
    For Each Key In RHS.Keys
        Upsert Key, 2
    Next
    
    Set This.Intersection = New Collection
    Set This.LeftOnly = New Collection
    Set This.RightOnly = New Collection
    
    For Each Key In This.Dictionary.Keys
        Select Case This.Dictionary.Item(Key)
            Case 1
                This.LeftOnly.Add Key
            Case 2
                This.RightOnly.Add Key
            Case 3
                This.Intersection.Add Key
        End Select
    Next Key
End Sub

Private Sub Upsert(ByVal Key As String, ByVal Value As Long)
    If This.Dictionary.Exists(Key) Then
        This.Dictionary.Item(Key) = This.Dictionary.Item(Key) + Value
    Else
        This.Dictionary.Item(Key) = Value
    End If
End Sub

Public Property Get ToTextTable() As String
    Dim MaxLines As Long
    MaxLines = This.LeftOnly.Count
    If This.Intersection.Count > MaxLines Then MaxLines = This.Intersection.Count
    If This.RightOnly.Count > MaxLines Then MaxLines = This.RightOnly.Count
    
    Dim Table() As Variant
    ReDim Table(1 To MaxLines - 1)
    
    Dim i As Long
    For i = 1 To (MaxLines - 1)
        Dim Row() As Variant
        Row = Array(vbNullString, vbNullString, vbNullString)
        With This
            If (i < .LeftOnly.Count) Then Row(0) = .LeftOnly.Item(i)
            If (i < .Intersection.Count) Then Row(1) = .Intersection.Item(i)
            If (i < .RightOnly.Count) Then Row(2) = .RightOnly.Item(i)
        End With
        Table(i) = Join(Row, vbTab)
    Next i
    
    ToTextTable = Join(Table, vbCrLf)
End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferInstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "TableTransfer"
Option Explicit

Private Type TState
    Source As ListObject
    Destination As ListObject
    SourceKey As ListColumn
    DestinationKey As ListColumn
    ValuePairs As Collection
    MapResult As Variant
    Flags As Integer
End Type

Private This As TState

' Properties
Public Property Get Flags() As Integer
    Set Flags = This.Source
End Property

Public Property Let Flags(ByVal Flags As Integer)
    This.Flags = Flags
End Property

Public Property Get Source() As ListObject
    Set Source = This.Source
End Property

Public Property Set Source(ByVal RHS As ListObject)
    Set This.Source = RHS
End Property

Public Property Get Destination() As ListObject
    Set Destination = This.Destination
End Property

Public Property Set Destination(ByVal RHS As ListObject)
    Set This.Destination = RHS
End Property

Public Property Get SourceKey() As ListColumn
    Set SourceKey = This.SourceKey
End Property

Public Property Set SourceKey(ByVal RHS As ListColumn)
    Set This.SourceKey = RHS
End Property

Public Property Get DestinationKey() As ListColumn
    Set DestinationKey = This.DestinationKey
End Property

Public Property Set DestinationKey(ByVal RHS As ListColumn)
    Set This.DestinationKey = RHS
End Property

Public Property Get ValuePairs() As Collection
    Set ValuePairs = This.ValuePairs
End Property

Public Property Set ValuePairs(ByVal RHS As Collection)
    Set This.ValuePairs = RHS
End Property

Public Property Get MapResult() As Variant
    If IsObject(This.MapResult) Then
        Set MapResult = This.MapResult
    Else
        MapResult = This.MapResult
    End If
End Property

Public Property Let MapResult(ByVal RHS As Variant)
    This.MapResult = RHS
End Property

Public Property Set MapResult(ByVal RHS As Variant)
    Set This.MapResult = RHS
End Property

' Ctor
Private Sub Class_Initialize()
    Set ValuePairs = New Collection
End Sub

' Methods
Public Sub Add(ByVal LHS As ListColumn, ByVal RHS As ListColumn)
    ValuePairs.Add ColumnPair.Create(LHS, RHS)
End Sub

Public Function IsValid() As Boolean
    If SourceKey Is Nothing Then Exit Function
    If DestinationKey Is Nothing Then Exit Function
    If ValuePairs Is Nothing Then Exit Function
    If ValuePairs.Count = 0 Then Exit Function
    IsValid = True
End Function

Public Sub Transfer()
    If Not IsValid Then Exit Sub

    If HasFlag(This.Flags, ClearDestinationFirst) Then
        Debug.Print "Clear first"
    Else
        Debug.Print "Don't clear first"
    End If

    SetMapResult
    
    Dim colPair As ColumnPair
    For Each colPair In ValuePairs
        TransferOne colPair
    Next colPair
End Sub

Private Sub SetMapResult()
    Dim compare As KeyColumnComparer
    Set compare = KeyColumnComparer.Create(KeyColumn.FromColumn(SourceKey), KeyColumn.FromColumn(DestinationKey))
    MapResult = compare.Map
End Sub

Private Sub TransferOne(ByVal colPair As ColumnPair)
   DoTransfer MapResult, colPair.LHS, colPair.RHS
End Sub

Private Sub DoTransfer(ByVal Map As Variant, ByVal Source As ListColumn, ByVal Destination As ListColumn)
    Dim i As Integer
    Dim arrLHS As Variant
    Dim arrRHS As Variant
    Dim arrLHSOffset As Long
    Dim arrRHSOffset As Long
    Dim newValue As Variant
    
    arrLHS = Source.DataBodyRange.Value2
    arrRHS = Destination.DataBodyRange.Value2
    arrLHSOffset = Source.DataBodyRange.Row
    arrRHSOffset = Destination.DataBodyRange.Row
    
    For i = LBound(Map) To UBound(Map)
        If Map(i) > -1 Then
            newValue = arrLHS(Map(i) - arrLHSOffset + 1, 1)
            Select Case VarType(newValue)
                Case vbString
                    If newValue <> vbNullString Then
                        arrRHS(i - arrRHSOffset + 1, 1) = newValue
                    End If
                Case Else
                    arrRHS(i - arrRHSOffset + 1, 1) = newValue
            End Select
        Else
            arrRHS(i - arrRHSOffset + 1, 1) = "Unmapped"
        End If
    Next i
    
    Destination.DataBodyRange.Value2 = arrRHS
End Sub

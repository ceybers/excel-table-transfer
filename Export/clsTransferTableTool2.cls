VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTransferTableTool2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "TransferTableTool2"
Option Explicit

Public Source As ListObject
Public Destination As ListObject
Public SourceKey As ListColumn
Public DestinationKey As ListColumn
Public ValuePairs As Collection
Public MapResult As Variant

Private Sub Class_Initialize()
    Set ValuePairs = New Collection
End Sub

Public Sub Add(ByVal LHS As ListColumn, ByVal RHS As ListColumn)
    ValuePairs.Add ColumnPair.Create(LHS, RHS)
End Sub

Public Sub Transfer()
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
   modTestTransferInstruction2.DoTransfer MapResult, colPair.LHS, colPair.RHS
End Sub

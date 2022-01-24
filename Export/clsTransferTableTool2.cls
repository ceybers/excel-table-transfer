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

Public Sub Add(ByRef lhs As ListColumn, ByRef RHS As ListColumn)
    ValuePairs.Add clsColumnPair.Create(lhs, RHS)
End Sub

Public Sub Transfer()
    SetMapResult
    
    Dim colPair As clsColumnPair
    For Each colPair In ValuePairs
        TransferOne colPair
    Next colPair
End Sub

Private Sub SetMapResult()
    Dim compare As clsCompareKeyColumns
    Set compare = clsCompareKeyColumns.Create(clsKeyColumn.FromColumn(SourceKey), clsKeyColumn.FromColumn(DestinationKey))
    MapResult = compare.Map
End Sub

Private Sub TransferOne(ByRef colPair As clsColumnPair)
   modTestTransferInstruction2.DoTransfer MapResult, colPair.lhs, colPair.RHS
End Sub

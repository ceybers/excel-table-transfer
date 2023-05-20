VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyColumnMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("TransferEngine")
Option Explicit

Private Type TKeyColumnMapper
    SrcKeyColumn As ListColumn
    DstKeyColumn As ListColumn
End Type
Private This As TKeyColumnMapper

Public Property Get SrcKeyColumn() As ListColumn
    Set SrcKeyColumn = This.SrcKeyColumn
End Property

Public Property Set SrcKeyColumn(ByVal RHS As ListColumn)
    Set This.SrcKeyColumn = RHS
End Property

Public Property Get DstKeyColumn() As ListColumn
    Set DstKeyColumn = This.DstKeyColumn
End Property

Public Property Set DstKeyColumn(ByVal RHS As ListColumn)
    Set This.DstKeyColumn = RHS
End Property

Public Function Create(ByVal SrcKeyColumn As ListColumn, ByVal DstKeyColumn As ListColumn) As KeyColumnMapper
    Dim Result As KeyColumnMapper
    Set Result = New KeyColumnMapper
    With Result
        Set .SrcKeyColumn = SrcKeyColumn
        Set .DstKeyColumn = DstKeyColumn
    End With
    Set Create = Result
End Function

Public Function GenerateMap() As Variant
    Dim Result As Variant
    'ReDim Result(1 To This.DstKeyColumn.DataBodyRange.Rows.Count, 1 to 1)
    ReDim Result(1 To This.DstKeyColumn.DataBodyRange.Rows.Count)
    
    Dim i As Long
    For i = 1 To UBound(Result, 1)
        Result(i) = This.SrcKeyColumn.DataBodyRange.Find(This.DstKeyColumn.DataBodyRange.Cells(i, 1)).Row - This.SrcKeyColumn.Range.Cells(1, 1).Row
        'Result(i) = This.SrcKeyColumn.DataBodyRange.Find(This.DstKeyColumn.DataBodyRange.Cells(i, 1)).Row
    Next i
    
    GenerateMap = Result
End Function

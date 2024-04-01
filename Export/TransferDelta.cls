VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferDelta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.Model.TransferDeltas"
Option Explicit

Private Type TState
    RowIndex As Long
    ColumnIndex As Long
    Key As String
    FieldSource As String
    FieldDestination As String
    ValueBefore As String
    ValueAfter As String
    DataType As VbVarType
    ChangeType As ChangeType2
End Type
Private This As TState

Public Property Get RowIndex() As Long
    RowIndex = This.RowIndex
End Property

Public Property Get ColumnIndex() As Long
    ColumnIndex = This.ColumnIndex
End Property

Public Property Get Key() As String
    Key = This.Key
End Property

Public Property Get FieldSource() As String
    FieldSource = This.FieldSource
End Property

Public Property Get FieldDestination() As String
    FieldDestination = This.FieldDestination
End Property

Public Property Get ValueBefore() As String
    ValueBefore = This.ValueBefore
End Property

Public Property Get ValueAfter() As String
    ValueAfter = This.ValueAfter
End Property

Public Property Get DataType() As VbVarType
    DataType = This.DataType
End Property

Public Property Get ChangeType() As ChangeType2
    ChangeType = This.ChangeType
End Property

Public Sub Load(ByVal DeltaVariant As Variant)
    If IsEmpty(DeltaVariant) Then Exit Sub
    
    With This
        .Key = DeltaVariant(DeltaIndex.tdKey)
        .RowIndex = DeltaVariant(DeltaIndex.tdRow)
        .ColumnIndex = DeltaVariant(DeltaIndex.tdCol)
        .FieldSource = DeltaVariant(DeltaIndex.tdFieldSrc)
        .FieldDestination = DeltaVariant(DeltaIndex.tdFieldDst)
        .ValueBefore = DeltaVariant(DeltaIndex.tdValueBefore)
        .ValueAfter = DeltaVariant(DeltaIndex.tdValueAfter)
        .DataType = DeltaVariant(DeltaIndex.tdVarType)
        .ChangeType = DeltaVariant(DeltaIndex.tdChangeType)
    End With
End Sub

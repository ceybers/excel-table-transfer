VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferDeltasViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM2.ViewModels"
Option Explicit

Private Const MAX_LIMIT_DELTAS As Long = 1000000

Private Type TState
    TransferDeltas As TransferDeltas
End Type
Private This As TState

Public Property Get Keys() As Collection
    Set Keys = This.TransferDeltas.Keys
End Property

Public Property Get Fields() As Collection
    Set Fields = This.TransferDeltas.Fields
End Property

Public Property Get Deltas() As Collection
    Set Deltas = This.TransferDeltas.Items
End Property

Public Sub Load(ByVal TransferDeltas As TransferDeltas)
    Set This.TransferDeltas = TransferDeltas
End Sub

Public Sub TrySelectKey(ByVal Key As String)
    This.TransferDeltas.KeyFilter = Key
End Sub

Public Sub TrySelectField(ByVal Field As String)
    This.TransferDeltas.FieldFilter = Field
End Sub

Public Function CanShowAll() As Boolean
    CanShowAll = This.TransferDeltas.Items.Count > This.TransferDeltas.Limit
End Function

Public Sub DoShowAll()
    This.TransferDeltas.Limit = MAX_LIMIT_DELTAS
End Sub

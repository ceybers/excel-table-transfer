VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TransferDeltas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.Model.TransferDeltas"
Option Explicit

Private Const DEFAULT_LIMIT As Long = 10

Private Type TState
    BaseItems As Collection
    Items As Collection
    Keys As Collection
    Fields As Collection
    
    KeyFilter As String
    FieldFilter As String
    Limit As Long
    
    DeltaVariants As Variant
End Type
Private This As TState

Public Property Get Keys() As Collection
    Set Keys = This.Keys
End Property

Public Property Get Fields() As Collection
    Set Fields = This.Fields
End Property

Public Property Get KeyFilter() As String
    KeyFilter = This.KeyFilter
End Property

Public Property Let KeyFilter(ByVal vNewValue As String)
    If This.KeyFilter = vNewValue Then Exit Property
    This.KeyFilter = vNewValue
    UpdateItems
End Property

Public Property Get FieldFilter() As String
    FieldFilter = This.FieldFilter
End Property

Public Property Let FieldFilter(ByVal vNewValue As String)
    If This.FieldFilter = vNewValue Then Exit Property
    This.FieldFilter = vNewValue
    UpdateItems
End Property

Public Property Get Items() As Collection
    Set Items = This.Items
End Property

Public Property Get Limit() As Long
    Limit = This.Limit
End Property

Public Property Let Limit(ByVal vNewValue As Long)
    If This.Limit = vNewValue Then Exit Property
    This.Limit = vNewValue
    UpdateBaseItems
    UpdateItems
End Property

Private Sub Class_Initialize()
    Set This.BaseItems = New Collection
    Set This.Items = New Collection
    Set This.Keys = New Collection
    Set This.Fields = New Collection
    This.Limit = DEFAULT_LIMIT
End Sub

Public Sub Load(ByVal DeltaVariants As Variant)
    This.DeltaVariants = DeltaVariants
    
    UpdateBaseItems
    UpdateItems
End Sub

Private Sub UpdateBaseItems()
    CollectionClear This.BaseItems
    
    Dim i As Long
    For i = LBound(This.DeltaVariants) To UBound(This.DeltaVariants)
        Dim TransferDelta As TransferDelta
        Set TransferDelta = New TransferDelta
        TransferDelta.Load This.DeltaVariants(i)
        This.BaseItems.Add Key:=CStr(i), Item:=TransferDelta
        If i >= This.Limit Then Exit For
    Next i
End Sub

Private Sub UpdateItems()
    CollectionClear This.Items
    
    Dim TransferDelta As TransferDelta
    For Each TransferDelta In This.BaseItems
        If TestFilter(TransferDelta) Then
            This.Items.Add Item:=TransferDelta
        End If
    Next TransferDelta
End Sub

Private Function TestFilter(ByVal TransferDelta As TransferDelta) As Boolean
    If This.KeyFilter <> vbNullString Then
        If TransferDelta.Key <> This.KeyFilter Then Exit Function
    End If
    
    If This.FieldFilter <> vbNullString Then
        If TransferDelta.FieldDestination <> This.FieldFilter Then Exit Function
    End If
    
    TestFilter = True
End Function

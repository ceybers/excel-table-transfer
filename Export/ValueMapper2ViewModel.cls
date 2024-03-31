VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueMapper2ViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVVM2.ViewModels")
Option Explicit
Implements IViewModel

Private Type TState
    Source As ValueColumns
    Destination As ValueColumns
    Mapped As ColumnPairs2
End Type

Private This As TState

Public Property Get Source() As ValueColumns
    Set Source = This.Source
End Property

Public Property Get Destination() As ValueColumns
    Set Destination = This.Destination
End Property

Public Property Get Mapped() As ColumnPairs2
    Set Mapped = This.Mapped
End Property

Private Sub Class_Initialize()
    Set This.Source = New ValueColumns
    Set This.Destination = New ValueColumns
    Set This.Mapped = New ColumnPairs2
End Sub

Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = True
End Function

Public Sub Load(ByVal Source As ListObject, ByVal Destination As ListObject)
    This.Source.Load Source
    This.Destination.Load Destination
End Sub

Public Function CanMapAdd() As Boolean
    If This.Source.Selected Is Nothing Then Exit Function
    If This.Destination.Selected Is Nothing Then Exit Function
    If Not This.Source.Selected.IsSelectable Then Exit Function
    If Not This.Destination.Selected.IsSelectable Then Exit Function
    If This.Mapped.Exists(This.Source.Selected.Name, This.Destination.Selected.Name) Then Exit Function
    
    CanMapAdd = True
End Function

Public Function CanMapRemove() As Boolean
    If This.Source.Selected Is Nothing Then Exit Function
    If This.Destination.Selected Is Nothing Then Exit Function
    If Not This.Source.Selected.IsSelectable Then Exit Function
    If Not This.Destination.Selected.IsSelectable Then Exit Function
    
    CanMapRemove = This.Mapped.Exists(This.Source.Selected.Name, This.Destination.Selected.Name)
End Function

Public Sub DoMapAdd()
    This.Mapped.Add This.Source.Selected.Name, This.Destination.Selected.Name
    This.Source.Selected.IsMapped = True
    This.Destination.Selected.IsMapped = True
End Sub

Public Sub DoMapRemove()
    This.Mapped.Remove This.Source.Selected.Name, This.Destination.Selected.Name
    This.Source.Selected.IsMapped = False
    This.Destination.Selected.IsMapped = False
End Sub

Public Function CanNext() As Boolean
    CanNext = This.Mapped.Items.Count > 0
End Function

Public Sub DoAutoMap()
    CollectionClear This.Mapped.Items
    
    Dim SourceColumn As ValueColumn
    For Each SourceColumn In This.Source.ValueColumns
        Dim DestinationColumn As ValueColumn
        For Each DestinationColumn In This.Destination.ValueColumns
            If SourceColumn.Name = DestinationColumn.Name And SourceColumn.IsSelectable And DestinationColumn.IsSelectable Then
                This.Mapped.Add SourceColumn.Name, DestinationColumn.Name
                SourceColumn.IsMapped = True
                DestinationColumn.IsMapped = True
            End If
        Next DestinationColumn
    Next SourceColumn
End Sub

Public Function CanRemoveAll() As Boolean
    CanRemoveAll = This.Mapped.Items.Count > 0
End Function

Public Sub DoRemoveAll()
    CollectionClear This.Mapped.Items
    This.Source.Reset
    This.Destination.Reset
End Sub


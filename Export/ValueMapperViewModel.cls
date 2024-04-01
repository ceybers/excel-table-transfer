VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueMapperViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.ViewModels"
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged
Implements IHandlePropertyChanged

Private Type TState
    Notifier As INotifyPropertyChanged
    
    Source As ValueColumns
    Destination As ValueColumns
    Mapped As ColumnPairs
End Type

Private This As TState

Public Property Get Source() As ValueColumns
    Set Source = This.Source
End Property

Public Property Get Destination() As ValueColumns
    Set Destination = This.Destination
End Property

Public Property Get Mapped() As ColumnPairs
    Set Mapped = This.Mapped
End Property

Private Sub Class_Initialize()
    Set This.Notifier = New PropertyChangeNotifier
    Set This.Source = New ValueColumns
    Set This.Destination = New ValueColumns
    
    Dim TestNotifyPropertyChanged As INotifyPropertyChanged
    Set TestNotifyPropertyChanged = This.Source
    TestNotifyPropertyChanged.RegisterHandler Me
    Set TestNotifyPropertyChanged = This.Destination
    TestNotifyPropertyChanged.RegisterHandler Me
    
'    This.Notifier.RegisterHandler This.Source
    
    Set This.Mapped = New ColumnPairs
End Sub

Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = True
End Function

Private Sub INotifyPropertyChanged_RegisterHandler(ByVal Handler As IHandlePropertyChanged)
    This.Notifier.RegisterHandler Handler
End Sub

Private Sub INotifyPropertyChanged_OnPropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    This.Notifier.OnPropertyChanged Source, PropertyName
End Sub

Private Sub OnPropertyChanged(ByVal PropertyName As String)
    INotifyPropertyChanged_OnPropertyChanged Me, PropertyName
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    If PropertyName = "Selected" Then
        OnPropertyChanged "CanMapAdd"
        OnPropertyChanged "CanMapRemove"
        OnPropertyChanged "CanRemoveAll"
        OnPropertyChanged "CanNext"
    End If
End Sub

Public Sub Load(ByVal Source As ListObject, ByVal Destination As ListObject)
    This.Source.Load Source
    This.Destination.Load Destination
End Sub

Public Property Get CanMapAdd() As Boolean
    If This.Source.Selected Is Nothing Then Exit Property
    If This.Destination.Selected Is Nothing Then Exit Property
    If Not This.Source.Selected.IsSelectable Then Exit Property
    If Not This.Destination.Selected.IsSelectable Then Exit Property
    If This.Mapped.Exists(This.Source.Selected.Name, This.Destination.Selected.Name) Then Exit Property
    
    CanMapAdd = True
End Property

Public Property Get CanMapRemove() As Boolean
    If This.Source.Selected Is Nothing Then Exit Property
    If This.Destination.Selected Is Nothing Then Exit Property
    If Not This.Source.Selected.IsSelectable Then Exit Property
    If Not This.Destination.Selected.IsSelectable Then Exit Property
    
    CanMapRemove = This.Mapped.Exists(This.Source.Selected.Name, This.Destination.Selected.Name)
End Property

Public Sub DoMapAdd()
    This.Mapped.Add This.Source.Selected.Name, This.Destination.Selected.Name
    This.Source.Selected.IsMapped = True
    This.Destination.Selected.IsMapped = True
    
    CollectionChanged
End Sub

Public Sub DoMapRemove()
    This.Mapped.Remove This.Source.Selected.Name, This.Destination.Selected.Name
    This.Source.Selected.IsMapped = False
    This.Destination.Selected.IsMapped = False
End Sub

Public Property Get CanNext() As Boolean
    CanNext = This.Mapped.Count > 0
End Property

Public Sub DoAutoMap()
    This.Mapped.Clear
    
    Dim SourceColumn As ValueColumn
    For Each SourceColumn In This.Source.ValueColumns
        Dim DestinationColumn As ValueColumn
        For Each DestinationColumn In This.Destination.ValueColumns
            TryAutoMapOne SourceColumn, DestinationColumn
        Next DestinationColumn
    Next SourceColumn
    
    CollectionChanged
End Sub

Private Sub TryAutoMapOne(ByVal Source As ValueColumn, ByVal Destination As ValueColumn)
    If Source.Name <> Destination.Name Then Exit Sub
    If Source.IsSelectable = False Then Exit Sub
    If Destination.IsSelectable = False Then Exit Sub

    This.Mapped.Add Source.Name, Destination.Name
    Source.IsMapped = True
    Destination.IsMapped = True
End Sub

Public Property Get CanRemoveAll() As Boolean
    CanRemoveAll = This.Mapped.Count > 0
End Property

Public Sub DoRemoveAll()
    This.Mapped.Clear
    This.Source.Reset
    This.Destination.Reset
    
    CollectionChanged
End Sub

Private Sub CollectionChanged()
    OnPropertyChanged "Mapped"
    OnPropertyChanged "CanNext"
    OnPropertyChanged "CanMapAdd"
    OnPropertyChanged "CanMapRemove"
    OnPropertyChanged "CanRemoveAll"
End Sub

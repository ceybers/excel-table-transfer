VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TablePickerViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.ViewModels"
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged

Private Type TState
    Notifier As INotifyPropertyChanged
    'Context As AppContext
    
    SelectedTable As ListObject
    Source As TableProps
    Destination As TableProps
    
    BaseItems As AvailableTables
    Items As Collection
End Type

Private This As TState

Public Property Get SelectedTable() As ListObject
    Set SelectedTable = This.SelectedTable
End Property

Public Property Set SelectedTable(ByVal RHS As ListObject)
    Set This.SelectedTable = RHS
    OnPropertyChanged "CanPickSelected"
End Property

Public Property Get SourceTable() As ListObject
    Set SourceTable = This.Source.ListObject
End Property

Public Property Set SourceTable(ByVal RHS As ListObject)
    This.Source.Load RHS
    GetNodeByObject(RHS).IsSelected = True
    OnPropertyChanged "SourceWorkbookName"
    OnPropertyChanged "SourceTableName"
    OnPropertyChanged "CanNext"
End Property

Public Property Get DestinationTable() As ListObject
    Set DestinationTable = This.Destination.ListObject
End Property

Public Property Set DestinationTable(ByVal RHS As ListObject)
     This.Destination.Load RHS
    GetNodeByObject(RHS).IsSelected = True
    OnPropertyChanged "DestinationWorkbookName"
    OnPropertyChanged "DestinationTableName"
    OnPropertyChanged "CanNext"
End Property

Public Property Get Items() As Collection
    Set Items = This.Items
End Property

Private Sub Class_Initialize()
    Set This.Source = New TableProps
    Set This.Destination = New TableProps
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

Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = True
End Function

Public Sub Load()
    Set This.Items = New Collection
    Set This.BaseItems = New AvailableTables
    This.BaseItems.Load
    
    UpdateNodes
End Sub

Private Sub UpdateNodes()
    CollectionClear This.Items
    
    Dim Node As AvailableTableNode
    For Each Node In This.BaseItems.Nodes
        This.Items.Add Node
    Next Node
End Sub

Public Function TrySelect(ByVal Key As String) As Boolean
    Set This.SelectedTable = Nothing
    Dim Node As AvailableTableNode
    For Each Node In This.Items
        If Node.Key = Key Then
            If Node.NodeType = ttListObject Then
                Set SelectedTable = Node.Object
                TrySelect = True
                Exit Function
            End If
        End If
    Next Node
End Function

Public Property Get CanPickSelected() As Boolean
    CanPickSelected = Not This.SelectedTable Is Nothing
End Property

Public Property Get CanNext() As Boolean
    If This.Source.IsValid = False Then Exit Property
    If This.Destination.IsValid = False Then Exit Property
    If This.Source.ListObject Is This.Destination.ListObject Then Exit Property
    
    CanNext = True
End Property

Public Sub PickSelectedTable(ByVal Direction As TtDirection)
    If Direction = ttSource Then
        If Not This.Source.Node Is Nothing Then
            This.Source.Node.IsSelected = False
        End If
        Set SourceTable = This.SelectedTable
    ElseIf Direction = ttDestination Then
        If Not This.Destination.Node Is Nothing Then
            This.Destination.Node.IsSelected = False
        End If
        Set DestinationTable = This.SelectedTable
    End If
End Sub

Private Function GetNodeByObject(ByVal Object As Object) As AvailableTableNode
    Dim Item As AvailableTableNode
    For Each Item In This.Items
        If Item.Object Is Object Then
            Set GetNodeByObject = Item
            Exit Function
        End If
    Next Item
End Function

Public Property Get SourceTableName() As String
    If This.Source.IsValid Then
        SourceTableName = This.Source.TableName
    Else
        SourceTableName = NO_TABLE_SELECTED
    End If
End Property

Public Property Get SourceWorkbookName() As String
    If This.Source.IsValid Then
        SourceWorkbookName = This.Source.WorkbookName
    Else
        SourceWorkbookName = NO_TABLE_SELECTED
    End If
End Property

Public Property Get DestinationTableName() As String
    If This.Destination.IsValid Then
        DestinationTableName = This.Destination.TableName
    Else
        DestinationTableName = NO_TABLE_SELECTED
    End If
End Property

Public Property Get DestinationWorkbookName() As String
    If This.Destination.IsValid Then
        DestinationWorkbookName = This.Destination.WorkbookName
    Else
        DestinationWorkbookName = NO_TABLE_SELECTED
    End If
End Property

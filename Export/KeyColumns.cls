VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.Model.KeyColumn"
Option Explicit
Implements INotifyPropertyChanged

Private Type TState
    Notifier As INotifyPropertyChanged
    
    ListObject As ListObject
    KeyColumns As Collection
    Selected As KeyColumn
End Type
Private This As TState

Public Property Get ListObject() As ListObject
    Set ListObject = This.ListObject
End Property

Public Property Get KeyColumns() As Collection
    Set KeyColumns = This.KeyColumns
End Property

Public Property Get IsSelected() As Boolean
    IsSelected = Not This.Selected Is Nothing
End Property

Public Property Get Selected() As KeyColumn
    Set Selected = This.Selected
End Property

Private Sub Class_Initialize()
    Set This.KeyColumns = New Collection
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

Public Sub Load(ByVal ListObject As ListObject)
    Set This.ListObject = ListObject
    GenerateKeyColumns
End Sub

Private Sub GenerateKeyColumns()
    Dim ListColumn As ListColumn
    For Each ListColumn In This.ListObject.ListColumns
        Dim NewKeyColumn As KeyColumn
        Set NewKeyColumn = KeyColumn.FromColumn(ListColumn)
        This.KeyColumns.Add Key:=ListColumn.Name, Item:=NewKeyColumn
    Next ListColumn
    
    OnPropertyChanged "KeyColumns"
End Sub

Public Sub TrySelect(ByVal ColumnName As String)
    Dim ListColumn As ListColumn
    If ListColumnHelpers.TryGetListColumn(This.ListObject, ColumnName, ListColumn) Then
        Set This.Selected = This.KeyColumns.Item(ColumnName)
        OnPropertyChanged "Caption"
    End If
End Sub

Public Property Get Caption() As String
    If This.Selected Is Nothing Then
        Caption = NO_COLUMN_SELECTED
    Else
        Caption = This.Selected.Count & " of " & CStr(This.Selected.Range.Cells.Count) & " distinct"
    End If
End Property

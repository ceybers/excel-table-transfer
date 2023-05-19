VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TablePickerViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.TablePicker.ViewModel"
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged

Private Type TState
    Notifier As INotifyPropertyChanged

    SourceTables As PickableTables
    DestinationTables As PickableTables
End Type
Private This As TState

Private Function IViewModel_IsValid() As Boolean
    IViewModel_IsValid = True
End Function

Private Function IViewModel_Self() As IViewModel
    Set IViewModel_Self = Me
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

Private Sub Class_Initialize()
    Set This.Notifier = New PropertyChangeNotifier

    Set This.SourceTables = New PickableTables
    Set This.DestinationTables = New PickableTables
End Sub

'---
Public Property Get SourceTables() As PickableTables
    Set SourceTables = This.SourceTables
End Property

Public Property Get DestinationTables() As PickableTables
    Set DestinationTables = This.DestinationTables
End Property

'---
Public Sub Load(ByVal ListObject As ListObject)
    Dim Worksheet As Worksheet
    Dim Workbook As Workbook
    Set Worksheet = ListObject.Parent
    Set Workbook = Worksheet.Parent
    
    This.SourceTables.Load Application
    This.DestinationTables.Load Application
    'Stop
End Sub

Public Sub TrySetColumnChecked(ByVal PropertyPath As String, ByVal ItemKey As String, ByVal vNewValue As Boolean)
    Dim PickableTables As PickableTables
    Set PickableTables = CallByName(Me, PropertyPath, VbGet)
    
    Dim PickableTable As PickableTable
    Set PickableTable = PickableTables.GetByKey(ItemKey)
    
    If PickableTable.Checked = vNewValue Then Exit Sub
    PickableTable.Checked = vNewValue
    'OnPropertyChanged "zzzTablePickerViewModel.TrySetColumnChecked"
End Sub

Public Sub TrySetColumnSelected(ByVal PropertyPath As String, ByVal ItemKey As String, ByVal vNewValue As Boolean)
    Dim PickableTables As PickableTables
    Set PickableTables = CallByName(Me, PropertyPath, VbGet)
    
    Dim PickableTable As PickableTable
    Set PickableTable = PickableTables.GetByKey(ItemKey)
    
    If PickableTable.Selected = vNewValue Then Exit Sub
    PickableTable.Selected = vNewValue
    'OnPropertyChanged "zzzTablePickerViewModel.TrySetColumnSelected"
End Sub

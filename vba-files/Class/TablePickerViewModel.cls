VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TablePickerViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.TableMapper.ViewModel"
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged

Private Type TState
    Notifier As INotifyPropertyChanged

    Caption As String
    Selected As PickableTable
    PickableTables As PickableTables
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
End Sub
'---

Public Property Get Selected() As PickableTable
    Set Selected = This.Selected
End Property

Public Property Set Selected(ByVal vNewValue As PickableTable)
    If Not This.Selected Is Nothing Then
        If This.Selected Is vNewValue Then Exit Property
    End If
    Set This.Selected = vNewValue
    OnPropertyChanged "Selected"
End Property

Public Property Get Item() As Scripting.Dictionary
    Set Item = This.PickableTables.Item
End Property

Public Property Get Caption() As String
    Caption = This.Caption
End Property

Public Property Let Caption(ByVal vNewValue As String)
    This.Caption = vNewValue
End Property

'---
Public Sub Load(ByVal ListObject As ListObject)
    Dim Worksheet As Worksheet
    Dim Workbook As Workbook
    Set Worksheet = ListObject.Parent
    Set Workbook = Worksheet.Parent
    
    Set This.PickableTables = New PickableTables
    This.PickableTables.Load Application
End Sub

'---
Public Sub TrySetColumnChecked(ByVal PropertyPath As String, ByVal ItemKey As String, ByVal vNewValue As Boolean)
    Dim PickableTable As PickableTable
    Set PickableTable = This.PickableTables.Item(ItemKey)
    
    If PickableTable.Checked = vNewValue Then Exit Sub
    PickableTable.Checked = vNewValue
    
    If vNewValue Then
        Set Selected = PickableTable
    End If
End Sub

Public Sub TrySetColumnSelected(ByVal PropertyPath As String, ByVal ItemKey As String, ByVal vNewValue As Boolean)
    Dim PickableTable As PickableTable
    Set PickableTable = This.PickableTables.Item(ItemKey)
    
    If PickableTable.Selected = vNewValue Then Exit Sub
    PickableTable.Selected = vNewValue
End Sub

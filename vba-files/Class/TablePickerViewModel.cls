VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TablePickerViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "MVVM.TableMapper.ViewModel"
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged
Implements IHandlePropertyChanged

Private Type TState
    Notifier As INotifyPropertyChanged

    Caption As String
    Direction As TransferDirection
    
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
Public Property Get SelectedAsText() As String
    If This.Selected Is Nothing Then
        SelectedAsText = "nothing"
        Exit Property
    End If
    SelectedAsText = This.Selected.Name
End Property

Public Property Get Selected() As PickableTable
    Set Selected = This.Selected
End Property

Public Property Set Selected(ByVal vNewValue As PickableTable)
    If Not This.Selected Is Nothing Then
        If This.Selected Is vNewValue Then Exit Property
    End If
    Set This.Selected = vNewValue
    OnPropertyChanged "Selected"
    OnPropertyChanged "SelectedAsText"
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

Public Property Get Direction() As TransferDirection
    Direction = This.Direction
End Property

Public Property Let Direction(ByVal vNewValue As TransferDirection)
    This.Direction = vNewValue
End Property

'---
Public Function Create(ByVal ListObject As Object, _
    ByVal Caption As String, ByVal Direction As TransferDirection, _
    ByVal Notifier As INotifyPropertyChanged) As TablePickerViewModel
    
    Dim Result As TablePickerViewModel
    Set Result = New TablePickerViewModel
    Result.Load ListObject
    Result.Caption = Caption
    Result.Direction = Direction
    Notifier.RegisterHandler Result
    
    Set Create = Result
End Function

Public Sub Load(ByVal ListObject As ListObject)
    Debug.Assert Not Me Is TablePickerViewModel
    
    Set This.PickableTables = New PickableTables
    This.PickableTables.Load Application
    SelectByListObject ListObject
End Sub

'---
Public Sub TrySetColumnChecked(ByVal PropertyPath As String, ByVal ItemKey As String, ByVal vNewValue As Boolean)
    Dim PickableTable As PickableTable
    Set PickableTable = This.PickableTables.Item(ItemKey)
    '???
    'If PickableTable.Checked = vNewValue Then Exit Sub
    
    If vNewValue Then
        ' This item cannot be checked. Redraw the ListView using the Source collection.
        If Not PickableTable.Enabled Then
            PickableTable.Checked = False
            OnPropertyChanged "Item"
            Exit Sub
        End If
        ' This item can be checked.
        ' If something else was previous checked, set its PickableTable to unchecked.
        If Not This.Selected Is Nothing Then
            This.Selected.Checked = False
        End If
        ' Set this new item as the Selected PickableTable
        Set Selected = PickableTable
        PickableTable.Checked = vNewValue
        OnPropertyChanged "Item"
    Else
        ' Always allow unchecking.
        ' Set both the PickableTable's property and the Selected PickableTable property in Collection wrapper.
        'If Not This.Selected Is Nothing Then
        '    This.Selected.Checked = False
        '    Set Selected = Nothing
        'End If
        'PickableTable.Checked = vNewValue
        'OnPropertyChanged "Item"
    End If
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    
    Select Case PropertyName
        Case "EnableProtected"
            This.PickableTables.DisableProtectedWorksheets = CallByName(Source, PropertyName, VbGet)
        Case "EnableReadonly"
            This.PickableTables.DisableReadonlyWorksheets = CallByName(Source, PropertyName, VbGet)
        Case Else
            Exit Sub
    End Select
    
    Dim PreviousSelectedKey As String
    If Not Selected Is Nothing Then
        PreviousSelectedKey = Selected.Key
    End If
    Set Selected = Nothing
    
    This.PickableTables.Reload
    
    If PreviousSelectedKey <> vbNullString Then
        Dim PreviousSelected As PickableTable
        Set PreviousSelected = This.PickableTables.GetByKey(PreviousSelectedKey)
        If PreviousSelected.Enabled Then
            PreviousSelected.Checked = True
            Set Selected = PreviousSelected
        End If
    End If
    
    OnPropertyChanged "Item"
End Sub

'@Description "Sets the Selected PickableTable to the correct instance matching the given ListObject."
Private Sub SelectByListObject(ByVal ListObject As ListObject)
Attribute SelectByListObject.VB_Description = "Sets the Selected PickableTable to the correct instance matching the given ListObject."
    If ListObject Is Nothing Then Exit Sub
    Dim PickableTable As PickableTable
    Set PickableTable = This.PickableTables.GetByListObject(ListObject)
    If PickableTable Is Nothing Then Exit Sub
    PickableTable.Checked = True
    Set Selected = PickableTable
End Sub

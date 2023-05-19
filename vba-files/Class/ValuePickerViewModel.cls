VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValuePickerViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("MVVM.ValueMapper.ViewModel")
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged
Implements IHandlePropertyChanged

Private Type TState
    Notifier As INotifyPropertyChanged

    Caption As String
    Direction As TransferDirection
    
    Selected As ValueColumn
    ValueColumns As ValueColumns
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

Public Property Get Selected() As ValueColumn
    Set Selected = This.Selected
End Property

Public Property Set Selected(ByVal vNewValue As ValueColumn)
    If Not This.Selected Is Nothing Then
        If This.Selected Is vNewValue Then Exit Property
    End If
    Set This.Selected = vNewValue
    OnPropertyChanged "Selected"
    OnPropertyChanged "SelectedAsText"
End Property

Public Property Get Item() As Scripting.Dictionary
    Set Item = This.ValueColumns.Item
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
    ByVal Notifier As INotifyPropertyChanged) As ValuePickerViewModel
    
    Dim Result As ValuePickerViewModel
    Set Result = New ValuePickerViewModel
    Result.Load ListObject
    Result.Caption = Caption
    Result.Direction = Direction
    Notifier.RegisterHandler Result
    
    Set Create = Result
End Function

Public Sub Load(ByVal ListObject As ListObject)
    Debug.Assert Not Me Is KeyPickerViewModel
   
    Set This.ValueColumns = New ValueColumns
    This.ValueColumns.Load ListObject
End Sub

'---
Public Sub TrySetColumnChecked(ByVal PropertyPath As String, ByVal ItemKey As String, ByVal vNewValue As Boolean)
    Dim ValueColumn  As ValueColumn
    Set ValueColumn = This.ValueColumns.Item(ItemKey)
    
    If ValueColumn.Checked = vNewValue Then Exit Sub
    ValueColumn.Checked = vNewValue
    
    If vNewValue Then
        If Not ValueColumn.Enabled Then
            ValueColumn.Checked = False
            OnPropertyChanged "Item"
            Exit Sub
        End If
        If Not This.Selected Is Nothing Then
            This.Selected.Checked = False
        End If
        Set Selected = ValueColumn
        OnPropertyChanged "Item"
    Else
        This.Selected.Checked = False
        Set Selected = Nothing
        OnPropertyChanged "Item"
    End If
End Sub

Public Sub TrySetColumnSelected(ByVal PropertyPath As String, ByVal ItemKey As String, ByVal vNewValue As Boolean)
    Dim ValueColumn  As ValueColumn
    Set ValueColumn = This.ValueColumns.Item(ItemKey)
    
    If ValueColumn.Selected = vNewValue Then Exit Sub
    ValueColumn.Selected = vNewValue
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    Select Case PropertyName
        Case "EnableNontext"
            This.ValueColumns.EnableNontext = CallByName(Source, PropertyName, VbGet)
            ReloadValueColumns
        Case "EnableNonunique"
            This.ValueColumns.EnableNonunique = CallByName(Source, PropertyName, VbGet)
            ReloadValueColumns
    End Select
End Sub

Private Sub ReloadValueColumns()
    Dim PreviousSelected As String
    If Not Selected Is Nothing Then
        PreviousSelected = Selected.Key
    End If
    
    Set Selected = Nothing
    This.ValueColumns.Reload
    
    If PreviousSelected <> vbNullString Then
        Dim ValueColumn As ValueColumn
        Set ValueColumn = This.ValueColumns.Item(PreviousSelected)
        If ValueColumn.Enabled Then
            ValueColumn.Checked = True
            ValueColumn.Selected = True
            Set Selected = ValueColumn
        End If
    End If
    
    OnPropertyChanged "Item"
End Sub



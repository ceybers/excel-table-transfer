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
    
    KeyColumn As ListColumn
    SelectedKey As String
    SelectedValueColumn As ValueColumn
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

Public Property Get Item() As Scripting.Dictionary
    Set Item = This.ValueColumns.Item
End Property

Public Property Get SelectedKey() As String
    SelectedKey = This.SelectedKey
End Property

Public Property Let SelectedKey(ByVal vNewValue As String)
    If This.SelectedKey = vNewValue Then Exit Property
    This.SelectedKey = vNewValue
    OnPropertyChanged "SelectedKey"
    
    Set This.SelectedValueColumn = This.ValueColumns.GetByKey(vNewValue)
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
Public Function Create(ByVal ListColumn As ListColumn, _
    ByVal Caption As String, ByVal Direction As TransferDirection, _
    ByVal Notifier As INotifyPropertyChanged) As ValuePickerViewModel
    
    Dim Result As ValuePickerViewModel
    Set Result = New ValuePickerViewModel
    With Result
        .Load ListColumn
        .Caption = Caption
        .Direction = Direction
    End With
    Notifier.RegisterHandler Result
    
    Set Create = Result
End Function

Public Sub Load(ByVal ListColumn As ListColumn)
    Debug.Assert Not Me Is KeyPickerViewModel
   
    Set This.KeyColumn = ListColumn
    
    Set This.ValueColumns = New ValueColumns
    This.ValueColumns.Load ListColumn.Parent
    SelectFirstEligbleValueColumn
End Sub

Private Sub SelectFirstEligbleValueColumn()
    Dim Key As Variant
    For Each Key In This.ValueColumns.Item.Keys
        Dim ValueColumn As ValueColumn
        Set ValueColumn = This.ValueColumns.Item(Key)
        If ValueColumn.Enabled = True Then
            SelectedKey = Key
            Exit Sub
        End If
    Next Key
End Sub

'---

' This runs when the user changes a checkbox via the UserForm, which causes the primary VM to update
' one of its properties, which we are listening to here.
Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    Select Case PropertyName
        Case "EnableNontext"
            'This.ValueColumns.EnableNontext = CallByName(Source, PropertyName, VbGet)
            'ReloadValueColumns
        Case "EnableNonunique"
            'This.ValueColumns.EnableNonunique = CallByName(Source, PropertyName, VbGet)
            'ReloadValueColumns
    End Select
End Sub

'@Description "Reloads ValueColums. Required after changing options that affect the collection."
Private Sub ReloadValueColumns()
Attribute ReloadValueColumns.VB_Description = "Reloads ValueColums. Required after changing options that affect the collection."
    Dim PreviousSelected As String
    ' If Not Selected Is Nothing Then
        ' PreviousSelected = Selected.Key
    ' End If
    
    ' Set Selected = Nothing
    ' This.ValueColumns.Reload
    
    ' If PreviousSelected <> vbNullString Then
        ' Dim ValueColumn As ValueColumn
        ' Set ValueColumn = This.ValueColumns.Item(PreviousSelected)
        ' If ValueColumn.Enabled Then
            ' ValueColumn.Checked = True
            ' ValueColumn.Selected = True
            ' Set Selected = ValueColumn
        ' End If
    ' End If
    
    OnPropertyChanged "Item"
End Sub

Public Sub Map(ByVal SrcColumnKey As String)
    Debug.Assert This.Direction = Destination
    Debug.Assert Not This.SelectedValueColumn Is Nothing
    
    This.SelectedValueColumn.MappedTo = SrcColumnKey
    
    OnPropertyChanged "Item"
End Sub

Public Sub Unmap(ByVal DstColumnKey As String)
    Debug.Assert This.Direction = Destination
    'Debug.Assert Not This.Selected Is Nothing
    'Set This.Selected.MappedTo = Nothing
    
    This.ValueColumns.Item.Item(DstColumnKey).MappedTo = vbNullString
    
    OnPropertyChanged "Item"
End Sub

Public Sub TryUpdateSelected(ByVal ListItem As ListItem)
    Dim CurrentSelection As ValueColumn
    Set CurrentSelection = This.ValueColumns.GetByKey(ListItem.Key)
    
    Dim PreviousSelection As ValueColumn
    If SelectedKey <> vbNullString Then
        Set PreviousSelection = This.ValueColumns.GetByKey(SelectedKey)
    End If
    
    If Not CurrentSelection.Enabled Then
        If Not PreviousSelection Is Nothing Then
            PreviousSelection.Selected = True
        End If
        CurrentSelection.Selected = False
        ' SelectedKey doesn't change
    Else
        If Not PreviousSelection Is Nothing Then
            PreviousSelection.Selected = False
        End If
        ' CurrentSelection is already set to True
        SelectedKey = ListItem.Key
    End If
    
    OnPropertyChanged "Item"
End Sub

Public Function CanUnmapAll() As Boolean
    Dim Key As Variant
    For Each Key In This.ValueColumns.Item.Keys
        If This.ValueColumns.Item.Item(Key).MappedTo <> vbNullString Then
            CanUnmapAll = True
            Exit Function
        End If
    Next Key
End Function

Public Sub DoUnmapAll()
    Dim Key As Variant
    For Each Key In This.ValueColumns.Item.Keys
        This.ValueColumns.Item.Item(Key).MappedTo = vbNullString
    Next Key
    
    OnPropertyChanged "Item"
End Sub

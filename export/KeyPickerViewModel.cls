VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyPickerViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("MVVM.KeyMapper.ViewModel")
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged
Implements IHandlePropertyChanged

Private Type TState
    Notifier As INotifyPropertyChanged
    Context As IAppContext
    Properties As TablePropViewModel
    
    Caption As String
    Direction As TransferDirection
    
    Selected As KeyColumn
    KeyColumns As KeyColumns
    ListObject As ListObject
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
Public Property Get TableName() As String
    TableName = This.ListObject.Name & " (" _
        & This.ListObject.Parent.Name & ") (" _
        & This.ListObject.Parent.Parent.Name & ")"
End Property

Public Property Get SelectedAsText() As String
    If This.Selected Is Nothing Then
        SelectedAsText = "nothing"
        Exit Property
    End If
    SelectedAsText = This.Selected.Name
End Property

Public Property Get Selected() As KeyColumn
    Set Selected = This.Selected
End Property

Public Property Set Selected(ByVal vNewValue As KeyColumn)
    If Not This.Selected Is Nothing Then
        If This.Selected Is vNewValue Then Exit Property
    End If
    Set This.Selected = vNewValue
    OnPropertyChanged "Selected"
    OnPropertyChanged "SelectedAsText"
End Property

Public Property Get Item() As Scripting.Dictionary
    Set Item = This.KeyColumns.Item
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

Public Property Get HasPreferred() As Boolean
    HasPreferred = This.KeyColumns.HasPreferred
End Property

'---
Public Function Create(ByVal Context As IAppContext, ByVal ListObject As Object, _
    ByVal Caption As String, ByVal Direction As TransferDirection, _
    ByVal Notifier As INotifyPropertyChanged) As KeyPickerViewModel
    
    Dim Result As KeyPickerViewModel
    Set Result = New KeyPickerViewModel
    With Result
        .Load Context, ListObject
        .Caption = Caption
        .Direction = Direction
    End With
    Notifier.RegisterHandler Result
    
    Set Create = Result
End Function

Public Sub Load(ByVal Context As IAppContext, ByVal ListObject As ListObject)
    Debug.Assert Not Me Is KeyPickerViewModel
   
    Set This.Context = Context
    Set This.ListObject = ListObject
    
    Set This.Properties = New TablePropViewModel
    This.Properties.Load This.Context, This.ListObject
    
    Set This.KeyColumns = New KeyColumns
    This.KeyColumns.Load ListObject
    
    This.KeyColumns.UpdatePreferredColumn This.Properties.TablePreferKeyVM.Selected
    ReloadKeyColumns
End Sub

'---
Public Sub TrySetColumnChecked(ByVal PropertyPath As String, ByVal ItemKey As String, ByVal vNewValue As Boolean)
    Dim KeyColumn  As KeyColumn
    Set KeyColumn = This.KeyColumns.Item(ItemKey)
    
    If KeyColumn.Checked = vNewValue Then Exit Sub
    KeyColumn.Checked = vNewValue
    
    If vNewValue Then
        If Not KeyColumn.Enabled Then
            KeyColumn.Checked = False
            OnPropertyChanged "Item"
            Exit Sub
        End If
        
        If Not This.Selected Is Nothing Then
            This.Selected.Checked = False
        End If
        
        Set Selected = KeyColumn
        OnPropertyChanged "Item"
    Else
        This.Selected.Checked = False
        Set Selected = Nothing
        OnPropertyChanged "Item"
    End If
End Sub

Public Sub TrySetColumnSelected(ByVal PropertyPath As String, ByVal ItemKey As String, ByVal vNewValue As Boolean)
    Dim KeyColumn  As KeyColumn
    Set KeyColumn = This.KeyColumns.Item(ItemKey)
    
    If KeyColumn.Selected = vNewValue Then Exit Sub
    KeyColumn.Selected = vNewValue
End Sub

Private Sub IHandlePropertyChanged_HandlePropertyChanged(ByVal Source As Object, ByVal PropertyName As String)
    Select Case PropertyName
        Case "EnableNontext"
            This.KeyColumns.EnableNontext = CallByName(Source, PropertyName, VbGet)
            ReloadKeyColumns
        Case "EnableNonunique"
            This.KeyColumns.EnableNonunique = CallByName(Source, PropertyName, VbGet)
            ReloadKeyColumns
    End Select
End Sub

Private Sub ReloadKeyColumns()
    Dim PreviousSelected As String
    If Not Selected Is Nothing Then
        PreviousSelected = Selected.Key
    End If
    
    Set Selected = Nothing
    This.KeyColumns.Reload
    
    If PreviousSelected <> vbNullString Then
        Dim KeyColumn As KeyColumn
        Set KeyColumn = This.KeyColumns.Item(PreviousSelected)
        If KeyColumn.Enabled Then
            KeyColumn.Checked = True
            KeyColumn.Selected = True
            Set Selected = KeyColumn
        End If
    End If
    
    OnPropertyChanged "Item"
End Sub

Public Function TryGetByKey(ByVal Key As String, ByRef OutKeyColumn As KeyColumn) As Boolean
    TryGetByKey = DictionaryEx.From(This.KeyColumns.Item).TryGetByKey(Key, OutKeyColumn)
End Function

Public Sub TryMap(ByVal Key As String)
    If Not Selected Is Nothing Then
        Selected.Checked = False
    End If
    
    Dim SelKeyColumn As KeyColumn
    Set SelKeyColumn = This.KeyColumns.Item.Item(Key)
    SelKeyColumn.Checked = True
    Set Selected = SelKeyColumn
    
    OnPropertyChanged "Item"
End Sub

Public Sub Deselect()
    If This.Selected Is Nothing Then Exit Sub
    This.Selected.Checked = False
    Set Selected = Nothing
    
    OnPropertyChanged "Item"
End Sub

Public Sub ShowTableProperties()
    Dim TablePropV As IView
    Set TablePropV = TablePropView.Create(This.Context, This.Properties)
    
    If TablePropV.ShowDialog Then
        This.Properties.Commit
        This.KeyColumns.UpdatePreferredColumn This.Properties.TablePreferKeyVM.Selected
        ReloadKeyColumns
    End If
    
    Set TablePropV = Nothing
End Sub

Public Sub SetToPreferred()
    If Not This.KeyColumns.HasPreferred Then Exit Sub
    
    Dim AKeyColumn As KeyColumn
    If TryGetByKey(This.KeyColumns.PreferredColumnName, AKeyColumn) Then
        With AKeyColumn
            .Enabled = True
            .IsPreferred = True
            .Checked = False
        End With
    End If
    
    TrySetColumnChecked vbEmpty, This.KeyColumns.PreferredColumnName, True
End Sub

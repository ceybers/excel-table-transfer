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
    Context As IAppContext
    Properties As TablePropViewModel
    
    ListObject As ListObject
    
    Caption As String
    Direction As TransferDirection
    
    KeyColumn As ListColumn
    SelectedKey As String
    SelectedValueColumn As ValueColumn
    ValueColumns As ValueColumns
    
    StarredColumns As Collection
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

Public Property Get TableName() As String
    TableName = This.ListObject.Name & " (" & This.ListObject.Parent.Name & ") (" & This.ListObject.Parent.Parent.Name & ")"
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
Public Function Create(ByVal Context As IAppContext, ByVal ListColumn As ListColumn, _
    ByVal Caption As String, ByVal Direction As TransferDirection, _
    ByVal Notifier As INotifyPropertyChanged) As ValuePickerViewModel
    
    Dim Result As ValuePickerViewModel
    Set Result = New ValuePickerViewModel
    With Result
        .Load Context, ListColumn
        .Caption = Caption
        .Direction = Direction
    End With
    Notifier.RegisterHandler Result
    
    Set Create = Result
End Function

Public Sub Load(ByVal Context As IAppContext, ByVal ListColumn As ListColumn)
    Debug.Assert Not Me Is KeyPickerViewModel
   
    Set This.Context = Context
    Set This.ListObject = ListColumn.Parent
    Set This.KeyColumn = ListColumn
    
    Set This.ValueColumns = New ValueColumns
    This.ValueColumns.Load ListColumn
    
    Set This.Properties = New TablePropViewModel
    This.Properties.Load This.Context, This.ListObject
    
    SelectFirstEligbleValueColumn
    
    UpdateStarredColumns
    OnPropertyChanged "Item"
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
    Debug.Print PropertyName
    Select Case PropertyName
        Case "EnableStarredOnly"
            This.ValueColumns.ShowStarredOnly = CallByName(Source, PropertyName, VbGet)
            ReloadValueColumns
        Case "EnableNonunique"
            'This.ValueColumns.EnableNonunique = CallByName(Source, PropertyName, VbGet)
            'ReloadValueColumns
    End Select
End Sub

'@Description "Reloads ValueColums. Required after changing option buttons on the main ViewModel that affect the collection."
Private Sub ReloadValueColumns()
Attribute ReloadValueColumns.VB_Description = "Reloads ValueColums. Required after changing options that affect the collection."
    Dim PreviousSelected As String
    'If Not Selected Is Nothing Then
    '    PreviousSelected = Selected.Key
    'End If
    
    'Set Selected = Nothing
    This.ValueColumns.Reload
    
    'If PreviousSelected <> vbNullString Then
    '    Dim ValueColumn As ValueColumn
    '    Set ValueColumn = This.ValueColumns.Item(PreviousSelected)
    '    If ValueColumn.Enabled Then
    '        ValueColumn.Checked = True
    '        ValueColumn.Selected = True
    '        Set Selected = ValueColumn
    '    End If
    'End If
    
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

Public Function GetMappedColumnTuples(ByVal SrcValueColumns As Scripting.Dictionary) As Collection
    Dim Result As Collection
    Set Result = New Collection
    
    Dim Key As Variant
    For Each Key In This.ValueColumns.Item.Keys
        Dim DstValueColumn As ValueColumn
        Set DstValueColumn = This.ValueColumns.Item.Item(Key)
        If DstValueColumn.MappedTo <> vbNullString Then
            Dim SrcValueColumn As ValueColumn
            Set SrcValueColumn = SrcValueColumns.Item(DstValueColumn.MappedTo)
            Dim AColumnTuple As ColumnTuple
            Set AColumnTuple = ColumnTuple.Create(SrcValueColumn.ListColumn, DstValueColumn.ListColumn)
            Result.Add Item:=AColumnTuple
        End If
    Next Key
    
    Set GetMappedColumnTuples = Result
End Function

Public Function GetEnabledColumnNames() As Collection
    Dim Result As Collection
    Set Result = New Collection
    
    Dim Key As Variant
    For Each Key In This.ValueColumns.Item
        Dim ThisValueColumn  As ValueColumn
        Set ThisValueColumn = This.ValueColumns.Item(Key)
        If ThisValueColumn.Enabled Then
            Result.Add Item:=ThisValueColumn.Name, Key:=ThisValueColumn.Name
        End If
    Next Key
    Set GetEnabledColumnNames = Result
End Function

Public Sub AutoMap(ByVal SrcValueColumns As Collection)
    Debug.Assert This.Direction = Destination
    
    Dim SrcColumnName As Variant
    For Each SrcColumnName In SrcValueColumns
        If This.ValueColumns.Item.Exists(SrcColumnName) Then
            Dim DstValueColumn As ValueColumn
            Set DstValueColumn = This.ValueColumns.Item.Item(SrcColumnName)
            
            If DstValueColumn.Enabled Then
                DstValueColumn.MappedTo = SrcColumnName
            Else
                DstValueColumn.MappedTo = vbNullString
            End If
        End If
    Next SrcColumnName
        
    OnPropertyChanged "Item"
End Sub

Public Function IsAutoMapped(ByVal SrcValueColumns As Collection) As Boolean
    IsAutoMapped = True
    
    Dim SrcColumnName As Variant
    For Each SrcColumnName In SrcValueColumns
        If This.ValueColumns.Item.Exists(SrcColumnName) Then
            Dim DstValueColumn As ValueColumn
            Set DstValueColumn = This.ValueColumns.Item.Item(SrcColumnName)
            
            If DstValueColumn.Enabled Then
                If DstValueColumn.MappedTo <> SrcColumnName Then
                    IsAutoMapped = False
                    Exit Function
                End If
            Else
                If DstValueColumn.MappedTo <> vbNullString Then
                    IsAutoMapped = False
                    Exit Function
                End If
            End If
        End If
    Next SrcColumnName
End Function

Public Sub ShowTableProperties()
    Dim TablePropV As IView
    Set TablePropV = TablePropView.Create(This.Context, This.Properties)
    
    If TablePropV.ShowDialog Then
        This.Properties.Commit
        UpdateStarredColumns
        'This.KeyColumns.UpdatePreferredColumn This.Properties.TablePreferKeyVM.Selected
        'ReloadKeyColumns
    End If
    
    Set TablePropV = Nothing
End Sub

Private Sub UpdateStarredColumns()
    ' TODO Do we need to store a copy here?
    Set This.StarredColumns = This.Properties.TableStarColumnsVM.GetStarredColumns
    This.ValueColumns.UpdateStarredColumns This.StarredColumns
    
    OnPropertyChanged "Item"
End Sub

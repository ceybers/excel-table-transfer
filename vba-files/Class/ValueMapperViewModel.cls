VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValueMapperViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.ValueMapper.ViewModel"
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged

Private Type TState
    Notifier As INotifyPropertyChanged

    SrcValueColumnVM As ValuePickerViewModel
    DstValueColumnVM As ValuePickerViewModel
    MappedValueColumns As Collection ' of type ColumnTuple
End Type
Private This As TState

Private Function IViewModel_IsValid() As Boolean
    'If This.SrcValueColumnVM.Selected Is Nothing Then Exit Function
    'If This.DstValueColumnVM.Selected Is Nothing Then Exit Function
    
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
Public Property Get SrcValueColumnVM() As ValuePickerViewModel
    Set SrcValueColumnVM = This.SrcValueColumnVM
End Property

Public Property Get DstValueColumnVM() As ValuePickerViewModel
    Set DstValueColumnVM = This.DstValueColumnVM
End Property

Public Property Get MappedValueColumns() As Collection
    Set MappedValueColumns = This.MappedValueColumns
End Property

'---
'Public Sub Load(ByVal SrcTable As ListObject, ByVal DstTable As ListObject)
Public Sub Load(ByVal SrcColumn As ListColumn, ByVal DstColumn As ListColumn)
    Set This.SrcValueColumnVM = ValuePickerViewModel.Create(SrcColumn, "Value Columns for Source Table", _
        Source, This.Notifier)
    Set This.DstValueColumnVM = ValuePickerViewModel.Create(DstColumn, "Value Columns for Destination Table", _
        Destination, This.Notifier)
End Sub

Public Function CanMap() As Boolean
    If This.SrcValueColumnVM.SelectedKey = vbNullString Then Exit Function
    If This.DstValueColumnVM.SelectedKey = vbNullString Then Exit Function
    
    CanMap = True
End Function

Public Function CanUnmap() As Boolean
    If This.DstValueColumnVM.SelectedKey = vbNullString Then Exit Function
    'If This.DstValueColumnVM.Selected.MappedTo = vbNullString Then Exit Function
    If This.DstValueColumnVM.Item.Item(This.DstValueColumnVM.SelectedKey).MappedTo = vbNullString Then Exit Function
    
    CanUnmap = True
End Function

Public Function CanUnmapAll() As Boolean
    CanUnmapAll = This.DstValueColumnVM.CanUnmapAll()
End Function

Public Function CanAutoMap() As Boolean
    CanAutoMap = Not This.DstValueColumnVM.IsAutoMapped(This.SrcValueColumnVM.Item)
End Function

Public Sub DoMap()
    'Dim SrcValueColumn As ValueColumn
    This.DstValueColumnVM.Map This.SrcValueColumnVM.SelectedKey
End Sub

Public Sub DoUnmap()
    This.DstValueColumnVM.Unmap This.DstValueColumnVM.SelectedKey
End Sub

Public Sub DoUnmapAll()
    This.DstValueColumnVM.DoUnmapAll
End Sub

Public Sub DoAutoMap()
    This.DstValueColumnVM.AutoMap This.SrcValueColumnVM.Item
End Sub

Public Sub ExtractMappedData()
    Set This.MappedValueColumns = This.DstValueColumnVM.GetMappedColumnTuples(This.SrcValueColumnVM.Item)
End Sub

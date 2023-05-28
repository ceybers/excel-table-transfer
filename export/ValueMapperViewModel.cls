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

Private Const ENABLE_STARREDONLY_FLAG As String = "VALMAPPER_ENABLESTARREDONLY"
Private Const ENABLE_OPTION2_FLAG As String = "VALMAPPER_ENABLEOPTION2"

Private Type TState
    Notifier As INotifyPropertyChanged
    Context As AppContext

    SrcValueColumnVM As ValuePickerViewModel
    DstValueColumnVM As ValuePickerViewModel
    MappedValueColumns As Collection ' of type ColumnTuple
    
    EnableStarredOnly As Boolean
    EnableOption2 As Boolean
End Type
Private This As TState

Private Function IViewModel_IsValid() As Boolean
    'If This.SrcValueColumnVM.Selected Is Nothing Then Exit Function
    'If This.DstValueColumnVM.Selected Is Nothing Then Exit Function
    If This.DstValueColumnVM.GetMappedColumnTuples(This.SrcValueColumnVM.Item).Count = 0 Then Exit Function
    
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

Public Property Get EnableStarredOnly() As Boolean
    EnableStarredOnly = This.EnableStarredOnly
End Property

Public Property Let EnableStarredOnly(ByVal vNewValue As Boolean)
    If This.EnableStarredOnly = vNewValue Then Exit Property
    This.EnableStarredOnly = vNewValue
    OnPropertyChanged "EnableStarredOnly"
    This.Context.Settings.User.SetFlag ENABLE_STARREDONLY_FLAG, vNewValue
End Property

Public Property Get EnableOption2() As Boolean
    EnableOption2 = This.EnableOption2
End Property

Public Property Let EnableOption2(ByVal vNewValue As Boolean)
    If This.EnableOption2 = vNewValue Then Exit Property
    This.EnableOption2 = vNewValue
    OnPropertyChanged "EnableOption2"
    This.Context.Settings.User.SetFlag ENABLE_OPTION2_FLAG, vNewValue
End Property

'---
'Public Sub Load(ByVal SrcTable As ListObject, ByVal DstTable As ListObject)
Public Sub Load(ByVal Context As IAppContext, ByVal SrcColumn As ListColumn, ByVal DstColumn As ListColumn)
    Set This.Context = Context
    
    Set This.SrcValueColumnVM = ValuePickerViewModel.Create(Context, SrcColumn, _
        "Value Columns for Source Table", Source, This.Notifier)
    Set This.DstValueColumnVM = ValuePickerViewModel.Create(Context, DstColumn, _
        "Value Columns for Destination Table", Destination, This.Notifier)
        
    Me.EnableStarredOnly = This.Context.Settings.User.GetFlag(ENABLE_STARREDONLY_FLAG)
    Me.EnableOption2 = This.Context.Settings.User.GetFlag(ENABLE_OPTION2_FLAG)
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
    CanAutoMap = Not This.DstValueColumnVM.IsAutoMapped(This.SrcValueColumnVM.GetEnabledColumnNames)
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
    'This.DstValueColumnVM.AutoMap This.SrcValueColumnVM.Item
    This.DstValueColumnVM.AutoMap This.SrcValueColumnVM.GetEnabledColumnNames
End Sub

Public Sub ExtractMappedData()
    Set This.MappedValueColumns = This.DstValueColumnVM.GetMappedColumnTuples(This.SrcValueColumnVM.Item)
End Sub



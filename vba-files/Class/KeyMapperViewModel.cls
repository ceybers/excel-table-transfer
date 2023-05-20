VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyMapperViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.KeyMapper.ViewModel"
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged

Private Const ENABLE_NONTEXT_FLAG As String = "KEYMAPPER_ENABLENONTEXT"
Private Const ENABLE_NONUNIQUE_FLAG As String = "KEYMAPPER_ENABLENONUNIQUE"

Private Type TState
    Notifier As INotifyPropertyChanged
    Context As AppContext

    SrcKeyColumnVM As KeyPickerViewModel
    DstKeyColumnVM As KeyPickerViewModel
    
    EnableNontext As Boolean
    EnableNonunique As Boolean
End Type
Private This As TState

Private Function IViewModel_IsValid() As Boolean
    If This.SrcKeyColumnVM.Selected Is Nothing Then Exit Function
    If This.DstKeyColumnVM.Selected Is Nothing Then Exit Function
    
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
Public Property Get SrcKeyColumnVM() As KeyPickerViewModel
    Set SrcKeyColumnVM = This.SrcKeyColumnVM
End Property

Public Property Get DstKeyColumnVM() As KeyPickerViewModel
    Set DstKeyColumnVM = This.DstKeyColumnVM
End Property

Public Property Get EnableNontext() As Boolean
    EnableNontext = This.EnableNontext
End Property

Public Property Let EnableNontext(ByVal vNewValue As Boolean)
    If This.EnableNontext = vNewValue Then Exit Property
    This.EnableNontext = vNewValue
    OnPropertyChanged "EnableNontext"
    This.Context.UserLevelSettings.SetFlag ENABLE_NONTEXT_FLAG, vNewValue
End Property

Public Property Get EnableNonunique() As Boolean
    EnableNonunique = This.EnableNonunique
End Property

Public Property Let EnableNonunique(ByVal vNewValue As Boolean)
    If This.EnableNonunique = vNewValue Then Exit Property
    This.EnableNonunique = vNewValue
    OnPropertyChanged "EnableNonunique"
    This.Context.UserLevelSettings.SetFlag ENABLE_NONUNIQUE_FLAG, vNewValue
End Property

'---
Public Sub Load(ByVal Context As IAppContext, ByVal SrcTable As ListObject, ByVal DstTable As ListObject)
    Set This.Context = Context
    Set This.SrcKeyColumnVM = KeyPickerViewModel.Create(SrcTable, "Key Column for Source Table", _
        Source, This.Notifier)
    Set This.DstKeyColumnVM = KeyPickerViewModel.Create(DstTable, "Key Column for Destination Table", _
        Destination, This.Notifier)
        
    Me.EnableNontext = This.Context.UserLevelSettings.GetFlag(ENABLE_NONTEXT_FLAG)
    Me.EnableNonunique = This.Context.UserLevelSettings.GetFlag(ENABLE_NONUNIQUE_FLAG)
End Sub

Public Function CanAutoMap() As Boolean
    CanAutoMap = True
End Function

Public Sub DoAutoMap()
    MsgBox "DoAutoMap Not Yet Implemented"
End Sub

Public Function CanMapPreferred() As Boolean
    CanMapPreferred = True
End Function

Public Sub DoMapPreferred()
    MsgBox "DoMapPreferred Not Yet Implemented"
End Sub

Public Function CanUnmapAll() As Boolean
    If Not This.SrcKeyColumnVM.Selected Is Nothing Then CanUnmapAll = True
    If Not This.DstKeyColumnVM.Selected Is Nothing Then CanUnmapAll = True
End Function

Public Sub DoUnmapAll()
    This.SrcKeyColumnVM.Deselect
    This.DstKeyColumnVM.Deselect
End Sub

Public Function CanMap(ByVal Direction As TransferDirection) As Boolean
    If This.SrcKeyColumnVM.SelectedAsText = This.DstKeyColumnVM.SelectedAsText Then Exit Function
    
    Dim CheckColumn As KeyColumn
    If Direction = TransferDirection.Source Then
        If This.SrcKeyColumnVM.Selected Is Nothing Then Exit Function
        If This.DstKeyColumnVM.TryGetByKey(This.SrcKeyColumnVM.SelectedAsText, CheckColumn) Then
            If Not CheckColumn.Enabled Then Exit Function
            If CheckColumn.Selected Then Exit Function
        Else
            Exit Function
        End If
    End If
    
    If Direction = TransferDirection.Destination Then
        If This.DstKeyColumnVM.Selected Is Nothing Then Exit Function
        If This.SrcKeyColumnVM.TryGetByKey(This.DstKeyColumnVM.SelectedAsText, CheckColumn) Then
            If Not CheckColumn.Enabled Then Exit Function
            If CheckColumn.Selected Then Exit Function
        Else
            Exit Function
        End If
    End If
    
    CanMap = True
End Function

Public Sub DoMap(ByVal Direction As TransferDirection)
    If Direction = TransferDirection.Source Then
        This.DstKeyColumnVM.TryMap This.SrcKeyColumnVM.SelectedAsText
    ElseIf Direction = TransferDirection.Destination Then
        This.SrcKeyColumnVM.TryMap This.DstKeyColumnVM.SelectedAsText
    End If
End Sub

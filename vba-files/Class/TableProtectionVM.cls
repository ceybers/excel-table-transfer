VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableProtectionVM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVVM.TableProps.ViewModel")
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged
Implements IChangeTracking

Private Const PROTECTION_ISNOCHANGE As String = "PROTECTION_ISNOCHANGE"
Private Const PROTECTION_ISTEMPORARY As String = "PROTECTION_ISTEMPORARY"
Private Const PROTECTION_ISPERMANENT As String = "PROTECTION_ISPERMANENT"
Private Const PROTECTION_ISPROTECTED As String = "PROTECTION_ISPROTECTED"

Private Type TState
    Notifier As INotifyPropertyChanged
    Settings As ISettings
    IsChanged As Boolean
    
    IsNoChange As Boolean
    IsTemporary As Boolean
    IsPermanent As Boolean
    IsTableProtected As Boolean
    IsPasswordProtected As Boolean ' Read-only

    LockTableCommand As ICommand
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
    This.IsChanged = True
End Sub

Private Sub Class_Initialize()
    Set This.Notifier = New PropertyChangeNotifier
End Sub

'---
Public Property Get IsNoChange() As Boolean
    IsNoChange = This.IsNoChange
End Property

Public Property Let IsNoChange(ByVal vNewValue As Boolean)
    If This.IsNoChange = vNewValue Then Exit Property
    This.IsNoChange = vNewValue
    OnPropertyChanged "IsNoChange"
    
    If vNewValue Then
        IsTemporary = False
        IsPermanent = False
    End If
End Property

Public Property Get IsTemporary() As Boolean
    IsTemporary = This.IsTemporary
End Property

Public Property Let IsTemporary(ByVal vNewValue As Boolean)
    If This.IsTemporary = vNewValue Then Exit Property
    This.IsTemporary = vNewValue
    OnPropertyChanged "IsTemporary"
    
    If vNewValue Then
        IsNoChange = False
        IsPermanent = False
    End If
End Property

Public Property Get IsPermanent() As Boolean
    IsPermanent = This.IsPermanent
End Property

Public Property Let IsPermanent(ByVal vNewValue As Boolean)
    If This.IsPermanent = vNewValue Then Exit Property
    This.IsPermanent = vNewValue
    OnPropertyChanged "IsPermanent"
    
    If vNewValue Then
        IsNoChange = False
        IsTemporary = False
    End If
End Property

Public Property Get IsTableProtected() As Boolean
    IsTableProtected = This.IsTableProtected
End Property

Public Property Let IsTableProtected(ByVal vNewValue As Boolean)
    If This.IsTableProtected = vNewValue Then Exit Property
    This.IsTableProtected = vNewValue
    OnPropertyChanged "IsTableProtected"
End Property

' ---
Public Sub Load(ByVal ListObject As ListObject, ByVal Settings As ISettings)
    Set This.Settings = Settings
    
    'Set LockTableCommand = New LockTableCommand(...)

    LoadSavedFlags
End Sub

Private Sub LoadSavedFlags()
    IsNoChange = This.Settings.GetFlag(PROTECTION_ISNOCHANGE)
    IsTemporary = This.Settings.GetFlag(PROTECTION_ISTEMPORARY)
    IsPermanent = This.Settings.GetFlag(PROTECTION_ISPERMANENT)
    IsTableProtected = This.Settings.GetFlag(PROTECTION_ISPROTECTED)
    
    If IsNoChange = False And IsTemporary = False And IsPermanent = False Then
        IsNoChange = True
    End If
End Sub

Public Sub Commit()
    If This.IsChanged = False Then Exit Sub
    This.Settings.SetFlag PROTECTION_ISNOCHANGE, IsNoChange
    This.Settings.SetFlag PROTECTION_ISTEMPORARY, IsTemporary
    This.Settings.SetFlag PROTECTION_ISPERMANENT, IsPermanent
    This.Settings.SetFlag PROTECTION_ISPROTECTED, IsTableProtected
End Sub

Private Property Get IChangeTracking_IsChanged() As Boolean
   IChangeTracking_IsChanged = This.IsChanged
End Property

Private Sub IChangeTracking_AcceptChanges()
    Commit
    This.IsChanged = False
End Sub


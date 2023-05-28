VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableProfileVM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVVM.TableProps.ViewModel")
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged

Private Const RESET_PROFILE_MESSAGE As String = "Are you sure you want to reset the profile?"
Private Const RESET_PROFILE_TITLE As String = "Reset Profile"

Private Type TState
    Notifier As INotifyPropertyChanged
    SettingsModel As ISettingsModel
    
    HasProfile As Boolean
    ActivateProfileCommand As ICommand
    RemoveProfileCommand As ICommand
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

Public Property Get HasProfile() As Boolean
    HasProfile = This.HasProfile ' True 'check settings
End Property

Public Property Get ActivateProfileCommand() As ICommand
    Set ActivateProfileCommand = This.ActivateProfileCommand
End Property

Public Property Set ActivateProfileCommand(ByVal vNewValue As ICommand)
    Set This.ActivateProfileCommand = vNewValue
End Property

Public Property Get RemoveProfileCommand() As ICommand
    Set RemoveProfileCommand = This.RemoveProfileCommand
End Property

Public Property Set RemoveProfileCommand(ByVal vNewValue As ICommand)
    Set This.RemoveProfileCommand = vNewValue
End Property

' ---
Public Sub Load(ByVal ListObject As ListObject, ByVal SettingsModel As ISettingsModel)
    Set This.SettingsModel = SettingsModel
    This.HasProfile = True 'TODO Impl function to actually check first
End Sub

Public Sub ActivateProfile()
    'MsgBox "ActivateProfile NYI"
    This.HasProfile = Not This.HasProfile
    OnPropertyChanged "HasProfile" ' Name doesn't matter, it just needs to trigger CommandManager's update CanExecute
    ' And we need to at least do a OneTimeBinding so that BindingManager will have seen this sub-VM at least once
End Sub

Public Sub RemoveProfile()
    If vbYes <> MsgBox(RESET_PROFILE_MESSAGE, vbExclamation + vbYesNo + vbDefaultButton2, RESET_PROFILE_TITLE) Then Exit Sub
    
    This.SettingsModel.Workbook.Reset
    
    This.HasProfile = Not This.HasProfile
    OnPropertyChanged "HasProfile" ' Name doesn't matter, it just needs to trigger CommandManager's update CanExecute
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableDirectionVM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("MVVM.TableProps.ViewModel")
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged
Implements IChangeTracking

Private Const DIRECTION_ISNEITHER As String = "DIRECTION_ISNEITHER"
Private Const DIRECTION_ISSOURCE As String = "DIRECTION_ISSOURCE"
Private Const DIRECTION_ISDESTIATION As String = "DIRECTION_ISDESTIATION"
Private Const DIRECTION_CONDITION As String = "DIRECTION_CONDITION"
Private Const DIRECTION_CONDITION_CRITERIA As String = "DIRECTION_CONDITION_CRITERIA"

Private Type TState
    Notifier As INotifyPropertyChanged
    Settings As ISettings
    IsChanged As Boolean
    
    IsNeither As Boolean
    IsSource As Boolean
    IsDestination As Boolean
    IsOnCondition As Boolean
    
    SelectedConditionDirection As String
    ConditionDirections As Scripting.Dictionary
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
    Set This.ConditionDirections = New Scripting.Dictionary
End Sub

'---
Public Property Get IsNeither() As Boolean
    IsNeither = This.IsNeither
End Property

Public Property Let IsNeither(ByVal vNewValue As Boolean)
    If This.IsNeither = vNewValue Then Exit Property
    This.IsNeither = vNewValue
    OnPropertyChanged "IsNeither"
    
    If vNewValue Then
        IsSource = False
        IsDestination = False
    End If
End Property

Public Property Get IsSource() As Boolean
    IsSource = This.IsSource
End Property

Public Property Let IsSource(ByVal vNewValue As Boolean)
    If This.IsSource = vNewValue Then Exit Property
    This.IsSource = vNewValue
    OnPropertyChanged "IsSource"
    
    If vNewValue Then
        IsNeither = False
        IsDestination = False
    End If
End Property

Public Property Get IsDestination() As Boolean
    IsDestination = This.IsDestination
End Property

Public Property Let IsDestination(ByVal vNewValue As Boolean)
    If This.IsDestination = vNewValue Then Exit Property
    This.IsDestination = vNewValue
    OnPropertyChanged "IsDestination"
    
    If vNewValue Then
        IsNeither = False
        IsSource = False
    End If
End Property

Public Property Get IsOnCondition() As Boolean
    IsOnCondition = This.IsOnCondition
End Property

Public Property Let IsOnCondition(ByVal vNewValue As Boolean)
    If This.IsOnCondition = vNewValue Then Exit Property
    This.IsOnCondition = vNewValue
    OnPropertyChanged "IsOnCondition"
End Property

Public Property Get SelectedConditionDirection() As String
    SelectedConditionDirection = This.SelectedConditionDirection
End Property

Public Property Let SelectedConditionDirection(ByVal vNewValue As String)
    If This.SelectedConditionDirection = vNewValue Then Exit Property
    This.SelectedConditionDirection = vNewValue
    OnPropertyChanged "SelectedConditionDirection"
End Property

Public Property Get ConditionDirections() As Scripting.Dictionary
    Set ConditionDirections = This.ConditionDirections
End Property

' ---
Public Sub Load(ByVal ListObject As ListObject, ByVal Settings As ISettings)
    Set This.Settings = Settings
    
    LoadSavedFlags
    LoadConditionDirections
    LoadSelectedConditionDirection
End Sub

Private Sub LoadSavedFlags()
    IsNeither = This.Settings.GetFlag(DIRECTION_ISNEITHER)
    IsSource = This.Settings.GetFlag(DIRECTION_ISSOURCE)
    IsDestination = This.Settings.GetFlag(DIRECTION_ISDESTIATION)
    IsOnCondition = This.Settings.GetFlag(DIRECTION_CONDITION)
    
    If IsNeither = False And IsSource = False And IsDestination = False Then
        IsNeither = True
    End If
End Sub

Private Sub LoadConditionDirections()
    Dim Tokens() As String
    Tokens = Split("Local,Network,OneDrive,SharePoint", ",")
    Dim i As Long
    For i = LBound(Tokens) To UBound(Tokens)
        This.ConditionDirections.Add Key:=Tokens(i), Item:=Tokens(i)
    Next i
    
    SelectedConditionDirection = Tokens(LBound(Tokens))
End Sub

Private Sub LoadSelectedConditionDirection()
    Dim SavedCriteria As String
    SavedCriteria = This.Settings.GetSetting(DIRECTION_CONDITION_CRITERIA)
    
    Dim Key As Variant
    For Each Key In This.ConditionDirections.Keys
        If This.ConditionDirections.Item(Key) = SavedCriteria Then
                SelectedConditionDirection = This.ConditionDirections.Item(Key)
            Exit Sub
        End If
    Next Key
End Sub

Public Sub Commit()
    If This.IsChanged = False Then Exit Sub
    This.Settings.SetFlag DIRECTION_ISNEITHER, IsNeither
    This.Settings.SetFlag DIRECTION_ISSOURCE, IsSource
    This.Settings.SetFlag DIRECTION_ISDESTIATION, IsDestination
    This.Settings.SetFlag DIRECTION_CONDITION, IsOnCondition
    This.Settings.SetSetting DIRECTION_CONDITION_CRITERIA, SelectedConditionDirection
End Sub

Private Property Get IChangeTracking_IsChanged() As Boolean
   IChangeTracking_IsChanged = This.IsChanged
End Property

Private Sub IChangeTracking_AcceptChanges()
    If This.IsChanged = False Then Exit Sub
    Commit
    This.IsChanged = False
End Sub


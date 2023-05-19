VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TablePropViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.TableProps.ViewModel"
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged

Private Const SETTINGS_NODE As String = "TableTransferTool"

Private Type TState
    Notifier As INotifyPropertyChanged
    TableSettings As ISettings
    
    ' General tab
    TableDetailsVM As TableDetailsVM
    TableLocationVM As TableLocationVM
    TableDirectionVM As TableDirectionVM

    ' Columns tab
    TablePreferKeyVM As TablePreferKeyVM
    TableStarColumnsVM As TableStarColumnsVM
    
    ' History tab
    TableProfileVM As TableProfileVM

    ' Tracking Tab
    TableHighlightingVM As TableHighlightingVM
    TableTimestampVM As TableTimestampVM
    TableProtectionVM As TableProtectionVM
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

    Set This.TableDetailsVM = New TableDetailsVM
    Set This.TableLocationVM = New TableLocationVM
    Set This.TableDirectionVM = New TableDirectionVM
    Set This.TablePreferKeyVM = New TablePreferKeyVM
    Set This.TableStarColumnsVM = New TableStarColumnsVM
    Set This.TableProfileVM = New TableProfileVM
    Set This.TableHighlightingVM = New TableHighlightingVM
    Set This.TableTimestampVM = New TableTimestampVM
    Set This.TableProtectionVM = New TableProtectionVM
End Sub

'---
Public Property Get TableDetailsVM() As TableDetailsVM
    Set TableDetailsVM = This.TableDetailsVM
End Property

Public Property Get TableLocationVM() As TableLocationVM
    Set TableLocationVM = This.TableLocationVM
End Property

Public Property Get TableDirectionVM() As TableDirectionVM
    Set TableDirectionVM = This.TableDirectionVM
End Property

Public Property Get TablePreferKeyVM() As TablePreferKeyVM
    Set TablePreferKeyVM = This.TablePreferKeyVM
End Property

Public Property Get TableStarColumnsVM() As TableStarColumnsVM
    Set TableStarColumnsVM = This.TableStarColumnsVM
End Property

Public Property Get TableProfileVM() As TableProfileVM
    Set TableProfileVM = This.TableProfileVM
End Property

Public Property Get TableHighlightingVM() As TableHighlightingVM
    Set TableHighlightingVM = This.TableHighlightingVM
End Property

Public Property Get TableTimestampVM() As TableTimestampVM
    Set TableTimestampVM = This.TableTimestampVM
End Property

Public Property Get TableProtectionVM() As TableProtectionVM
    Set TableProtectionVM = This.TableProtectionVM
End Property

'---
Public Sub Load(ByVal ListObject As ListObject)
    Dim Worksheet As Worksheet
    Dim Workbook As Workbook
    Set Worksheet = ListObject.Parent
    Set Workbook = Worksheet.Parent

    Dim SettingsModel As ISettingsModel
    Set SettingsModel = XMLSettingsModel.Create(Workbook, SETTINGS_NODE)
    Set This.TableSettings = SettingsModel.Table(ListObject.Name)
    
    This.TableDetailsVM.Load ListObject
    This.TableLocationVM.Load Workbook
    This.TableDirectionVM.Load ListObject, This.TableSettings
    This.TablePreferKeyVM.Load ListObject, This.TableSettings
    This.TableStarColumnsVM.Load ListObject, This.TableSettings

    TableProfileVM.Load ListObject, SettingsModel
    TableHighlightingVM.Load ListObject
    TableTimestampVM.Load ListObject, This.TableSettings
    TableProtectionVM.Load ListObject, This.TableSettings
    
    Commit ' Since we don't have a method to reset .IsChanged, we'll just resave the XML
End Sub

Public Sub Commit()
    CommitChanges This.TableDirectionVM
    CommitChanges This.TablePreferKeyVM
    CommitChanges This.TableStarColumnsVM

    CommitChanges This.TableTimestampVM
    CommitChanges This.TableProtectionVM
End Sub

Private Sub CommitChanges(ByVal ChangeTracker As IChangeTracking)
    ChangeTracker.AcceptChanges
End Sub

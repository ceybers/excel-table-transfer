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
    
    Foobar As String
    
    TableSettings As ISettings
    
    TableDetailsVM As TableDetailsVM
    TableLocationVM As TableLocationVM
    TableDirectionVM As TableDirectionVM
    TablePreferKeyVM As TablePreferKeyVM
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
'---
Public Property Get Foobar() As String
    Foobar = This.Foobar
End Property

Public Property Let Foobar(ByVal vNewValue As String)
    This.Foobar = vNewValue
End Property

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
End Sub

Public Sub Commit()
    CommitChanges This.TableDirectionVM
    CommitChanges This.TablePreferKeyVM
End Sub

Private Sub CommitChanges(ByVal ChangeTracker As IChangeTracking)
    ChangeTracker.AcceptChanges
End Sub

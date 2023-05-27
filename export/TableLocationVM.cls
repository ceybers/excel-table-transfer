VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TableLocationVM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "MVVM.Common.ViewModel"
Option Explicit
Implements IViewModel
Implements INotifyPropertyChanged

Private Type TState
    Notifier As INotifyPropertyChanged
    
    Properties As WorkbookProperty
    
    IsLocalStorage As Boolean
    IsNetworkStorage As Boolean
    IsOneDriveStorage As Boolean
    IsSharePointStorage As Boolean
    
    SelectedLocation As String
    Locations As Scripting.Dictionary
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
Public Property Get IsLocalStorage() As Boolean
    IsLocalStorage = This.IsLocalStorage
End Property

Public Property Get IsNetworkStorage() As Boolean
    IsNetworkStorage = This.IsNetworkStorage
End Property

Public Property Get IsOneDriveStorage() As Boolean
    IsOneDriveStorage = This.IsOneDriveStorage
End Property

Public Property Get IsSharePointStorage() As Boolean
    IsSharePointStorage = This.IsSharePointStorage
End Property

Public Property Get SelectedLocation() As String
    SelectedLocation = This.SelectedLocation
End Property

Public Property Get Locations() As Scripting.Dictionary
    Set Locations = This.Locations
End Property
    
'---
Public Sub Load(ByVal Workbook As Workbook)
    Set This.Properties = New WorkbookProperty
    Set This.Properties.Workbook = Workbook
    
    Set This.Locations = CollectionExConverters.TextToDictionary("Local,Network,OneDrive,SharePoint")
    
    LoadProperties
End Sub

Private Sub LoadProperties()
    With This.Properties
        This.IsLocalStorage = (.StorageLocation = LocalStorage)
        This.IsNetworkStorage = (.StorageLocation = RemoteStorage)
        This.IsOneDriveStorage = (.StorageLocation = OneDriveStorage)
        This.IsSharePointStorage = (.StorageLocation = SharePointStorage)
    End With
    
    Select Case This.Properties.StorageLocation
        Case ttWorkbookLocation.LocalStorage
            This.SelectedLocation = "Local"
        Case ttWorkbookLocation.RemoteStorage
            This.SelectedLocation = "Remote"
        Case ttWorkbookLocation.OneDriveStorage
            This.SelectedLocation = "OneDrive"
        Case ttWorkbookLocation.SharePointStorage
            This.SelectedLocation = "SharePoint"
    End Select
End Sub

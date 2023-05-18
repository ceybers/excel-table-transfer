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

Event PropertyChanged(ByVal PropertyName As String)

Private Type TState
    ListObject As ListObject
    ColumnProperties As ColumnProperties
    KeyColumns As KeyColumns
    WorkbookProperty As WorkbookProperty
    
    ActivateListObjectCommand As ActivateListObjectCommand
    ApplyViewModelCommand As ApplyViewModelCommand
    
    SettingsModel As ISettingsModel
    
    IsDirty As Boolean
End Type
Private This As TState

Public Property Get IsDirty() As Boolean
    IsDirty = This.IsDirty
End Property

Public Property Let IsDirty(ByVal vNewValue As Boolean)
    This.IsDirty = vNewValue
End Property

Public Property Get ApplyViewModelCommand() As ApplyViewModelCommand
    Set ApplyViewModelCommand = This.ApplyViewModelCommand
End Property

Public Property Get ActivateListObjectCommand() As ActivateListObjectCommand
    Set ActivateListObjectCommand = This.ActivateListObjectCommand
End Property

Public Property Get WorkbookProperty() As WorkbookProperty
    Set WorkbookProperty = This.WorkbookProperty
End Property

Public Property Get KeyColumns() As KeyColumns
    Set KeyColumns = This.KeyColumns
End Property

Public Property Get ColumnProperties() As ColumnProperties
    Set ColumnProperties = This.ColumnProperties
End Property

Public Property Get TableName() As String
    TableName = This.ListObject.Name
End Property

Public Property Get WorkSheetName() As String
    WorkSheetName = This.ListObject.Parent.Name
End Property

Public Property Get WorkBookName() As String
    WorkBookName = This.ListObject.Parent.Parent.Name
End Property

Public Sub Load(ByVal ListObject As ListObject)
    Set This.WorkbookProperty = New WorkbookProperty
    Set This.WorkbookProperty.Workbook = ListObject.Parent.Parent
    
    Set This.ListObject = ListObject
    
    Set This.ColumnProperties = New ColumnProperties
    This.ColumnProperties.Load ListObject
    
    Set This.KeyColumns = New KeyColumns
    This.KeyColumns.Load ListObject
    
    LoadCommands
    
    LoadSettings
End Sub

Private Sub LoadCommands()
    Set This.ActivateListObjectCommand = New ActivateListObjectCommand
    This.ActivateListObjectCommand.Create This.ListObject
    
    Set This.ApplyViewModelCommand = New ApplyViewModelCommand
    This.ApplyViewModelCommand.Create Me
End Sub


Private Sub LoadSettings()
    Set This.SettingsModel = XMLSettingsModel.Create(This.ListObject.Parent, "TableProperties")
    
    Dim TableSettings As ISettings
    Set TableSettings = This.SettingsModel.Table(This.ListObject.Name)
    
    Dim SavedKeyColumnName As String
    SavedKeyColumnName = TableSettings.GetSetting("KeyColumn")
    
    Debug.Print "Restored SavedKeyColumnName = '"; SavedKeyColumnName; "'"
    If This.KeyColumns.TrySelectByName(SavedKeyColumnName) Then
        Debug.Print "TrySelect worked"
        OnPropertyChanged "SavedKeyColumnName"
    Else
        Debug.Print "TrySelect failed"
    End If
End Sub

Private Sub SaveSettings()
End Sub

Private Sub OnPropertyChanged(ByVal PropertyName As String)
    Debug.Print "PropertyChanged('"; PropertyName; "')"
    RaiseEvent PropertyChanged(PropertyName)
End Sub

Public Sub DoActiveListObject()
    This.ListObject.Parent.Activate
End Sub

Public Sub Apply()
    If IsDirty Then
        ' Do apply and save settings
        Dim vNewValue As String
        Debug.Print "Saving column name..."
        
        If Not This.KeyColumns.Selected Is Nothing Then
            vNewValue = This.KeyColumns.Selected.Name
        End If
        
        This.SettingsModel.Table(This.ListObject.Name).SetSettings "KeyColumn", vNewValue
        This.IsDirty = False
    End If
End Sub

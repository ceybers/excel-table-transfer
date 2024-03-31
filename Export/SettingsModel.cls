VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SettingsModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "PersistentStorage.Model"
Option Explicit
Implements ISettingsModel

Private Type TState
    UserSettings As ISettings
    WorkbookSettings As ISettings
    TableSettings As Scripting.Dictionary ' of type ISettings
End Type
Private This As TState

Private Function ISettingsModel_User() As ISettings
    Debug.Assert Not Me Is SettingsModel
    Debug.Assert Not This.UserSettings Is Nothing
    Set ISettingsModel_User = This.UserSettings
End Function

Private Function ISettingsModel_Workbook() As ISettings
    Debug.Assert Not Me Is SettingsModel
    Debug.Assert Not This.WorkbookSettings Is Nothing
    Set ISettingsModel_Workbook = This.WorkbookSettings
End Function

Private Function ISettingsModel_Table(ByVal TableName As String) As ISettings
    Debug.Assert Not Me Is SettingsModel
    Debug.Assert Not This.TableSettings Is Nothing
    
    If This.TableSettings.Exists(TableName) Then
        Set ISettingsModel_Table = This.TableSettings.Item(TableName)
        Exit Function
    End If
End Function

Private Sub Class_Initialize()
    Set This.TableSettings = New Scripting.Dictionary
End Sub

'@Description "Creates a new XMLSettingsModel for the given Workbook and RootNode."
Public Function Create() As SettingsModel
Attribute Create.VB_Description = "Creates a new XMLSettingsModel for the given Workbook and RootNode."
    Dim Result As SettingsModel
    Set Result = New SettingsModel

    Set Create = Result
End Function

Public Function AddUserSettings(ByVal Settings As ISettings) As SettingsModel
    Debug.Assert This.UserSettings Is Nothing
    Set This.UserSettings = Settings
    Set AddUserSettings = Me
End Function

Public Function AddWorkbookSettings(ByVal Settings As ISettings) As SettingsModel
    Debug.Assert This.WorkbookSettings Is Nothing
    Set This.WorkbookSettings = Settings
    Set AddWorkbookSettings = Me
    AddTables
End Function

Private Sub AddTables()
    Dim XMLSettings As XMLSettings
    Set XMLSettings = This.WorkbookSettings
    
    Dim Worksheet As Worksheet
    Dim ListObject As ListObject
    
    For Each Worksheet In XMLSettings.Workbook.Worksheets
        For Each ListObject In Worksheet.ListObjects
            AddTable ListObject.Name
        Next ListObject
    Next Worksheet
End Sub

Private Sub AddTable(ByVal TableName As String)
    This.TableSettings.Add _
        Key:=TableName, _
        Item:=XMLSettingsFactory.CreateTableSettings(This.WorkbookSettings, TableName)
End Sub

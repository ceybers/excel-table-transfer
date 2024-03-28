Attribute VB_Name = "TTTUserSettings"
'@Folder("PersistentStorage.MyDocSettings")
Option Explicit

Private Const TABLE_TRANSFER_TOOL_SETTINGS_FILENAME As String = "tabletransfertool.ini"
Private Const TABLE_TRANSFER_TOOL_SETTINGS_UUID As String = "{eb37a119-14e2-47c5-ab73-e46804ff84b5}"

Public Function GetMyDocSettings() As ISettings
    Dim Settings As MyDocSettings
    Set Settings = New MyDocSettings
    
    With Settings
        .Filename = TABLE_TRANSFER_TOOL_SETTINGS_FILENAME
        .UUID = TABLE_TRANSFER_TOOL_SETTINGS_UUID
        .Load
    End With
    
    Set GetMyDocSettings = Settings
End Function


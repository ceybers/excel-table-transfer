VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder "PersistentStorage.Abstract"
Option Explicit

'@Description "Gets/sets a boolean value for the given Flag name."
Public Function GetFlag(ByVal FlagName As String) As Boolean
Attribute GetFlag.VB_Description = "Gets/sets a boolean value for the given Flag name."
End Function

Public Sub SetFlag(ByVal FlagName As String, ByVal Value As Boolean)
End Sub

'@Description "Gets/sets a text value for the given Setting name."
Public Function GetSetting(ByVal SettingName As String) As String
Attribute GetSetting.VB_Description = "Gets/sets a text value for the given Setting name."
End Function

Public Sub SetSetting(ByVal SettingName As String, ByVal Value As String)
End Sub

'@Description "Gets/sets a Collection of text values for the given Collection name."
Public Function GetCollection(ByVal CollectionName As String) As Collection
Attribute GetCollection.VB_Description = "Gets/sets a Collection of text values for the given Collection name."
End Function

Public Sub SetCollection(ByVal CollectionName As String, ByVal Collection As Collection)
End Sub

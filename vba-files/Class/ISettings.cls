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
'@Folder "SomeSettingsModel"
Option Explicit

Public Function GetFlag(ByVal FlagName As String) As Boolean
End Function

Public Sub SetFlag(ByVal FlagName As String, ByVal Value As Boolean)
End Sub

Public Function GetSetting(ByVal SettingName As String) As String
End Function

Public Sub SetSetting(ByVal SettingName As String, ByVal Value As String)
End Sub

Public Function GetCollection(ByVal CollectionName As String) As Collection
End Function

Public Sub SetCollection(ByVal CollectionName As String, ByVal Collection As Collection)
End Sub


VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISettingsModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Interface
'@Folder "PersistentStorage.Abstract"
Option Explicit

'@Description "Gets the Settings object for the user-level settings."
Public Function User() As ISettings
Attribute User.VB_Description = "Gets the Settings object for the user-level settings."
End Function

'@Description "Gets the Settings object for the workbook-level settings."
Public Function Workbook() As ISettings
Attribute Workbook.VB_Description = "Gets the Settings object for the workbook-level settings."
End Function

'@Description "Gets the Settings object for the given table's table-level settings."
Public Function Table(ByVal TableName As String) As ISettings
Attribute Table.VB_Description = "Gets the Settings object for the given table's table-level settings."
End Function

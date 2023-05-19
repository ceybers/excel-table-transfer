Attribute VB_Name = "DebugPrintXML"
'@Folder("PersistentStorage")
Option Explicit

Public Sub DebugPrintXML()
    Dim SettingsModel As ISettingsModel
    Set SettingsModel = XMLSettingsModel.Create(ThisWorkbook, "TableTransferTool")
    'Set This.TableSettings = SettingsModel.Table(ThisWorkbook.Worksheets(1).ListObjects(1).Name)
    
    Dim XSettingsModel As XMLSettingsModel
    Set XSettingsModel = SettingsModel
    XSettingsModel.DebugPrint
End Sub

Attribute VB_Name = "XMLSettingsFactory"
'@Folder("PersistentStorage.XMLSettings")
Option Explicit

Private Const WORKBOOK_NODE As String = "Workbook"
Private Const TABLES_NODE As String = "Tables"
Private Const TABLE_NODE As String = "Table"

Public Function CreateWorkbookSettings(ByVal Workbook As Workbook, ByVal RootNode As String) As ISettings
    Dim Result As XMLSettings
    Set Result = XMLSettings.Create(Workbook, "/" & RootNode & "/" & WORKBOOK_NODE)
    
    With Result
        .Load
    End With
    
    Set CreateWorkbookSettings = Result
End Function

Public Function CreateTableSettings(ByVal WorkbookSettings As XMLSettings, ByVal TableName As String) As ISettings
    Dim TableXPath As String
    TableXPath = WorkbookSettings.XPathPrefix & "/" & TABLES_NODE & "/" & TABLE_NODE & "[@Name='" & TableName & "']"
    
    Dim Result As XMLSettings
    Set Result = XMLSettings.Create(WorkbookSettings.Workbook, TableXPath)
    
    With Result
        .Load
    End With
    
    Set CreateTableSettings = Result
End Function

Attribute VB_Name = "modTestPerTable"
'@Folder("VBAProject")
Option Explicit

Public Sub AAATESTA()
    Dim XMLSettingsPerTable As XMLSettingsPerTable
    Set XMLSettingsPerTable = New XMLSettingsPerTable
    XMLSettingsPerTable.Load ThisWorkbook, "Table1"
    XMLSettingsPerTable.Reset
    
    Dim PreferDirection As String
    PreferDirection = XMLSettingsPerTable.GetTextValue("PreferredDirection")
    Debug.Print "Preferred Direction = "; PreferDirection
    
    Dim KeyColumn As String
    KeyColumn = XMLSettingsPerTable.GetTextValue("KeyColumn")
    Debug.Print "KeyColumn (before) = "; KeyColumn
    
    XMLSettingsPerTable.SetTextValue "KeyColumn", "NewIDColumn"
    KeyColumn = XMLSettingsPerTable.GetTextValue("KeyColumn")
    Debug.Print "KeyColumn (after) = "; KeyColumn
    
    Dim NewNode As String
    NewNode = XMLSettingsPerTable.GetTextValue("NewNode")
    Debug.Print "NewNode (before) = "; NewNode
    
    XMLSettingsPerTable.SetTextValue "NewNode", "I am a new node"
    NewNode = XMLSettingsPerTable.GetTextValue("NewNode")
    Debug.Print "NewNode (after) = "; NewNode
    
    Dim StarCols As Collection
    Set StarCols = XMLSettingsPerTable.GetList("StarredColumns/StarredColumn")
    Debug.Print "Star Cols Count (before) = "; StarCols.Count
    
    Dim NewCol As Collection
    Set NewCol = New Collection
    NewCol.Add "Alpha"
    NewCol.Add "Bravo"
    XMLSettingsPerTable.SetList "StarredColumns/StarredColumn", NewCol
    Set StarCols = XMLSettingsPerTable.GetList("StarredColumns/StarredColumn")
    Debug.Print "Star Cols Count (after) = "; StarCols.Count
End Sub

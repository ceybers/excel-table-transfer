Attribute VB_Name = "ListColumnHelpers"
'@Folder("Helpers.Objects")
Option Explicit

Public Function TryGetListColumn(ByVal ListObject As ListObject, ByVal ListColumnName As String, ByRef OutListColumn As ListColumn) As Boolean
    Debug.Assert Not ListObject Is Nothing
    Debug.Assert ListColumnName <> vbNullString
    
    Dim ListColumn As ListColumn
    For Each ListColumn In ListObject.ListColumns
        If ListColumn.Name = ListColumnName Then
            Set OutListColumn = ListColumn
            TryGetListColumn = True
            Exit Function
        End If
    Next ListColumn
End Function

Public Function Exists(ByVal ListObject As ListObject, ByVal ListColumnName As String) As Boolean
    Dim ListColumn As ListColumn
    Exists = TryGetListColumn(ListObject, ListColumnName, ListColumn)
End Function

Attribute VB_Name = "CollectionHelpers"
'@Folder "Helpers.Objects"
Option Explicit

Public Sub ClearCollection(ByVal coll As Collection)
    Debug.Assert Not coll Is Nothing
    Dim i As Long
    For i = coll.Count To 1 Step -1
        coll.Remove i
    Next i
End Sub


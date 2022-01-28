Attribute VB_Name = "Module2"
'@Folder("HelperFunctions")
Option Explicit

Public Function ClearCollection(ByVal coll As Collection)
    Debug.Assert Not coll Is Nothing
    Dim i As Long
    For i = coll.Count To 1 Step -1
        coll.Remove i
    Next i
End Function

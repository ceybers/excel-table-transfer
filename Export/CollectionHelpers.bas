Attribute VB_Name = "CollectionHelpers"
'@IgnoreModule UseMeaningfulName, ProcedureNotUsed
'@Folder "Helpers.Objects"
Option Explicit

'@Description "Returns True if the given Value exists in a Collection."
Public Function ExistsInCollection(ByVal Collection As Object, ByVal Value As Variant) As Boolean
Attribute ExistsInCollection.VB_Description = "Returns True if the given Value exists in a Collection."
    Debug.Assert Not Collection Is Nothing
    
    Dim ThisValue As Variant
    For Each ThisValue In Collection
        'If ThisValue = Value Then
        If CStr(ThisValue) = CStr(Value) Then
        'If StrComp(ThisValue, Value) Then ' Run-time error '458' Variable uses an Automation Type supported in Visual Basic
            ExistsInCollection = True
            Exit Function
        End If
    Next ThisValue
End Function

'@Description "Removes all items in a Collection."
Public Sub CollectionClear(ByVal Collection As Collection)
Attribute CollectionClear.VB_Description = "Removes all items in a Collection."
    Debug.Assert Not Collection Is Nothing
    
    Dim i As Long
    For i = Collection.Count To 1 Step -1
        Collection.Remove i
    Next i
End Sub

'@Description "Copies all items in collection LHS to RHS. Does not copy keys."
Public Sub Clone(ByVal LHS As Collection, ByVal RHS As Collection)
Attribute Clone.VB_Description = "Copies all items in collection LHS to RHS. Does not copy keys."
    Dim Item As Variant
    For Each Item In LHS
        RHS.Add Item
    Next Item
End Sub

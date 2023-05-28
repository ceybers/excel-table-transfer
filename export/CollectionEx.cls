VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Helpers.CollectionEx"
'@PredeclaredId
Option Explicit
Implements ICollectionEx

Private Const COLLECTION_LBOUND As Long = 1
Private Const RESULT_IF_NOT_FOUND As Long = -1

Private Type TState
    Collection As Collection
End Type
Private This As TState

Public Function From(ByVal Collection As Collection) As ICollectionEx
    If Collection Is Nothing Then Exit Function
    If Not TypeOf Collection Is Collection Then Exit Function
    
    Dim Result As New CollectionEx
    Set Result = New CollectionEx
    Result.SetCollection Collection
    Set From = Result
End Function

Public Sub SetCollection(ByVal Collection As Collection)
    Set This.Collection = Collection
End Sub

Private Function ICollectionEx_Contains(ByVal Value As Variant) As Boolean
    Dim ThisVariant As Variant
    For Each ThisVariant In This.Collection
        If ThisVariant = Value Then
            ICollectionEx_Contains = True
            Exit Function
        End If
    Next ThisVariant
End Function

Private Function ICollectionEx_ContainsByProperty(ByVal PropertyName As String, ByVal Value As Variant) As Boolean
    Dim ThisVariant As Variant
    Dim ThisValue As Variant
    For Each ThisVariant In This.Collection
        On Error GoTo ErrHandler
        ThisValue = CallByName(ThisVariant, PropertyName, VbGet)
        On Error GoTo 0
        If ThisValue = Value Then
            ICollectionEx_ContainsByProperty = True
            Exit Function
        End If
    Next ThisVariant
    Exit Function
ErrHandler:
    If Err.Number = 438 Then
    Debug.Print "Error"; Err.Number; ": "; Err.Description; ". No property named '"; PropertyName; "' in variant."
    End If
    Err.Clear
End Function

Private Function ICollectionEx_IndexOf(ByVal Value As Variant) As Long
    ICollectionEx_IndexOf = RESULT_IF_NOT_FOUND
    Dim i As Long
    For i = COLLECTION_LBOUND To This.Collection.Count
        If This.Collection(i) = Value Then
            ICollectionEx_IndexOf = i
            Exit Function
        End If
    Next i
End Function

Private Sub ICollectionEx_Remove(ByVal Value As Variant)
    ICollectionEx_TryRemove Value
End Sub

Private Function ICollectionEx_TryRemove(ByVal Value As Variant) As Boolean
    Dim Index As Long
    Index = ICollectionEx_IndexOf(Value)
    If Index = RESULT_IF_NOT_FOUND Then
        ICollectionEx_TryRemove = False
    Else
        This.Collection.Remove Index
        ICollectionEx_TryRemove = True
    End If
End Function

Private Sub ICollectionEx_RemoveAt(ByVal Index As Long)
    ICollectionEx_TryRemoveAt Index
End Sub

Private Function ICollectionEx_TryRemoveAt(ByVal Index As Long) As Boolean
    If Not IsIndexInBounds(Index) Then Exit Function
    This.Collection.Remove Index
    ICollectionEx_TryRemoveAt = True
End Function

Private Function ICollectionEx_Count() As Long
    ICollectionEx_Count = This.Collection.Count
End Function

Private Sub ICollectionEx_Clear()
    Dim i As Long
    For i = This.Collection.Count To COLLECTION_LBOUND Step -1
        This.Collection.Remove i
    Next i
End Sub

Private Sub ICollectionEx_Add(ByVal Key As Variant, ByVal Value As Variant)
    ICollectionEx_TryAdd Key, Value
End Sub

Private Function ICollectionEx_TryAdd(ByVal Key As Variant, ByVal Value As Variant) As Boolean
    Dim vv As Variant
    On Error Resume Next
    vv = This.Collection.Item(Key)
    On Error GoTo 0
    
    If IsEmpty(vv) = True Then
        This.Collection.Add Item:=Value, Key:=Key
        ICollectionEx_TryAdd = True
    Else
        ICollectionEx_TryAdd = False
    End If
End Function

Private Sub ICollectionEx_Insert(ByVal Index As Long, ByVal Value As Variant)
    ICollectionEx_TryInsert Index, Value
End Sub

Private Function ICollectionEx_TryInsert(ByVal Index As Long, ByVal Value As Variant) As Boolean
    If Not IsIndexInBounds(Index) Then Exit Function
    This.Collection.Add Item:=Value, Before:=Index
    ICollectionEx_TryInsert = True
End Function

Private Function ICollectionEx_GetByIndex(ByVal Index As Long) As Variant
    Dim Result As Variant
    If ICollectionEx_TryGetByIndex(Index, Result) Then
        ICollectionEx_GetByIndex = Result
    End If
End Function

Private Function ICollectionEx_TryGetByIndex(ByVal Index As Long, ByRef OutValue As Variant) As Boolean
    If Not IsIndexInBounds(Index) Then Exit Function
    OutValue = This.Collection.Item(Index)
    ICollectionEx_TryGetByIndex = True
End Function

Private Function ICollectionEx_GetByKey(ByVal Key As Variant) As Variant
    Dim Result As Variant
    If ICollectionEx_TryGetByKey(Key, Result) Then
        ICollectionEx_GetByKey = Result
    End If
End Function

Private Function ICollectionEx_TryGetByKey(ByVal Key As Variant, ByRef OutValue As Variant) As Boolean
    Dim vv As Variant
    On Error Resume Next
    vv = This.Collection.Item(Key)
    On Error GoTo 0
    
    If IsEmpty(vv) = False Then
        OutValue = vv
        ICollectionEx_TryGetByKey = True
    Else
        ICollectionEx_TryGetByKey = False
    End If
End Function

Public Sub DebugPrint()
    Dim i As Long
    Debug.Print "CollectionEx.DebugPrint()"
    For i = 1 To This.Collection.Count
        Debug.Print " "; i; "- "; This.Collection.Item(i)
    Next i
    Debug.Print "---"
End Sub

Private Sub ICollectionEx_ForEach(ByVal Object As Object, ByVal MethodName As Variant)
    Dim ThisVariant As Variant
    Dim ThisValue As Variant
    For Each ThisVariant In This.Collection
        If TypeOf ThisVariant Is Object  Then
            On Error GoTo ErrHandler
            CallByName Object, MethodName, VbMethod, ThisVariant
            On Error GoTo 0
        End If
    Next ThisVariant
    Exit Sub
ErrHandler:
    If Err.Number = 438 Then
    Debug.Print "Error"; Err.Number; ": "; Err.Description; ". No method named '"; MethodName; "' in object."
    End If
    Err.Clear
End Sub

Private Function IsIndexInBounds(ByVal Index As Long) As Boolean
    If Index < COLLECTION_LBOUND Then Exit Function
    If Index > This.Collection.Count Then Exit Function
    IsIndexInBounds = True
End Function

Private Function ICollectionEx_ToArray() As Variant
    Dim Result As Variant
    Result = Array()
    ReDim Result(0 To (This.Collection.Count - 1))
    
    If UBound(Result) > 0 Then
        Dim i As Long
        For i = 0 To UBound(Result)
            Result(i) = This.Collection.Item(i + 1)
        Next i
    End If
    
    ICollectionEx_ToArray = Result
End Function

Private Function ICollectionEx_ToCollection() As Collection
    Set ICollectionEx_ToCollection = This.Collection
End Function

Private Function ICollectionEx_ToDictionary() As Scripting.Dictionary
    Dim Result As Scripting.Dictionary
    Set Result = New Scripting.Dictionary
    
    Dim i As Long
    For i = 1 To This.Collection.Count
        Result.Add Key:="I" & i, Item:=This.Collection.Item(i)
    Next i
    
    Set ICollectionEx_ToDictionary = Result
End Function

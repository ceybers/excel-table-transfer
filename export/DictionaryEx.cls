VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DictionaryEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "Helpers.CollectionEx"
Option Explicit
Implements ICollectionEx

Private Const COLLECTION_LBOUND As Long = 0
Private Const RESULT_IF_NOT_FOUND As Long = -1

Private Type TState
    Dictionary As Scripting.Dictionary
End Type
Private This As TState

Public Function From(ByVal Dictionary As Scripting.Dictionary) As ICollectionEx
    If Dictionary Is Nothing Then Exit Function
    If Not TypeOf Dictionary Is Scripting.Dictionary Then Exit Function
    
    Dim Result As New DictionaryEx
    Set Result = New DictionaryEx
    Result.SetDictionary Dictionary
    Set From = Result
End Function

Public Sub SetDictionary(ByVal Dictionary As Scripting.Dictionary)
    Set This.Dictionary = Dictionary
End Sub

Private Function ICollectionEx_Contains(ByVal Value As Variant) As Boolean
    Dim ThisVariant As Variant
    For Each ThisVariant In This.Dictionary.Items
        If ThisVariant = Value Then
            ICollectionEx_Contains = True
            Exit Function
        End If
    Next ThisVariant
End Function

Private Function ICollectionEx_ContainsByProperty(ByVal PropertyName As String, ByVal Value As Variant) As Boolean
    Dim ThisVariant As Variant
    Dim ThisValue As Variant
    For Each ThisVariant In This.Dictionary.Items
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
    For i = COLLECTION_LBOUND To (This.Dictionary.Count - 1)
        If This.Dictionary.Items(i) = Value Then
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
        'This.Collection.Remove Index
        'This.Dictionary.Remove(key) ??? NYI
        ICollectionEx_TryRemove = False 'True
        Debug.Assert False
    End If
End Function

Private Sub ICollectionEx_RemoveAt(ByVal Index As Long)
    ICollectionEx_TryRemoveAt Index
End Sub

Private Function ICollectionEx_TryRemoveAt(ByVal Index As Long) As Boolean
    'If Not IsIndexInBounds(Index) Then Exit Function
    'This.Collection.Remove Index
    ' dict remove uses key not index
    'ICollectionEx_TryRemoveAt = True
    Debug.Assert False
End Function

Private Function ICollectionEx_Count() As Long
    ICollectionEx_Count = This.Dictionary.Count
End Function

Private Sub ICollectionEx_Clear()
    This.Dictionary.RemoveAll
End Sub

Private Sub ICollectionEx_Add(ByVal Key As Variant, ByVal Value As Variant)
    ICollectionEx_TryAdd Key, Value
End Sub

Private Function ICollectionEx_TryAdd(ByVal Key As Variant, ByVal Value As Variant) As Boolean
    If This.Dictionary.Exists(Key) = False Then
        This.Dictionary.Add Item:=Value, Key:=Key
        ICollectionEx_TryAdd = True
    End If
End Function

Private Sub ICollectionEx_Insert(ByVal Index As Long, ByVal Value As Variant)
    Debug.Assert False
    Err.Raise 42, Description:="NYI, Cannot be implemented?"
End Sub

Private Function ICollectionEx_TryInsert(ByVal Index As Long, ByVal Value As Variant) As Boolean
    Debug.Assert False
    Err.Raise 42, Description:="NYI, Cannot be implemented?"
End Function

Private Function ICollectionEx_GetByIndex(ByVal Index As Long) As Variant
    Dim Result As Variant
    If ICollectionEx_TryGetByIndex(Index, Result) Then
        ICollectionEx_GetByIndex = Result
    End If
End Function

Private Function ICollectionEx_TryGetByIndex(ByVal Index As Long, ByRef OutValue As Variant) As Boolean
    If Not IsIndexInBounds(Index) Then Exit Function
    OutValue = This.Dictionary.Items(Index)
    ICollectionEx_TryGetByIndex = True
End Function

Private Function ICollectionEx_GetByKey(ByVal Key As Variant) As Variant
    Dim Result As Variant
    If ICollectionEx_TryGetByKey(Key, Result) Then
        ICollectionEx_GetByKey = Result
    End If
End Function

Private Function ICollectionEx_TryGetByKey(ByVal Key As Variant, ByRef OutValue As Variant) As Boolean
    If This.Dictionary.Exists(Key) Then
        Set OutValue = This.Dictionary.Item(Key)
        ICollectionEx_TryGetByKey = True
    End If
End Function

Public Sub DebugPrint()
    'Dim i As Long
    'Debug.Print "CollectionEx.DebugPrint()"
    'For i = 1 To This.Collection.Count
    '    Debug.Print " "; i; "- "; This.Collection.Item(i)
    'Next i
    'Debug.Print "---"
End Sub

Private Sub ICollectionEx_ForEach(ByVal Object As Object, ByVal MethodName As Variant)
    Dim ThisVariant As Variant
    Dim ThisValue As Variant
    For Each ThisVariant In This.Dictionary.Items
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
    If Index > (This.Dictionary.Count - 1) Then Exit Function
    IsIndexInBounds = True
End Function

Private Function ICollectionEx_ToArray() As Variant
    ICollectionEx_ToArray = This.Dictionary.Items
End Function

Private Function ICollectionEx_ToCollection() As Collection
    Dim Result As Collection
    Set Result = New Collection
    
    Dim i As Long
    For i = 0 To (This.Dictionary.Count - 1)
        Result.Add This.Dictionary.Keys(i), This.Dictionary.Items(i)
    Next i
    Set ICollectionEx_ToCollection = Result
End Function

Private Function ICollectionEx_ToDictionary() As Scripting.Dictionary
    Set ICollectionEx_ToDictionary = This.Dictionary
End Function

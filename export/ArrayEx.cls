VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "Helpers.CollectionEx"
Option Explicit
Implements ICollectionEx

'Private Const COLLECTION_LBOUND As Long = 0
Private Const RESULT_IF_NOT_FOUND As Long = -1

Private Type TState
    Array As Variant
End Type
Private This As TState

Public Function From(ByVal ArrayToLoad As Variant) As ICollectionEx
    If Not IsArray(ArrayToLoad) Then Exit Function
    ' Need to test that this is a single rank array
    
    Dim Result As New ArrayEx
    Set Result = New ArrayEx
    Result.SetArray ArrayToLoad
    Set From = Result
End Function

Public Sub SetArray(ByVal ArrayToLoad As Variant)
    This.Array = ArrayToLoad
End Sub

Private Function ICollectionEx_Contains(ByVal Value As Variant) As Boolean
    Dim ThisVariant As Variant
    For Each ThisVariant In This.Array
        If ThisVariant = Value Then
            ICollectionEx_Contains = True
            Exit Function
        End If
    Next ThisVariant
End Function

Private Function ICollectionEx_ContainsByProperty(ByVal PropertyName As String, ByVal Value As Variant) As Boolean
    Dim ThisVariant As Variant
    Dim ThisValue As Variant
    For Each ThisVariant In This.Array
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
    For i = LBound(This.Array) To UBound(This.Array)
        If This.Array(i) = Value Then
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
    
    If Index = RESULT_IF_NOT_FOUND Then Exit Function
    
    Dim i As Long
    For i = UBound(This.Array) - 1 To (Index) Step -1
        This.Array(i) = This.Array(i + 1)
    Next i
    ReDim Preserve This.Array(LBound(This.Array) To UBound(This.Array) - 1)
    ICollectionEx_TryRemove = True
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
    ICollectionEx_Count = UBound(This.Array) - LBound(This.Array) + 1
End Function

Private Sub ICollectionEx_Clear()
    This.Array = Array()
End Sub

Private Sub ICollectionEx_Add(ByVal Key As Variant, ByVal Value As Variant)
    ICollectionEx_TryAdd Key, Value
End Sub

Private Function ICollectionEx_TryAdd(ByVal Key As Variant, ByVal Value As Variant) As Boolean
    ReDim Preserve This.Array(LBound(This.Array) To UBound(This.Array) + 1)
    This.Array(UBound(This.Array)) = Value
    ICollectionEx_TryAdd = True
End Function

Private Sub ICollectionEx_Insert(ByVal Index As Long, ByVal Value As Variant)
    Debug.Assert False
    Err.Raise 42, Description:="NYI TODO"
End Sub

Private Function ICollectionEx_TryInsert(ByVal Index As Long, ByVal Value As Variant) As Boolean
    Debug.Assert False
    Err.Raise 42, Description:="NYI TODO"
    ' redim, start from back, shift everything down once
End Function

Private Function ICollectionEx_GetByIndex(ByVal Index As Long) As Variant
    Dim Result As Variant
    If ICollectionEx_TryGetByIndex(Index, Result) Then
        ICollectionEx_GetByIndex = Result
    End If
End Function

Private Function ICollectionEx_TryGetByIndex(ByVal Index As Long, ByRef outValue As Variant) As Boolean
    If Not IsIndexInBounds(Index) Then Exit Function
    outValue = This.Array(Index)
    ICollectionEx_TryGetByIndex = True
End Function

Private Function ICollectionEx_GetByKey(ByVal Key As Variant) As Variant
    Debug.Assert False
    Err.Raise 42, Description:="NYI No keys in arrays"
End Function

Private Function ICollectionEx_TryGetByKey(ByVal Key As Variant, ByRef outValue As Variant) As Boolean
    Debug.Assert False
    Err.Raise 42, Description:="NYI No keys in arrays"
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
    For Each ThisVariant In This.Array
        If TypeOf ThisVariant Is Object  Then
            On Error GoTo ErrHandler
            Call CallByName(Object, MethodName, VbMethod, ThisVariant)
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
    If Index < LBound(This.Array) Then Exit Function
    If Index > UBound(This.Array) Then Exit Function
    IsIndexInBounds = True
End Function

Private Function ICollectionEx_ToArray() As Variant
    ICollectionEx_ToArray = This.Array
End Function

Private Function ICollectionEx_ToCollection() As Collection
    Dim Result As Collection
    Set Result = New Collection
    
    Dim i As Long
    For i = LBound(This.Array) To UBound(This.Array)
        Result.Add Key:="I" & i, Item:=This.Array(i)
    Next i
    Set ICollectionEx_ToCollection = Result
End Function

Private Function ICollectionEx_ToDictionary() As Scripting.Dictionary
    Dim Result As Scripting.Dictionary
    Set Result = New Scripting.Dictionary
    
    Dim i As Long
    For i = LBound(This.Array) To UBound(This.Array)
        Result.Add Key:="I" & i, Item:=This.Array(i)
    Next i
    Set ICollectionEx_ToDictionary = Result
End Function

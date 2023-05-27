VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder "Helpers.Collection"
Option Explicit
Implements ICollectionEx

Private Type TState
    Collection As Collection
End Type
Private This As TState

Public Function From(ByVal Collection As Collection) As CollectionEx
    If Collection Is Nothing Then Exit Function
    If Not TypeOf Collection Is Collection Then Exit Function
    
    Set From = New CollectionEx
    Dim Interface As ICollectionEx
    Set Interface = From
    Interface.SetCollection Collection
End Function

Private Sub ICollectionEx_SetCollection(ByVal Collection As Collection)
    Set This.Collection = Collection
End Sub

Public Function Exists(ByVal Value As Variant) As Boolean
    Dim ThisVariant As Variant
    For Each ThisVariant In This.Collection
        If ThisVariant = Value Then
            Exists = True
            Exit Function
        End If
    Next ThisVariant
End Function

Public Function ExistsByProperty(ByVal PropertyName As String, ByVal Value As Variant) As Boolean
    Dim ThisVariant As Variant
    Dim ThisValue As Variant
    For Each ThisVariant In This.Collection
        On Error GoTo ErrHandler
        ThisValue = CallByName(ThisVariant, PropertyName, VbGet)
        On Error GoTo 0
        If ThisValue = Value Then
            ExistsByProperty = True
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

Public Function IndexOf(ByVal Value As Variant) As Long
    IndexOf = -1
    Dim i As Long
    For i = 1 To This.Collection.Count
        If This.Collection(i) = Value Then
            IndexOf = i
            Exit Function
        End If
    Next i
End Function

Public Function TryRemove(ByVal Value As Variant) As Boolean
    Dim Index As Long
    Index = IndexOf(Value)
    If Index = -1 Then
        TryRemove = False
    Else
        This.Collection.Remove Index
        TryRemove = True
    End If
End Function

Public Sub Clear()
    Dim i As Long
    For i = This.Collection.Count To 1 Step -1
        This.Collection.Remove i
    Next i
End Sub

Public Function TryAdd(ByVal Key As Variant, ByVal Value As Variant) As Boolean
    Dim vv As Variant
    On Error Resume Next
    vv = This.Collection.Item(Key)
    On Error GoTo 0
    
    If IsEmpty(vv) = True Then
        This.Collection.Add Item:=Value, Key:=Key
        TryAdd = True
    Else
        TryAdd = False
    End If
End Function

Public Function TryGet(ByVal Key As Variant, ByRef outValue As Variant) As Boolean
    Dim vv As Variant
    On Error Resume Next
    vv = This.Collection.Item(Key)
    On Error GoTo 0
    
    If IsEmpty(vv) = False Then
        outValue = vv
        TryGet = True
    Else
        TryGet = False
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

Public Function ForEach(ByVal Object As Object, ByVal MethodName As Variant) As Boolean
    Dim ThisVariant As Variant
    Dim ThisValue As Variant
    For Each ThisVariant In This.Collection
        If TypeOf ThisVariant Is Object  Then
            On Error GoTo ErrHandler
            Call CallByName(Object, MethodName, VbMethod, ThisVariant)
            On Error GoTo 0
        End If
    Next ThisVariant
    Exit Function
ErrHandler:
    If Err.Number = 438 Then
    Debug.Print "Error"; Err.Number; ": "; Err.Description; ". No method named '"; MethodName; "' in object."
    End If
    Err.Clear
End Function

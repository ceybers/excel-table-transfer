VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MostRecentlyUsed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Data structure for storing a Most Recently Used list."
'@ModuleDescription "Data structure for storing a Most Recently Used list."
'@PredeclaredId
'@Folder("Helpers.CollectionEx")
Option Explicit

Implements IMostRecentlyUsed

Private Const INITIAL_MRU_LENGTH As Long = 4
Private Const MAXIMUM_MRU_LENGTH As Long = 255

Private Type TState
    Arr As Variant
    Tail As Long
    MaxLength As Long
End Type
Private This As TState

Private Sub Class_Initialize()
    This.Arr = Array()
    ReDim This.Arr(0 To (INITIAL_MRU_LENGTH - 1))
    This.Tail = -1
    This.MaxLength = INITIAL_MRU_LENGTH
End Sub

Private Sub IMostRecentlyUsed_SetMaximumLength(ByVal Length As Long)
    Debug.Assert Not Me Is MostRecentlyUsed
    Debug.Assert Length > 0
    Debug.Assert Length <= MAXIMUM_MRU_LENGTH
    This.MaxLength = Length
    ReDim Preserve This.Arr(0 To (This.MaxLength - 1))
    If This.Tail = -1 > Length Then
        Debug.Assert False
    End If
End Sub

Private Sub IMostRecentlyUsed_Add(ByVal Value As Variant)
    Debug.Assert Not Me Is MostRecentlyUsed
    Dim Index As Long
    Index = IndexOf(Value)
    
    If Index > -1 Then
        AddExisting Value, Index
    Else
        AddNew Value
    End If
End Sub

Private Sub AddNew(ByVal Value As Variant)
    This.Tail = This.Tail + 1
    If This.Tail > UBound(This.Arr) Then
        This.Tail = UBound(This.Arr)
    End If
    
    Dim i As Long
    For i = This.Tail To 1 Step -1
        This.Arr(i) = This.Arr(i - 1)
    Next i
    
    This.Arr(0) = Value
End Sub

Private Sub AddExisting(ByVal Value As Variant, ByVal Index As Long)
    Dim i As Long
    For i = Index To 1 Step -1
        This.Arr(i) = This.Arr(i - 1)
    Next i
    This.Arr(0) = Value
End Sub

Private Sub IMostRecentlyUsed_Remove(ByVal Value As Variant)
    Debug.Assert Not Me Is MostRecentlyUsed
    Dim Index As Long
    Index = IndexOf(Value)
    If Index = -1 Then Exit Sub
    IMostRecentlyUsed_RemoveAt Index
End Sub

Private Sub IMostRecentlyUsed_RemoveAt(ByVal Index As Variant)
    Debug.Assert Not Me Is MostRecentlyUsed
    If (Index < 0) Or (Index > This.Tail) Then
        Err.Raise vbObjectError + 5, "MostRecentlyUsed", "Index out of bounds"
    End If
    
    Dim i As Long
    For i = Index To (This.Tail - 1)
        This.Arr(i) = This.Arr(i + 1)
    Next i
    
    This.Arr(This.Tail) = Empty
    This.Tail = This.Tail - 1
End Sub

Private Function IMostRecentlyUsed_Item(ByVal Index As Long) As Variant
    Debug.Assert Not Me Is MostRecentlyUsed
    If (Index < 0) Or (Index > This.Tail) Then
        Err.Raise vbObjectError + 5, "MostRecentlyUsed", "Index out of bounds"
    End If
    IMostRecentlyUsed_Item = This.Arr(Index)
End Function

Private Function IMostRecentlyUsed_Count() As Long
    Debug.Assert Not Me Is MostRecentlyUsed
    IMostRecentlyUsed_Count = (This.Tail + 1)
End Function

Private Sub IMostRecentlyUsed_Clear()
    Debug.Assert Not Me Is MostRecentlyUsed
    Dim i As Long
    For i = UBound(This.Arr) To 0 Step -1
        This.Arr(i) = Empty
    Next i
    This.Tail = -1
End Sub

Private Function IndexOf(ByVal Value As Variant) As Long
    Dim i As Long
    For i = 0 To This.Tail 'UBound(This.Arr)
        If This.Arr(i) = Value Then
            IndexOf = i
            Exit Function
        End If
    Next i
    IndexOf = -1
End Function

Private Function IMostRecentlyUsed_ToCollection() As Collection
    Dim Result As Collection
    Set Result = New Collection
    
    Dim i As Long
    For i = 0 To This.Tail
        Result.Add Item:=This.Arr(i), Key:=CStr(i)
    Next i
    
    Set IMostRecentlyUsed_ToCollection = Result
End Function

Private Sub IMostRecentlyUsed_FromCollection(ByVal Collection As Collection)
    Dim i As Long
    For i = Collection.Count To 1 Step -1
        IMostRecentlyUsed_Add Collection.Item(i)
    Next i
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICollectionEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Helpers.CollectionEx.Abstract"
Option Explicit

Public Function Contains(ByVal Value As Variant) As Boolean
End Function

Public Function ContainsByProperty(ByVal PropertyName As String, ByVal Value As Variant) As Boolean
End Function

Public Function IndexOf(ByVal Value As Variant) As Long
End Function

Public Sub Add(ByVal Key As Variant, ByVal Value As Variant)
End Sub

Public Function TryAdd(ByVal Key As Variant, ByVal Value As Variant) As Boolean
End Function

Public Sub Insert(ByVal Index As Long, ByVal Value As Variant)
End Sub

Public Function TryInsert(ByVal Index As Long, ByVal Value As Variant) As Boolean
End Function

Public Function GetByIndex(ByVal Index As Long) As Variant
End Function

Public Function TryGetByIndex(ByVal Index As Long, ByRef outValue As Variant) As Boolean
End Function

Public Function GetByKey(ByVal Key As Variant) As Variant
End Function

Public Function TryGetByKey(ByVal Key As Variant, ByRef outValue As Variant) As Boolean
End Function

Public Sub Remove(ByVal Value As Variant)
End Sub

Public Function TryRemove(ByVal Value As Variant) As Boolean
End Function

Public Sub RemoveAt(ByVal Index As Long)
End Sub

Public Function TryRemoveAt(ByVal Index As Long) As Boolean
End Function

Public Sub ForEach(ByVal Object As Object, ByVal MethodName As Variant)
End Sub

Public Function Count() As Long
End Function

Public Sub Clear()
End Sub

Public Function ToArray() As Variant
End Function

Public Function ToCollection() As Collection
End Function

Public Function ToDictionary() As Scripting.Dictionary
End Function

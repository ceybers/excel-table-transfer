VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMostRecentlyUsed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@Folder("Helpers.CollectionEx")
'@Interface

'@Description "Sets the maximum length of items in the structure. Reducing the size trunctuates existing elements."
Public Sub SetMaximumLength(ByVal Length As Long)
Attribute SetMaximumLength.VB_Description = "Sets the maximum length of items in the structure. Reducing the size trunctuates existing elements."
End Sub

'@Description "Adds an element to the top of the list. If it already exists in the list, it is moved to the top."
Public Sub Add(ByVal Value As Variant)
Attribute Add.VB_Description = "Adds an element to the top of the list. If it already exists in the list, it is moved to the top."
End Sub

'@Description "Removes a given element from the list if it exists."
Public Sub Remove(ByVal Value As Variant)
Attribute Remove.VB_Description = "Removes a given element from the list if it exists."
End Sub

'@Description "Removes the element at the given index."
Public Sub RemoveAt(ByVal Index As Variant)
Attribute RemoveAt.VB_Description = "Removes the element at the given index."
End Sub

'@Description "Gets the element at the given index."
Public Function Item(ByVal Index As Long) As Variant
Attribute Item.VB_Description = "Gets the element at the given index."
End Function

'@Description "Returns the count of elements in the list."
Public Function Count() As Long
Attribute Count.VB_Description = "Returns the count of elements in the list."
End Function

'@Description "Removes all elements in the list."
Public Sub Clear()
Attribute Clear.VB_Description = "Removes all elements in the list."
End Sub

'@Description "Converts the MostRecentlyUsed object to a Collection object."
Public Function ToCollection() As Collection
Attribute ToCollection.VB_Description = "Converts the MostRecentlyUsed object to a Collection object."
End Function

'@Description "Appends the contents of a Collection object to this MostRecentlyUsed object. Item 1 in the Collection will be the most recent element in the MRU."
Public Sub FromCollection(ByVal Collection As Collection)
Attribute FromCollection.VB_Description = "Appends the contents of a Collection object to this MostRecentlyUsed object. Item 1 in the Collection will be the most recent element in the MRU."
End Sub
